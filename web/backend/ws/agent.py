"""Agent WebSocket 端点 — 仪器 Agent 通过此端点连接服务器"""

import json

from fastapi import APIRouter, WebSocket, WebSocketDisconnect

from ..services.agent_manager import agent_manager
from ..services.ws_manager import ws_manager

router = APIRouter()


@router.websocket("/agent/ws")
async def agent_websocket(ws: WebSocket):
    """仪器 Agent 专用 WebSocket"""
    await ws.accept()
    agent_id = None

    try:
        while True:
            raw = await ws.receive_text()
            try:
                msg = json.loads(raw)
            except json.JSONDecodeError:
                continue

            msg_type = msg.get("type", "")

            if msg_type == "register":
                agent_id = msg.get("agent_id", "")
                hostname = msg.get("hostname", "")
                modules = msg.get("modules", [])
                await agent_manager.register(agent_id, ws, hostname, modules)
                print(f"[AgentWS] {agent_id} 注册成功")

            elif msg_type == "pong":
                pass  # 心跳应答，忽略

            elif agent_id:
                # 转发 Agent 消息到前端客户端（log / progress / completed / error / data）
                module = msg.get("module", "")
                test_id = msg.get("test_id", "")

                await ws_manager.broadcast({
                    "type": msg_type,
                    "module": module,
                    "agent_id": agent_id,
                    "test_id": test_id,
                    "payload": {
                        "message": msg.get("message", ""),
                        "status": msg.get("status"),
                        "result": msg.get("result"),
                        "phase": msg.get("phase"),
                        "segment": msg.get("segment"),
                        "total_segments": msg.get("total_segments"),
                        "progress": msg.get("progress"),
                        "error": msg.get("error"),
                    },
                })

    except WebSocketDisconnect:
        pass
    except Exception as e:
        print(f"[AgentWS] 异常: {e}")
    finally:
        if agent_id:
            await agent_manager.unregister(agent_id)
