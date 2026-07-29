"""WebSocket 端点 — 测试日志/进度实时推送"""

from fastapi import APIRouter, WebSocket, WebSocketDisconnect

from ..services.ws_manager import ws_manager

router = APIRouter()


@router.websocket("/ws")
async def websocket_global(ws: WebSocket):
    """全局 WebSocket — 接收所有广播消息"""
    await ws_manager.connect(ws)
    try:
        while True:
            # 保持连接，接收客户端消息（如心跳）
            _ = await ws.receive_text()
    except WebSocketDisconnect:
        pass
    finally:
        await ws_manager.disconnect(ws)


@router.websocket("/ws/{module_name}")
async def websocket_module(ws: WebSocket, module_name: str):
    """模块级 WebSocket — 订阅特定模块的消息"""
    await ws_manager.connect(ws, module_name=module_name)
    try:
        while True:
            _ = await ws.receive_text()
    except WebSocketDisconnect:
        pass
    finally:
        await ws_manager.disconnect(ws, module_name=module_name)
