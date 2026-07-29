"""Agent 连接管理器 — 维护仪器 Agent 连接池，路由命令，转发消息"""

import asyncio
import json
from typing import Callable

from fastapi import WebSocket

from .ws_manager import ws_manager


class AgentInfo:
    def __init__(self, agent_id: str, ws: WebSocket, hostname: str = "", modules: list[str] | None = None):
        self.agent_id = agent_id
        self.ws = ws
        self.hostname = hostname
        self.modules = modules or []
        self.online = True


class AgentManager:
    """管理所有连接的仪器 Agent"""

    def __init__(self):
        self._agents: dict[str, AgentInfo] = {}  # agent_id → AgentInfo
        self._lock = asyncio.Lock()

    # -----------------------------------------------------------------
    # 注册 / 注销
    # -----------------------------------------------------------------
    async def register(self, agent_id: str, ws: WebSocket, hostname: str = "", modules: list[str] | None = None):
        """Agent 注册上线"""
        async with self._lock:
            info = AgentInfo(agent_id=agent_id, ws=ws, hostname=hostname, modules=modules or [])
            self._agents[agent_id] = info
        await ws_manager.broadcast_global({
            "type": "agent_online",
            "payload": {"agent_id": agent_id, "hostname": hostname, "modules": modules},
        })
        print(f"[AgentManager] Agent 上线: {agent_id} ({hostname}) 模块: {modules}")

    async def unregister(self, agent_id: str):
        """Agent 断线"""
        async with self._lock:
            self._agents.pop(agent_id, None)
        await ws_manager.broadcast_global({
            "type": "agent_offline",
            "payload": {"agent_id": agent_id},
        })
        print(f"[AgentManager] Agent 下线: {agent_id}")

    # -----------------------------------------------------------------
    # 查询
    # -----------------------------------------------------------------
    async def get_online_agents(self) -> list[dict]:
        """获取在线 Agent 列表"""
        async with self._lock:
            return [
                {
                    "agent_id": a.agent_id,
                    "hostname": a.hostname,
                    "modules": a.modules,
                    "online": a.online,
                }
                for a in self._agents.values()
            ]

    async def get_agent_for_module(self, module: str) -> AgentInfo | None:
        """找到能执行指定模块的 Agent"""
        async with self._lock:
            for info in self._agents.values():
                if module in info.modules:
                    return info
        return None

    async def get_agent(self, agent_id: str) -> AgentInfo | None:
        async with self._lock:
            return self._agents.get(agent_id)

    # -----------------------------------------------------------------
    # 命令下发
    # -----------------------------------------------------------------
    async def send_command(self, agent_id: str, command: dict) -> bool:
        """向指定 Agent 下发命令"""
        info = await self.get_agent(agent_id)
        if not info:
            return False
        try:
            text = json.dumps(command, ensure_ascii=False, default=str)
            await info.ws.send_text(text)
            return True
        except Exception:
            await self.unregister(agent_id)
            return False

    async def send_to_module_agent(self, module: str, command: dict) -> str | None:
        """找到能执行该模块的 Agent 并发送命令，返回 agent_id 或 None"""
        info = await self.get_agent_for_module(module)
        if not info:
            return None
        ok = await self.send_command(info.agent_id, command)
        return info.agent_id if ok else None


# 全局单例
agent_manager = AgentManager()
