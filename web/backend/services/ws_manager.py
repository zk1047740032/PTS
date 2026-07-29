"""WebSocket 连接管理器"""

import asyncio
import json
from typing import Any

from fastapi import WebSocket


class WSManager:
    """管理所有 WebSocket 连接，支持按模块订阅广播"""

    def __init__(self):
        # module_name -> set of WebSocket
        self._module_subscribers: dict[str, set[WebSocket]] = {}
        # 全局广播连接（Dashboard 等）
        self._global_connections: set[WebSocket] = set()
        self._lock = asyncio.Lock()

    async def connect(self, ws: WebSocket, module_name: str | None = None):
        """接受连接并注册"""
        await ws.accept()
        async with self._lock:
            if module_name:
                self._module_subscribers.setdefault(module_name, set()).add(ws)
            else:
                self._global_connections.add(ws)

    async def disconnect(self, ws: WebSocket, module_name: str | None = None):
        """移除连接"""
        async with self._lock:
            if module_name and module_name in self._module_subscribers:
                self._module_subscribers[module_name].discard(ws)
            self._global_connections.discard(ws)

    async def broadcast(self, message: dict[str, Any]):
        """广播到所有连接的客户端"""
        text = json.dumps(message, ensure_ascii=False, default=str)
        dead: list[WebSocket] = []
        async with self._lock:
            targets = list(self._global_connections)
            for module_set in self._module_subscribers.values():
                targets.extend(module_set)
        # 去重
        for ws in set(targets):
            try:
                await ws.send_text(text)
            except Exception:
                dead.append(ws)
        if dead:
            async with self._lock:
                for ws in dead:
                    self._global_connections.discard(ws)
                    for s in self._module_subscribers.values():
                        s.discard(ws)

    async def send_to_module(self, module_name: str, message: dict[str, Any]):
        """发送消息给订阅了特定模块的客户端"""
        text = json.dumps(message, ensure_ascii=False, default=str)
        dead: list[WebSocket] = []
        async with self._lock:
            subs = self._module_subscribers.get(module_name, set())
        for ws in list(subs):
            try:
                await ws.send_text(text)
            except Exception:
                dead.append(ws)
        if dead:
            async with self._lock:
                for ws in dead:
                    subs.discard(ws)

    async def broadcast_global(self, message: dict[str, Any]):
        """仅广播到全局连接"""
        text = json.dumps(message, ensure_ascii=False, default=str)
        dead: list[WebSocket] = []
        async with self._lock:
            targets = list(self._global_connections)
        for ws in targets:
            try:
                await ws.send_text(text)
            except Exception:
                dead.append(ws)
        if dead:
            async with self._lock:
                for ws in dead:
                    self._global_connections.discard(ws)


# 全局单例
ws_manager = WSManager()
