"""PTS 仪器 Agent — 部署在连接仪器的 PC 上，接收服务器命令执行测试"""

import argparse
import asyncio
import json
import signal
import socket
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

import websockets

# 确保能 import 项目模块
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

from web.agent.test_executor import TestExecutor


class InstrumentAgent:
    """仪器 Agent：连接服务器 → 注册 → 接收命令 → 执行测试 → 回传结果"""

    def __init__(self, server_url: str, agent_name: str):
        self.server_url = server_url
        self.agent_name = agent_name
        self.hostname = socket.gethostname()
        self.ws: websockets.ClientConnection | None = None
        self.executor = TestExecutor(callback=self._send)
        self._running = True
        self._reconnect_delay = 3

    # -----------------------------------------------------------------
    # 生命周期
    # -----------------------------------------------------------------
    async def run(self):
        """主循环：连接 → 注册 → 监听命令"""
        while self._running:
            try:
                await self._connect_and_listen()
            except (websockets.ConnectionClosed, OSError) as e:
                print(f"[Agent] 连接断开: {e}，{self._reconnect_delay}s 后重连...")
            except Exception as e:
                print(f"[Agent] 异常: {e}")
                traceback.print_exc()
            if self._running:
                await asyncio.sleep(self._reconnect_delay)

    async def _connect_and_listen(self):
        async with websockets.connect(self.server_url, ping_interval=30) as ws:
            self.ws = ws
            # 注册
            await self._register()
            print(f"[Agent] 已注册到服务器 {self.server_url}，在线模块: {self.executor.available_modules}")
            # 监听命令
            async for raw in ws:
                try:
                    msg = json.loads(raw)
                except json.JSONDecodeError:
                    continue
                await self._handle_command(msg)

    def stop(self):
        self._running = False
        if self.ws:
            asyncio.ensure_future(self.ws.close())

    # -----------------------------------------------------------------
    # 注册
    # -----------------------------------------------------------------
    async def _register(self):
        await self._send({
            "type": "register",
            "agent_id": self.agent_name,
            "hostname": self.hostname,
            "modules": self.executor.available_modules,
        })

    # -----------------------------------------------------------------
    # 命令分发
    # -----------------------------------------------------------------
    async def _handle_command(self, msg: dict):
        cmd_type = msg.get("type", "")
        test_id = msg.get("test_id", "")

        if cmd_type == "start_test":
            module = msg.get("module", "")
            params = msg.get("params", {})
            print(f"[Agent] 收到启动命令: {module} (test_id={test_id})")
            # 在新线程中执行测试（不阻塞 WS 消息循环）
            threading.Thread(
                target=self.executor.run_test,
                args=(module, params, test_id),
                daemon=True,
            ).start()

        elif cmd_type == "stop_test":
            print(f"[Agent] 收到停止命令: test_id={test_id}")
            self.executor.stop_test(test_id)

        elif cmd_type == "ping":
            await self._send({"type": "pong"})

    # -----------------------------------------------------------------
    # 发送消息到服务器
    # -----------------------------------------------------------------
    async def _send(self, data: dict):
        if self.ws:
            try:
                await self.ws.send(json.dumps(data, ensure_ascii=False, default=str))
            except websockets.ConnectionClosed:
                pass


# =====================================================================
# 入口
# =====================================================================
def main():
    parser = argparse.ArgumentParser(description="PTS 仪器 Agent")
    parser.add_argument("--server", default="ws://127.0.0.1:8088/agent/ws", help="服务器 WebSocket 地址")
    parser.add_argument("--name", default=socket.gethostname(), help="Agent 名称")
    args = parser.parse_args()

    agent = InstrumentAgent(server_url=args.server, agent_name=args.name)
    print(f"PTS Agent [{args.name}] 启动中，连接 {args.server} ...")
    print(f"本机: {agent.hostname}")

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    # 优雅退出
    def shutdown(sig, frame):
        print("\n[Agent] 正在关闭...")
        agent.stop()
        loop.stop()
    signal.signal(signal.SIGINT, shutdown)
    signal.signal(signal.SIGTERM, shutdown)

    try:
        loop.run_until_complete(agent.run())
    except KeyboardInterrupt:
        pass
    finally:
        loop.close()


if __name__ == "__main__":
    main()
