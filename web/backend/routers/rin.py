"""RIN 测试 API 路由 — 通过 Agent 或本地执行"""

import asyncio
import threading
import uuid

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse

from ..services.agent_manager import agent_manager
from ..services.ws_manager import ws_manager

router = APIRouter(prefix="/api/rin", tags=["rin"])

# 全局状态
_rin_result: dict = {}
_rin_run_id: str = ""

# ---------- 本地执行（无 Agent 时 fallback）----------
_local_thread: threading.Thread | None = None
_local_stop_flag: threading.Event = threading.Event()
_local_main_loop: asyncio.AbstractEventLoop | None = None


def _run_rin_locally(params: dict):
    """本地后台线程执行 RIN 测试"""
    global _rin_result, _local_main_loop
    from ..services.rin_service import HeadlessRinRunner

    _local_stop_flag.clear()

    def _broadcast(msg_type, msg, **extra):
        payload = {"message": msg, **extra}
        loop = _local_main_loop
        if loop and not loop.is_closed():
            asyncio.run_coroutine_threadsafe(
                ws_manager.broadcast({
                    "type": msg_type, "module": "Rin_FSV3004", "payload": payload,
                }),
                loop,
            )

    def _cb(msg, data=None):
        data = data or {}
        t = data.pop("type", "log")
        _broadcast(t, msg, **data)

    runner = HeadlessRinRunner(
        dc_value=params.get("dc_value", 2.4),
        ip_address=params.get("ip_address"),
        save_path=params.get("save_path"),
        progress_callback=_cb,
        stop_flag=_local_stop_flag,
    )

    try:
        _broadcast("status", "RIN 测试已启动 [本地]", status="running")
        result = runner.run()
        _rin_result = result
        s = result.get("status", "error")
        if s == "completed":
            _broadcast("completed", "RIN 测试完成", status="completed", result=result)
        elif s == "stopped":
            _broadcast("status", "RIN 测试已停止", status="stopped")
        else:
            _broadcast("error", f"RIN 测试异常", status="error", error=result.get("error", ""))
    except Exception as e:
        _rin_result = {"status": "error", "error": str(e)}
        _broadcast("error", str(e), status="error")


@router.post("/start")
async def start_rin(params: dict):
    """启动 RIN 测试（优先走 Agent，无 Agent 时本地执行）"""
    global _rin_run_id, _local_main_loop

    dc_value = params.get("dc_value", 2.4)
    test_id = uuid.uuid4().hex[:12]
    _rin_run_id = test_id

    # 尝试通过 Agent 执行
    agent_id = await agent_manager.send_to_module_agent("Rin_FSV3004", {
        "type": "start_test",
        "test_id": test_id,
        "module": "Rin_FSV3004",
        "params": {"dc_value": dc_value, "ip_address": params.get("ip_address")},
    })

    if agent_id:
        return {
            "run_id": test_id, "status": "started", "mode": "agent",
            "agent_id": agent_id,
            "message": f"RIN 测试已下发到 Agent [{agent_id}]，通过 WebSocket 接收进度",
        }

    # 无 Agent，本地执行
    global _local_thread
    if _local_thread and _local_thread.is_alive():
        raise HTTPException(409, "RIN 测试已在本地运行中")

    _local_main_loop = asyncio.get_running_loop()
    _local_thread = threading.Thread(target=_run_rin_locally, args=(params,), daemon=True)
    _local_thread.start()

    return {
        "run_id": test_id, "status": "started", "mode": "local",
        "message": "RIN 测试已在本机启动（无可用 Agent），通过 WebSocket 接收进度",
    }


@router.get("/status")
async def get_rin_status():
    """获取当前 RIN 测试状态"""
    local_running = _local_thread is not None and _local_thread.is_alive()
    status = "running" if local_running else (_rin_result.get("status", "idle"))
    return {
        "running": local_running,
        "status": status,
        "run_id": _rin_run_id,
    }


@router.post("/stop")
async def stop_rin():
    """停止当前 RIN 测试"""
    # 先尝试 Agent
    agent = await agent_manager.get_agent_for_module("Rin_FSV3004")
    if agent:
        await agent_manager.send_command(agent.agent_id, {
            "type": "stop_test", "test_id": _rin_run_id,
        })
        return {"status": "stopping", "mode": "agent", "agent_id": agent.agent_id}

    # 本地停止
    if not (_local_thread and _local_thread.is_alive()):
        raise HTTPException(404, "没有正在运行的 RIN 测试")
    _local_stop_flag.set()
    return {"status": "stopping", "mode": "local"}


@router.get("/result")
async def get_rin_result():
    """获取最近一次 RIN 测试结果"""
    if not _rin_result:
        return {"status": "idle", "message": "尚无测试结果"}
    return _rin_result


@router.get("/chart")
async def get_rin_chart():
    """获取 RIN 图表 PNG"""
    import os
    png_path = _rin_result.get("rin_png", "")
    if not png_path or not os.path.exists(png_path):
        raise HTTPException(404, "图表文件不存在")
    return FileResponse(png_path, media_type="image/png")


@router.get("/agents")
async def list_agents():
    """获取在线 Agent 列表"""
    return {"agents": await agent_manager.get_online_agents()}
