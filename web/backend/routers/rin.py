"""RIN 测试 API 路由"""

import asyncio
import json
import threading
import uuid

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse

from ..services.ws_manager import ws_manager
from ..services.rin_service import HeadlessRinRunner

router = APIRouter(prefix="/api/rin", tags=["rin"])

# 全局状态（单线程运行 RIN 测试）
_rin_thread: threading.Thread | None = None
_rin_runner: HeadlessRinRunner | None = None
_rin_stop_flag: threading.Event = threading.Event()
_rin_result: dict = {}
_rin_logs: list[dict] = []
_rin_run_id: str = ""
_main_loop: asyncio.AbstractEventLoop | None = None


def _ensure_main_loop():
    """获取或缓存主线程事件循环"""
    global _main_loop
    if _main_loop is None or _main_loop.is_closed():
        try:
            _main_loop = asyncio.get_running_loop()
        except RuntimeError:
            _main_loop = asyncio.new_event_loop()
            asyncio.set_event_loop(_main_loop)
    return _main_loop


def _broadcast(msg_type: str, message: str, data: dict | None = None):
    """同步方式广播 WebSocket 消息（从后台线程调用）"""
    payload = {"message": message}
    if data:
        payload.update(data)
    loop = _ensure_main_loop()
    asyncio.run_coroutine_threadsafe(
        ws_manager.broadcast({
            "type": msg_type,
            "module": "Rin_FSV3004",
            "payload": payload,
        }),
        loop,
    )


def _progress_callback(msg: str, data: dict | None = None):
    """RIN 测试进度回调"""
    data = data or {}
    msg_type = data.pop("type", "log")
    _rin_logs.append({
        "timestamp": None, "level": msg_type, "message": msg, "data": data,
    })
    _broadcast(msg_type, msg, data)


def _run_rin_thread(params: dict):
    """在后台线程中运行 RIN 测试"""
    global _rin_result, _rin_runner
    _rin_stop_flag.clear()
    _rin_logs.clear()

    _rin_runner = HeadlessRinRunner(
        dc_value=params.get("dc_value", 2.4),
        ip_address=params.get("ip_address"),
        save_path=params.get("save_path"),
        progress_callback=_progress_callback,
        stop_flag=_rin_stop_flag,
    )

    try:
        _broadcast("status", "RIN 测试已启动", {"status": "running"})
        result = _rin_runner.run()
        _rin_result = result

        status = result.get("status", "error")
        if status == "completed":
            _broadcast("completed", "RIN 测试完成", {"status": "completed", "result": result})
        elif status == "stopped":
            _broadcast("status", "RIN 测试已停止", {"status": "stopped"})
        else:
            _broadcast("error", f"RIN 测试异常: {result.get('error', '未知错误')}",
                       {"status": "error"})
    except Exception as e:
        _rin_result = {"status": "error", "error": str(e)}
        _broadcast("error", f"RIN 测试线程异常: {e}", {"status": "error"})
    finally:
        _rin_runner = None


@router.post("/start")
async def start_rin(params: dict):
    """启动 RIN 测试（后台线程）

    Body: {
        "dc_value": 2.4,
        "ip_address": "192.168.7.10",    // 可选
        "save_path": "C:\\PTS\\..."      // 可选
    }
    """
    global _rin_thread, _rin_run_id, _main_loop

    # 在主线程中缓存事件循环，供后台线程广播用
    _main_loop = asyncio.get_running_loop()

    if _rin_thread and _rin_thread.is_alive():
        raise HTTPException(409, "RIN 测试已在运行中")

    _rin_run_id = uuid.uuid4().hex[:12]
    _rin_thread = threading.Thread(
        target=_run_rin_thread,
        args=(params,),
        daemon=True,
    )
    _rin_thread.start()

    return {
        "run_id": _rin_run_id,
        "status": "started",
        "message": "RIN 测试已启动，通过 WebSocket 接收实时进度",
    }


@router.get("/status")
async def get_rin_status():
    """获取当前 RIN 测试状态"""
    running = _rin_thread is not None and _rin_thread.is_alive()
    status = "running" if running else (_rin_result.get("status", "idle"))
    return {
        "running": running,
        "status": status,
        "run_id": _rin_run_id,
        "logs": _rin_logs[-100:],  # 最近 100 条日志
    }


@router.post("/stop")
async def stop_rin():
    """停止当前 RIN 测试"""
    if not (_rin_thread and _rin_thread.is_alive()):
        raise HTTPException(404, "没有正在运行的 RIN 测试")
    _rin_stop_flag.set()
    return {"status": "stopping", "message": "停止信号已发送"}


@router.get("/result")
async def get_rin_result():
    """获取最近一次 RIN 测试结果"""
    if not _rin_result:
        return {"status": "idle", "message": "尚无测试结果"}
    return _rin_result


@router.get("/chart")
async def get_rin_chart():
    """获取 RIN 图表 PNG"""
    png_path = _rin_result.get("rin_png", "")
    if not png_path or not __import__('os').path.exists(png_path):
        raise HTTPException(404, "图表文件不存在")
    return FileResponse(png_path, media_type="image/png")
