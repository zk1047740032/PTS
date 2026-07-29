"""测试生命周期管理 — 通过 multiprocessing.Process 拉起测试程序"""

import asyncio
import multiprocessing
import queue
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from ..models.test_result import TestRun, TestLog
from .ws_manager import ws_manager

# 确保项目根目录在 sys.path 中
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

# 模块注册表（与 main_platform.py 保持一致）
MODULE_REGISTRY: dict[str, tuple[str, str]] = {
    "Rin_FSV3004":    ("path_a.Rin_FSV3004",    "RinGUI"),
    "线宽_FSV3004":   ("path_a.LineWidth_FSV3004", "LineWidth_FSV3004_GUI"),
    "时域":           ("path_a.TimeDomain",      "TimeDomainGUI"),
    "信噪比":         ("path_a.SpectrumSNR",     "SpectrumSNRGUI"),
    "单频":           ("path_a.SingleFrequency", "SingleFrequencyGUI"),
    "功率":           ("path_b.Power",           "PowerGUI"),
    "PZT调制":        ("path_a.WaveLength",      "WaveLengthTestGUI"),
    "相噪":           ("path_a.PhaseNoise",      "PhaseNoiseGUI"),
}

# 模块到通道的映射
MODULE_CHANNEL_MAP: dict[str, str] = {
    "PZT调制": "通道1",
    "时域": "通道2",
    "Rin_FSV3004": "通道3",
    "线宽_FSV3004": "通道3",
    "信噪比": "通道3",
    "单频": "通道3",
    "相噪": "通道4",
    "功率": "光路B",
}

# 模块启动方法名
MODULE_START_METHODS: dict[str, str] = {
    "Rin_FSV3004": "start_rin",
    "线宽_FSV3004": "start_measurement",
    "时域": "start_test",
    "信噪比": "start_test",
    "单频": "start_test",
    "功率": "start_test",
    "PZT调制": "start_test",
    "相噪": "start_test",
}


def _run_module_in_subprocess(module_key: str, cmd_queue: multiprocessing.Queue, msg_queue: multiprocessing.Queue):
    """在子进程中运行测试模块 GUI（与 main_platform.py 逻辑一致）"""
    import importlib

    module_path, gui_class_name = MODULE_REGISTRY[module_key]
    gui_module = importlib.import_module(module_path)
    gui_class = getattr(gui_module, gui_class_name)

    # 应用用户配置覆盖
    try:
        from utils.config_params_panel import apply_user_overrides_to_cfg
        apply_user_overrides_to_cfg()
    except Exception:
        pass

    app_instance = gui_class(None)
    if hasattr(app_instance, 'root'):
        app_instance.root.title(f"{module_key} [就绪]")

    msg_queue.put((module_key, "running", f"{module_key} 窗口已打开"))

    def check_queue():
        """轮询命令队列"""
        try:
            cmd = cmd_queue.get_nowait()
            if cmd == "START":
                start_method_name = MODULE_START_METHODS.get(module_key, "start_test")
                if hasattr(app_instance, start_method_name):
                    msg_queue.put((module_key, "running", f"开始执行 {module_key} 测试..."))
                    getattr(app_instance, start_method_name)()
        except queue.Empty:
            pass
        except Exception as e:
            msg_queue.put((module_key, "error", str(e)))
        # 继续轮询
        if hasattr(app_instance, 'root'):
            app_instance.root.after(200, check_queue)

    if hasattr(app_instance, 'root'):
        app_instance.root.after(500, check_queue)
        app_instance.root.mainloop()

    msg_queue.put((module_key, "completed", f"{module_key} 测试完成"))


class TestService:
    """管理测试进程池"""

    def __init__(self):
        self._processes: dict[str, multiprocessing.Process] = {}
        self._cmd_queues: dict[str, multiprocessing.Queue] = {}
        self._msg_queue: multiprocessing.Queue = multiprocessing.Queue()
        self._run_records: dict[str, TestRun] = {}
        self._log_seq: dict[str, int] = {}
        self._polling = False

    @property
    def running_modules(self) -> list[str]:
        return [k for k, p in self._processes.items() if p.is_alive()]

    async def start_test(self, module_key: str, auto_start: bool = True, db_session=None) -> str | None:
        """启动测试模块，返回 run_id"""
        if module_key in self._processes and self._processes[module_key].is_alive():
            return None  # 已经在运行

        cmd_queue: multiprocessing.Queue = multiprocessing.Queue()
        self._cmd_queues[module_key] = cmd_queue

        proc = multiprocessing.Process(
            target=_run_module_in_subprocess,
            args=(module_key, cmd_queue, self._msg_queue),
            daemon=True,
        )
        proc.start()
        self._processes[module_key] = proc

        if auto_start:
            cmd_queue.put("START")

        # 创建数据库记录
        run = TestRun(
            module_name=module_key,
            module_label=module_key,
            channel=MODULE_CHANNEL_MAP.get(module_key, ""),
            status="running",
            start_time=datetime.now(),
        )
        self._run_records[module_key] = run
        self._log_seq[module_key] = 0

        if db_session:
            db_session.add(run)
            await db_session.commit()

        # 启动消息轮询
        if not self._polling:
            self._polling = True
            asyncio.create_task(self._poll_messages())

        return run.id

    async def stop_test(self, module_key: str) -> bool:
        """停止测试"""
        proc = self._processes.get(module_key)
        if not proc or not proc.is_alive():
            return False
        proc.terminate()
        proc.join(timeout=5)
        if proc.is_alive():
            proc.kill()
        self._processes.pop(module_key, None)
        self._cmd_queues.pop(module_key, None)

        # 更新记录
        run = self._run_records.get(module_key)
        if run:
            run.status = "stopped"
            run.end_time = datetime.now()

        await ws_manager.broadcast({
            "type": "status",
            "module": module_key,
            "payload": {"status": "stopped"},
        })
        return True

    async def _poll_messages(self):
        """轮询子进程消息队列，广播到 WebSocket"""
        while self._processes:
            try:
                msg = self._msg_queue.get_nowait()
                module_key, msg_type, text = msg

                seq = self._log_seq.get(module_key, 0)
                self._log_seq[module_key] = seq + 1

                await ws_manager.send_to_module(module_key, {
                    "type": msg_type,
                    "module": module_key,
                    "payload": {"message": text, "seq": seq},
                })
                await ws_manager.broadcast_global({
                    "type": msg_type,
                    "module": module_key,
                    "payload": {"message": text, "seq": seq},
                })

                # 更新 run 状态
                if msg_type == "completed" or msg_type == "error":
                    run = self._run_records.get(module_key)
                    if run:
                        run.status = msg_type
                        run.end_time = datetime.now()
                        if run.start_time:
                            run.duration_seconds = (run.end_time - run.start_time).total_seconds()

            except queue.Empty:
                pass
            except Exception:
                pass

            # 清理已结束的进程
            dead = [k for k, p in self._processes.items() if not p.is_alive()]
            for k in dead:
                self._processes.pop(k, None)
                self._cmd_queues.pop(k, None)

            await asyncio.sleep(0.3)

        self._polling = False


# 全局单例
test_service = TestService()
