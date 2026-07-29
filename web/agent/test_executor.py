"""测试执行器 — 接收模块名 + 参数 → 加载测试 → 执行 → 回传进度"""

import threading
import traceback
from typing import Callable


class TestExecutor:
    """管理测试运行的生命周期"""

    def __init__(self, callback: Callable | None = None):
        self._callback = callback or (lambda d: None)
        self._stop_flags: dict[str, threading.Event] = {}
        self._threads: dict[str, threading.Thread] = {}
        # 注册本 Agent 能跑的测试模块
        self.available_modules = self._discover_modules()

    # -----------------------------------------------------------------
    # 模块发现
    # -----------------------------------------------------------------
    def _discover_modules(self) -> list[str]:
        """扫描可用的测试模块（目前仅 RIN，后续陆续加入）"""
        modules = ["Rin_FSV3004"]
        # TODO: 后续加入 "信噪比", "时域", "线宽_FSV3004", "单频", "功率", "PZT调制", "相噪"
        return modules

    # -----------------------------------------------------------------
    # 运行测试
    # -----------------------------------------------------------------
    def run_test(self, module: str, params: dict, test_id: str):
        """在新线程中运行指定模块的测试"""
        if module == "Rin_FSV3004":
            self._run_rin(params, test_id)
        else:
            self._send({
                "type": "error",
                "test_id": test_id,
                "module": module,
                "error": f"不支持的模块: {module}",
            })

    def _run_rin(self, params: dict, test_id: str):
        """运行 RIN 测试（Headless）"""
        from web.backend.services.rin_service import HeadlessRinRunner

        stop_flag = threading.Event()
        self._stop_flags[test_id] = stop_flag

        def do_run():
            self._send({"type": "log", "test_id": test_id, "module": "Rin_FSV3004",
                         "message": "正在导入仪器控制模块..."})

            runner = HeadlessRinRunner(
                dc_value=params.get("dc_value", 2.4),
                ip_address=params.get("ip_address"),
                save_path=params.get("save_path"),
                progress_callback=lambda msg, data: self._on_progress(test_id, msg, data),
                stop_flag=stop_flag,
            )

            try:
                self._send({"type": "log", "test_id": test_id, "module": "Rin_FSV3004",
                             "message": "RIN 测试已启动"})
                result = runner.run()

                status = result.get("status", "error")
                if status == "completed":
                    self._send({"type": "completed", "test_id": test_id,
                                 "module": "Rin_FSV3004",
                                 "message": "RIN 测试完成", "result": result})
                elif status == "stopped":
                    self._send({"type": "status", "test_id": test_id,
                                 "module": "Rin_FSV3004",
                                 "message": "RIN 测试已停止", "status": "stopped"})
                else:
                    self._send({"type": "error", "test_id": test_id,
                                 "module": "Rin_FSV3004",
                                 "message": f"测试异常: {result.get('error', '')}"})
            except Exception as e:
                self._send({"type": "error", "test_id": test_id,
                             "module": "Rin_FSV3004",
                             "message": f"测试线程异常: {e}\n{traceback.format_exc()}"})
            finally:
                self._stop_flags.pop(test_id, None)

        thread = threading.Thread(target=do_run, daemon=True)
        self._threads[test_id] = thread
        thread.start()

    # -----------------------------------------------------------------
    # 停止测试
    # -----------------------------------------------------------------
    def stop_test(self, test_id: str):
        flag = self._stop_flags.get(test_id)
        if flag:
            flag.set()

    # -----------------------------------------------------------------
    # 回调
    # -----------------------------------------------------------------
    def _on_progress(self, test_id: str, msg: str, data: dict | None = None):
        data = data or {}
        msg_type = data.pop("type", "log")
        payload = {"type": msg_type, "test_id": test_id, "module": "Rin_FSV3004",
                   "message": msg}
        payload.update(data)
        self._send(payload)

    def _send(self, data: dict):
        try:
            self._callback(data)
        except Exception:
            pass
