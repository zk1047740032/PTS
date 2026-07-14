#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core.base_runner —— 测试流程编排基类（模板方法）

吸收 path_a / path_b 中重复出现的测试流程骨架：
    清空输出目录 → 连接仪器 → 配置参数 → 循环测量 → 收尾 → 关闭资源

参考 path_a/Rin_FSV3004.py 中已存在的 TestRunner 雏形（签名绑死 run_rin），
将其泛化为通用模板方法。

不强制所有测试都套用本基类——如 path_a/SingleFrequency.py 的双并行控制
线程（温度 + 电流）+ 实时参数轮询等复杂编排，可选择只继承 VisaInstrument
+ BaseTestGUI 而不用 runner。
"""

from __future__ import annotations

import threading
import traceback
from tkinter import messagebox
from typing import Callable, Optional, TYPE_CHECKING

if TYPE_CHECKING:  # 仅用于类型注解，避免运行时强制依赖 tkinter
    from .base_gui import BaseTestGUI

__all__ = ["BaseTestRunner"]


class BaseTestRunner:
    """
    测试流程编排基类（模板方法模式）。

    模板方法 ``run()`` 固化通用骨架，子类按需覆写下列 hook（默认空实现/默认值）：

        prepare()      —— 清空输出目录等初始化
        connect()      —— 连接仪器，返回 bool（False 则终止）
        configure()    —— 写入仪器测量参数
        measure_loop() —— 测试主体（内部自查 stop_flag，调用 interruptible_sleep）
        finalize()     —— 收尾（数据可视化、结果弹窗等）
        cleanup()      —— 关闭仪器、恢复初始状态

    使用示例::

        class LinewidthRunner(BaseTestRunner):
            def connect(self) -> bool:
                self.tester = LinewidthTester(log_callback=self.log)
                return self.tester.connect(self.ip)
            def measure_loop(self):
                for span in self.span_values:
                    if self.stop_flag.is_set(): break
                    ...
            def cleanup(self):
                if self.tester: self.tester.close()

    属性:
        gui: 关联的 BaseTestGUI（可空），用于共享 stop_flag 与 root.after。
        log: 日志回调（优先用 gui.log，其次 log_func，再退化为 print）。
        stop_flag: 停止事件。关联 gui 时与其共享同一个 Event，保证 GUI 的
            "停止"按钮能中断 runner。
    """

    def __init__(
        self,
        *,
        gui: Optional["BaseTestGUI"] = None,
        log_func: Optional[Callable[[str], None]] = None,
    ) -> None:
        """
        初始化测试编排基类。

            参数:
                gui (BaseTestGUI, optional): 关联的 GUI。提供后，日志走 gui.log、
                    stop_flag 与 GUI 共享、异常弹窗走 gui.root.after。
                log_func (callable, optional): 无 gui 时的日志回调。
        """
        self.gui = gui
        # 日志优先级：gui.log > log_func > print
        if gui is not None and hasattr(gui, "log"):
            self.log: Callable[[str], None] = gui.log
        elif log_func is not None:
            self.log = log_func
        else:
            self.log = print

        # stop_flag 与 GUI 共享：GUI 的"停止"按钮 set 后，本 runner 的循环自查即中断
        self.stop_flag: threading.Event = (
            gui.stop_flag if gui is not None else threading.Event()
        )

    # ------------------------------------------------------------------
    # 模板方法
    # ------------------------------------------------------------------
    def run(self) -> None:
        """
        执行测试流程模板。调用顺序：
        prepare → connect(失败则终止) → configure → measure_loop → finalize，
        全程异常捕获并记录，finally 中执行 cleanup 并（关联 GUI 时）自动关窗。

        自动关窗延时默认 2000ms，子类可覆写 auto_close_delay_ms 改变。
        """
        try:
            self.prepare()
            if not self.connect():
                self.log("[流程] 连接失败，终止测试")
                return
            self.configure()
            self.measure_loop()
            self.finalize()
        except Exception as e:
            tb = traceback.format_exc()
            self.log(f"[错误] 测试失败：{e}\n{tb}")
            if self.gui is not None:
                self.gui.root.after(
                    0, lambda err=str(e): messagebox.showerror("错误", err)
                )
        finally:
            self.cleanup()
            if self.gui is not None:
                self.gui.auto_close(self.auto_close_delay_ms)

    # ------------------------------------------------------------------
    # 子类钩子（默认实现）
    # ------------------------------------------------------------------
    auto_close_delay_ms: int = 2000
    """一键测试模式自动关窗延时（毫秒）。子类可覆写为类属性。"""

    def prepare(self) -> None:
        """初始化准备，如清空输出目录、创建子控制器。默认空。"""
        return None

    def connect(self) -> bool:
        """连接仪器。返回 False 则终止后续流程。默认返回 True（无连接需求）。"""
        return True

    def configure(self) -> None:
        """写入仪器测量参数。默认空。"""
        return None

    def measure_loop(self) -> None:
        """测试主体循环。子类应在循环中自查 self.stop_flag。默认空。"""
        return None

    def finalize(self) -> None:
        """收尾（可视化、结果弹窗等）。默认空。"""
        return None

    def cleanup(self) -> None:
        """关闭仪器、恢复初始状态。默认空。即使前面抛异常也会执行。"""
        return None
