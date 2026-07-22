#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core.base_gui —— 测试程序 GUI 基类

吸收 path_a / path_b 中 8 个 *GUI 类的共性：
    - 嵌套式窗口构造（parent is None 建 tk.Tk，否则复用 parent）
    - 线程安全日志（root.after + log_box insert/see）
    - 后台测试线程（daemon Thread + stop_flag）
    - 可中断睡眠
    - 文件/目录浏览对话框
    - run() 主循环

布局逻辑（_build_ui）留给子类——各测试界面参数框数量、按钮个数、
是否有 1μm/1.5μm 切换差异太大，强行抽象布局只会更乱。
"""

from __future__ import annotations

import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Callable, Optional, Tuple

from .utils import interruptible_sleep

__all__ = ["BaseTestGUI"]


class BaseTestGUI:
    """
    测试程序 GUI 基类。

    使用示例::

        class PhaseNoiseGUI(BaseTestGUI):
            def __init__(self, parent=None):
                super().__init__(parent, title="PhaseNoise",
                                 geometry="1140x420", icon="PreciLasers.ico")
                self._build_ui()   # 仅保留子类布局，创建 self.log_box
            def _build_ui(self):
                ...
                self.log_box = tk.Text(log_frame)
                ...

    约定：
        - 子类在 _build_ui 中必须创建 ``self.log_box``（tk.Text），log() 才有输出目标。
        - 自动关窗（root.after(2000~3000, self.root.destroy)）是"一键测试模式"行为，
          且延时数值各程序不同，故不放入基类，由子类的 finally 块自行调用
          ``self.auto_close(delay_ms=3000)``。

    属性:
        parent: 父控件（集成模式下复用）。
        root: 实际承载控件的根窗口（独立模式为 tk.Tk，集成模式为 parent）。
        stop_flag: threading.Event，与后台线程共享以实现中断。
        worker: 当前后台测试线程。
    """

    def __init__(
        self,
        parent: Optional[tk.Widget] = None,
        *,
        title: str = "测试",
        geometry: str = "1100x600",
        icon: Optional[str] = None,
        resizable: Tuple[bool, bool] = (True, True),
    ) -> None:
        """
        初始化 GUI 基类。

            参数:
                parent (tk.Widget, optional): 父控件。None 则创建独立窗口，否则复用 parent 作为 root。
                title (str): 独立模式窗口标题。
                geometry (str): 独立模式窗口尺寸，如 "1140x420"。
                icon (str, optional): 图标文件路径（.ico）。加载失败静默忽略。
                resizable (tuple): (宽可调, 高可调)。
        """
        self.parent = parent

        if parent is None:
            self.root = tk.Tk()
            self.root.title(title)
            self.root.geometry(geometry)
            self.root.resizable(*resizable)
            if icon:
                try:
                    self.root.iconbitmap(icon)
                except Exception:
                    pass
        else:
            self.root = parent

        self.stop_flag = threading.Event()
        self.worker: Optional[threading.Thread] = None

    # ------------------------------------------------------------------
    # 线程安全日志
    # ------------------------------------------------------------------
    def log(self, msg: str) -> None:
        """
        线程安全地输出日志：通过 root.after 切回主线程写入 log_box。

        子类必须在 _build_ui 中创建 self.log_box（tk.Text），否则日志仅打印到控制台。

            参数:
                msg (str): 日志消息（不含时间戳，由本方法统一加前缀）。
        """
        t = time.strftime("[%H:%M:%S]")
        self.root.after(0, lambda: self._safe_log_append(f"{t} {msg}\n"))

    def _safe_log_append(self, text: str) -> None:
        """
        在主线程中向日志框追加文本并滚动到末尾；无 log_box 时退化为 print。

            参数:
                text (str): 已格式化的日志文本（含换行）。
        """
        log_box = getattr(self, "log_box", None)
        if log_box is not None:
            try:
                log_box.insert(tk.END, text)
                log_box.see(tk.END)
            except Exception:
                pass
        print(text, end="")

    # ------------------------------------------------------------------
    # 后台线程
    # ------------------------------------------------------------------
    def start_worker(
        self,
        target: Callable[[], None],
        *,
        busy_msg: str = "测试已在进行中",
    ) -> bool:
        """
        启动后台 daemon 线程执行 target。已有进行中的线程时弹提示并拒绝。

        会先清除 stop_flag，交给 target 内部在循环中自查。

            参数:
                target (callable): 线程入口，无参无返回（异常自行 try/except 或交给本框架）。
                busy_msg (str): 重复启动时的提示文案。

            返回:
                bool: True 表示已启动新线程，False 表示已有线程在跑未启动。
        """
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", busy_msg)
            return False
        self.stop_flag.clear()
        self.worker = threading.Thread(target=target, daemon=True)
        self.worker.start()
        return True

    # ------------------------------------------------------------------
    # 可中断睡眠
    # ------------------------------------------------------------------
    def interruptible_sleep(
        self, seconds: float, check_interval: float = 0.5
    ) -> bool:
        """
        可中断的睡眠，绑定本 GUI 的 stop_flag。

            参数:
                seconds (float): 总睡眠时间（秒）。
                check_interval (float): 检查停止标志的间隔（秒）。

            返回:
                bool: True 正常完成，False 被 stop_flag 中断。
        """
        return interruptible_sleep(seconds, self.stop_flag, check_interval)

    # ------------------------------------------------------------------
    # 浏览对话框
    # ------------------------------------------------------------------
    def browse_file(
        self,
        entry: tk.Entry,
        title: str = "选择文件",
        filetypes=None,
    ) -> None:
        """
        弹出"打开文件"对话框，选中后填入 entry。

            参数:
                entry (tk.Entry): 结果写入的输入框。
                title (str): 对话框标题。
                filetypes: 文件类型过滤，默认 [("所有文件", "*.*")]。
        """
        if filetypes is None:
            filetypes = [("所有文件", "*.*")]
        filename = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if filename:
            entry.delete(0, tk.END)
            entry.insert(0, filename)

    def browse_dir(self, entry: tk.Entry, title: str = "选择保存目录") -> None:
        """
        弹出"选择目录"对话框，选中后填入 entry。

            参数:
                entry (tk.Entry): 结果写入的输入框。
                title (str): 对话框标题。
        """
        dirname = filedialog.askdirectory(title=title)
        if dirname:
            entry.delete(0, tk.END)
            entry.insert(0, dirname)

    # ------------------------------------------------------------------
    # 一键测试模式标记 & 自动关窗
    # ------------------------------------------------------------------
    _one_click_test: bool = False
    """是否由"一键测试"启动。仅当为 True 时 auto_close() 才真正关闭窗口；
    通过"打开"按钮或双击独立打开的程序设为 False，测试结束后窗口保持打开。"""

    def auto_close(self, delay_ms: int = 3000) -> None:
        """
        延时关闭窗口（仅"一键测试模式"生效）。

        若窗口不是由"一键测试"启动（_one_click_test=False），则跳过关闭，
        测试结束后窗口保持打开以便用户查看结果。

            参数:
                delay_ms (int): 关闭前延时（毫秒），默认 3000。
        """
        if not self._one_click_test:
            return
        self.root.after(delay_ms, self.root.destroy)

    # ------------------------------------------------------------------
    # 主循环
    # ------------------------------------------------------------------
    def run(self) -> None:
        """启动 Tk 主事件循环（仅独立窗口模式有意义）。"""
        if self.root.winfo_exists():
            self.root.mainloop()
