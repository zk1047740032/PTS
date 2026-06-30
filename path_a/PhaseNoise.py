#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
种子激光器相位噪声测试模块

该模块用于自动化执行种子激光器的相位噪声测量，
支持基于波长自动选择1μm或1.5μm程序并进行自动点击操作。
"""
from __future__ import annotations

import os
import time
import csv
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from typing import Optional
from pywinauto import Desktop
import ctypes

if os.name == 'nt':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        dpi = ctypes.windll.user32.GetDpiForSystem()
        scaling_factor = dpi / 96.0
    except Exception:
        scaling_factor = 1.0
else:
    scaling_factor = 1.0


def click_button():
    """
    点击目标窗口中的指定按钮

    使用pywinauto模拟鼠标点击操作，自动定位并点击
    LaserNoiseMeasurement_Manufactory窗口中的目标按钮。

    返回:
        bool: 点击成功返回True，失败返回False
    """
    try:
        desktop = Desktop(backend="win32")
        matched_windows = desktop.windows(
            class_name="SunAwtFrame",
            title="LaserNoiseMeasurement_Manufactory",
            visible_only=True
        )
        if not matched_windows:
            print("未找到目标窗口。")
            return False

        win = desktop.window(handle=matched_windows[0].handle)

        try:
            win.set_focus()
        except Exception:
            pass
        time.sleep(0.5)

        win.click_input(coords=(296, 675))
        print("已成功点击目标按钮！(相对坐标: X=296, Y=675)")
        return True

    except Exception as e:
        print(f"点击失败: {e}")
        return False


class PhaseNoiseGUI:
    """
    相位噪声测试GUI控制类

    提供图形化界面用于配置测试参数、控制测试流程、显示运行日志。
    支持根据种子波长自动选择对应的程序进行测试。
    """

    def __init__(self, parent: Optional[tk.Widget] = None) -> None:
        """
        初始化GUI

        参数:
            parent (tk.Widget, optional): 父控件。若为None则创建独立窗口。
        """
        self.parent = parent

        if parent is None:
            self.root = tk.Tk()
            self.root.title("PhaseNoise")
            self.root.geometry("1140x420")
            self.root.resizable(True, True)
            try:
                self.root.iconbitmap("PreciLasers.ico")
            except Exception:
                pass
        else:
            self.root = parent

        self.wavelength_path = tk.StringVar(value=r"C:\PTS\zhongzi\WaveLength\wavelength.csv")
        self.program_1um = tk.StringVar(value="")
        self.program_1_5um = tk.StringVar(value="")

        self.worker: Optional[threading.Thread] = None
        self.stop_flag = threading.Event()

        self._build_ui()

    def _build_ui(self) -> None:
        """
        构建GUI界面布局
        """
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        config_frame = tk.LabelFrame(left_frame, text="配置", padx=10, pady=10)
        config_frame.pack(fill=tk.X, pady=8)

        tk.Label(config_frame, text="种子波长文件:", anchor="w").pack(fill=tk.X)
        path_frame = tk.Frame(config_frame)
        path_frame.pack(fill=tk.X, pady=(0, 5))
        tk.Entry(path_frame, textvariable=self.wavelength_path, width=28).pack(side=tk.LEFT, fill=tk.X, expand=True)
        #tk.Button(path_frame, text="浏览", command=self._browse_wavelength, cursor="hand2", width=6).pack(side=tk.RIGHT, padx=(2, 0))

        tk.Label(config_frame, text="1μm程序:", anchor="w").pack(fill=tk.X)
        entry_1um = tk.Entry(config_frame, textvariable=self.program_1um, width=35)
        entry_1um.pack(pady=(0, 5))
        #tk.Button(config_frame, text="浏览", command=lambda: self._browse_program(self.program_1um), cursor="hand2", width=6).pack(anchor="e", padx=(0, 0))

        tk.Label(config_frame, text="1.5μm程序:", anchor="w").pack(fill=tk.X)
        entry_1_5um = tk.Entry(config_frame, textvariable=self.program_1_5um, width=35)
        entry_1_5um.pack(pady=(0, 5))
        #tk.Button(config_frame, text="浏览", command=lambda: self._browse_program(self.program_1_5um), cursor="hand2", width=6).pack(anchor="e")

        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=20)

        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor="center")

        self.start_btn = tk.Button(
            inner_btn_frame, text="开始测试",
            bg="#4CAF50", fg="#FFFFFF",
            command=self.start_test,
            cursor="hand2",
            width=12
        )
        self.start_btn.pack(side=tk.LEFT)

        log_frame = tk.LabelFrame(right_frame, text="运行日志", padx=6, pady=6)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(self.log_box)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_box.yview)

    def _browse_wavelength(self) -> None:
        """
        浏览选择种子波长CSV文件
        """
        filename = filedialog.askopenfilename(
            title="选择种子波长文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if filename:
            self.wavelength_path.set(filename)

    def _browse_program(self, var: tk.StringVar) -> None:
        """
        浏览选择程序文件

        参数:
            var: 程序路径变量
        """
        filename = filedialog.askopenfilename(
            title="选择程序",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if filename:
            var.set(filename)

    def log(self, msg: str) -> None:
        """
        线程安全地输出日志信息

        参数:
            msg (str): 日志消息
        """
        t = time.strftime("[%H:%M:%S]")
        self.root.after(0, lambda: self._safe_log_append(f"{t} {msg}\n"))

    def _safe_log_append(self, text: str) -> None:
        """
        线程安全地向日志文本框追加内容

        参数:
            text (str): 日志文本
        """
        self.log_box.insert(tk.END, text)
        self.log_box.see(tk.END)

    def _read_wavelength(self) -> Optional[float]:
        """
        读取种子波长CSV文件

        返回:
            float: 波长值(nm)，读取失败返回None
        """
        path = self.wavelength_path.get().strip()
        if not path:
            self.log("[错误] 种子波长文件路径为空")
            return None

        if not os.path.exists(path):
            self.log(f"[错误] 文件不存在: {path}")
            return None

        try:
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    if row and row[0].strip():
                        wavelength = float(row[0].strip())
                        self.log(f"[波长] 读取到波长: {wavelength} nm")
                        return wavelength
            self.log("[错误] 文件中未找到波长数据")
            return None
        except Exception as e:
            self.log(f"[错误] 读取波长文件失败: {e}")
            return None

    def start_test(self) -> None:
        """
        启动测试流程
        """
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "测试已在进行中")
            return

        self.start_btn.config(state=tk.DISABLED)

        self.stop_flag.clear()

        self.worker = threading.Thread(target=self._test_task, daemon=True)
        self.worker.start()

    def _test_task(self) -> None:
        """
        后台测试任务线程
        """
        try:
            self.log("=" * 30)
            self.log("[开始] 相位噪声测试")
            self.log("=" * 30)

            wavelength = self._read_wavelength()
            if wavelength is None:
                return

            if wavelength >= 1500:
                program_path = self.program_1_5um.get().strip()
                seed_type = "1.5μm"
            else:
                program_path = self.program_1um.get().strip()
                seed_type = "1μm"

            self.log(f"[判定] 种子类型: {seed_type} (波长: {wavelength} nm)")

            if not program_path:
                self.log(f"[错误] {seed_type}程序路径为空")
                return

            if not os.path.exists(program_path):
                self.log(f"[错误] 程序文件不存在: {program_path}")
                return

            self.log(f"[启动] 正在启动程序: {program_path}")
            # 获取程序所在的文件夹路径
            program_dir = os.path.dirname(program_path)
            # 指定 cwd 为程序所在文件夹
            subprocess.Popen([program_path], cwd=program_dir)
            self.log("[启动] 程序已启动，等待窗口加载...")
            time.sleep(3)

            if not self._interruptible_sleep(2):
                self.log("[停止] 用户中断测试")
                return

            self.log("[点击] 正在执行点击操作...")
            max_retries = 3
            retry_count = 0
            click_success = False
            
            while retry_count < max_retries and not click_success:
                if retry_count > 0:
                    self.log(f"[重试] 第 {retry_count} 次重试点击操作...")
                    time.sleep(1)
                if click_button():
                    click_success = True
                else:
                    retry_count += 1
            
            if click_success:
                self.log("[成功] 点击操作成功完成")
            else:
                self.log(f"[错误] 点击操作失败，已重试 {max_retries} 次")

        except Exception as e:
            self.log(f"[错误] 测试失败: {e}")
        finally:
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            # 一键测试模式：自动关闭窗口，触发进程退出
            self.root.after(3000, self.root.destroy)

    def _interruptible_sleep(self, seconds: float, check_interval: float = 0.5) -> bool:
        """
        可中断的睡眠

        参数:
            seconds: 总睡眠时间（秒）
            check_interval: 检查间隔（秒）

        返回:
            bool: True表示正常完成，False表示被中断
        """
        elapsed = 0.0
        while elapsed < seconds:
            if self.stop_flag.is_set():
                return False
            sleep_time = min(check_interval, seconds - elapsed)
            time.sleep(sleep_time)
            elapsed += sleep_time
        return True

    def run(self) -> None:
        """
        以独立窗口模式运行
        """
        if self.root.winfo_exists():
            self.root.mainloop()

if __name__ == "__main__":
    gui = PhaseNoiseGUI()
    gui.run()

"""
& "d:\Coding\Project\PTS\zhongzi\.venv\Scripts\python.exe" -m PyInstaller path_a/PhaseNoise.py --onefile --noconsole --hidden-import=pywinauto --add-data "PreciLasers.ico;."
"""