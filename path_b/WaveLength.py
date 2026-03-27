#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
种子激光器波长测量与PZT调制测试模块

该模块用于自动化执行种子激光器的波长测量与PZT调制测试，
支持信号源控制和HighFinesse波长计数据采集。
"""
from __future__ import annotations

import os
import time
import csv
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from typing import Callable, Optional, List, Tuple
from datetime import datetime
from pywinauto import Desktop
import pyvisa
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

from drivers import wlmData

# ===============  上位机控制（pywinauto）  ===============
try:
    from pywinauto.application import Application
    from pywinauto import timings
    PYW_AVAILABLE = True
except Exception:
    PYW_AVAILABLE = False

class WlmController:
    def __init__(self, window_title=".*Wavelength Meter.*", log_func=print):
        self.exe_path = exe_path
        self.window_title = window_title
        self.app = None
        self.win = None
        self.log = log_func

    def start_or_connect(self, timeout=15.0):
        try:
            # 优先连接已运行实例
            self.app = Application(backend="uia").connect(title_re=self.window_title, timeout=5)
            self.win = self.app.window(title_re=self.window_title)
            self.win.set_focus()
            self.log("[上位机] 已连接到运行中的窗口")
        except Exception:
            self.log("[上位机] 未找到运行实例，请打开…")
        timings.wait_until_passes(5, 0.5, lambda: self.win.exists() and self.win.is_visible())
        return True

class SignalGenerator:
    """
    信号源控制类
    
    通过LAN接口远程控制信号源（如Keysight/Agilent 33500B系列），
    支持正弦波、三角波等波形输出。
    """
    
    WAVEFORM_MAP = {
        'SIN': 'SINusoid',
        'RAMP': 'RAMP',
        'SQUARE': 'SQUare',
        'PULSE': 'PULSe',
        'NOISE': 'NOISe',
        'DC': 'DC',
    }
    
    def __init__(self, log_callback: Optional[Callable[[str], None]] = None) -> None:
        """
        初始化信号源控制类
        
        参数:
            log_callback (callable, optional): 日志回调函数
        """
        self.rm: Optional[pyvisa.ResourceManager] = None
        self.inst: Optional[pyvisa.Resource] = None
        self.log = log_callback or (lambda msg: None)
    
    def connect(self, ip_address: str, max_retries: int = 3, retry_interval: int = 2) -> bool:
        """
        连接信号源（带重试机制）
        
        参数:
            ip_address (str): 信号源的IP地址
            max_retries (int, optional): 最大重试次数
            retry_interval (int, optional): 重试间隔时间（秒）
            
        返回:
            bool: 连接成功返回True，失败返回False
        """
        for attempt in range(max_retries + 1):
            try:
                self.rm = pyvisa.ResourceManager()
                self.inst = self.rm.open_resource(f'TCPIP0::{ip_address}::INSTR')
                self.inst.timeout = 10000
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'
                idn = self.inst.query('*IDN?')
                self.log(f"[信号源] 已连接到信号源（第{attempt + 1}次尝试）")
                self.log(f"[信号源] 仪器标识: {idn.strip()}")
                return True
            except Exception as e:
                if attempt < max_retries:
                    self.log(f"[信号源] 第{attempt + 1}次连接失败：{e}，{retry_interval}秒后重试...")
                    time.sleep(retry_interval)
                else:
                    self.log(f"[信号源] 连接失败，共尝试{max_retries + 1}次")
                    return False
        return False
    
    def configure(self, waveform: str = "RAMP", freq: float = 0.1, 
                  volt: float = 10.0, offset: float = 5.0) -> bool:
        """
        配置信号源输出参数
        
        参数:
            waveform (str, optional): 波形类型，如"RAMP"表示三角波。默认"RAMP"。
            freq (float, optional): 输出频率，单位Hz。默认0.1 Hz。
            volt (float, optional): 输出幅值，单位Vpp。默认10.0 Vpp。
            offset (float, optional): 直流偏移，单位Vdc。默认5.0 Vdc。
            
        返回:
            bool: 配置成功返回True，失败返回False
        """
        if not self.inst:
            self.log("[信号源] 未连接到信号源")
            return False
        try:
            waveform_cmd = self.WAVEFORM_MAP.get(waveform.upper(), waveform.upper())
            self.inst.write(f":SOUR1:FUNC {waveform_cmd}")
            self.log(f"[信号源] 设置波形: {waveform.upper()}")
            
            self.inst.write(f":SOUR1:FREQ {freq}Hz")
            self.log(f"[信号源] 设置频率: {freq} Hz")
            
            self.inst.write(f":SOUR1:VOLT {volt}")
            self.log(f"[信号源] 设置幅值: {volt} Vpp")
            
            self.inst.write(f":SOUR1:VOLT:OFFS {offset}")
            self.log(f"[信号源] 设置偏置: {offset} Vdc")
            
            return True
        except Exception as e:
            self.log(f"[信号源] 配置失败: {e}")
            return False
    
    def set_output(self, on: bool = True) -> bool:
        """
        设置信号源输出开关状态
        
        参数:
            on (bool, optional): True表示打开输出，False表示关闭。默认True。
            
        返回:
            bool: 设置成功返回True，失败返回False
        """
        if not self.inst:
            self.log("[信号源] 未连接到信号源")
            return False
        try:
            if on:
                self.inst.write(":OUTP1 ON")
                self.log("[信号源] 打开输出")
            else:
                self.inst.write(":OUTP1 OFF")
                self.log("[信号源] 关闭输出")
            return True
        except Exception as e:
            self.log(f"[信号源] 设置输出状态失败: {e}")
            return False
    
    def close(self) -> None:
        """
        关闭信号源连接，释放VISA资源
        """
        if self.inst:
            try:
                self.inst.close()
                self.log("[信号源] 已关闭信号源连接")
            except Exception:
                pass
            self.inst = None
        if self.rm:
            try:
                self.rm.close()
            except Exception:
                pass
            self.rm = None


class WavemeterController:
    """
    波长计控制器类
    
    封装HighFinesse/Angstrom WS7-1波长计的wlmData DLL调用，
    提供连接、初始化、数据采集等功能。
    """
    
    def __init__(self, log_callback: Optional[Callable[[str], None]] = None) -> None:
        """
        初始化波长计控制器
        
        参数:
            log_callback (callable, optional): 日志回调函数
        """
        self.log = log_callback or (lambda msg: None)
        self.dll_loaded = False
        self.connected = False
        self.stop_flag = threading.Event()
    
    def load_dll(self, dll_path: str) -> bool:
        """
        加载wlmData.dll动态链接库
        
        参数:
            dll_path (str): DLL文件的完整路径
            
        返回:
            bool: 加载成功返回True，失败返回False
        """
        try:
            if not os.path.exists(dll_path):
                self.log(f"[波长计] DLL文件不存在: {dll_path}")
                return False
            
            wlmData.LoadDLL(dll_path)
            self.dll_loaded = True
            self.log(f"[波长计] 已加载DLL: {dll_path}")
            return True
        except Exception as e:
            self.log(f"[波长计] 加载DLL失败: {e}")
            return False
    
    def connect(self) -> bool:
        """
        连接波长计（调用Instantiate）
        
        返回:
            bool: 连接成功返回True，失败返回False
        """
        if not self.dll_loaded:
            self.log("[波长计] DLL未加载，请先加载DLL文件")
            return False
        try:
            result = wlmData.dll.Instantiate(0, 1, 0, 0)
            if result:
                self.connected = True
                self.log("[波长计] 已连接到波长计")
                return True
            else:
                self.log("[波长计] 连接波长计失败，请确保波长计软件已启动")
                return False
        except Exception as e:
            self.log(f"[波长计] 连接异常: {e}")
            return False
    
    def initialize(self, exposure_ms: int = 2) -> bool:
        """
        初始化波长计参数
        
        参数:
            exposure_ms (int, optional): 曝光时间（毫秒）。默认2ms。
            
        返回:
            bool: 初始化成功返回True，失败返回False
        """
        if not self.connected:
            self.log("[波长计] 未连接到波长计")
            return False
        try:
            result = wlmData.dll.SetExposureMode(False)
            if result >= 0:
                self.log("[波长计] 已关闭自动曝光模式")
            else:
                self.log("[波长计] 关闭自动曝光模式失败")
                return False
            
            result = wlmData.dll.SetExposureNum(1, 1, exposure_ms)
            if result >= 0:
                self.log(f"[波长计] 已设置Expo.1曝光时间为{exposure_ms}ms")
            else:
                self.log("[波长计] 设置曝光时间失败")
                return False
            
            return True
        except Exception as e:
            self.log(f"[波长计] 初始化异常: {e}")
            return False
    
    def get_interference_peak(self) -> Optional[int]:
        """
        获取干涉峰位置
        
        返回:
            int: 干涉峰位置值，失败返回None
        """
        if not self.connected:
            return None
        try:
            peak = wlmData.dll.GetMaxPeak(0)
            return peak
        except Exception:
            return None
    
    def get_frequency(self) -> Optional[float]:
        """
        获取当前频率值（THz单位）
        
        返回:
            float: 频率值（THz），失败返回None
        """
        if not self.connected:
            return None
        try:
            freq = wlmData.dll.GetFrequency(0.0)
            if freq > 0:
                return freq
            return None
        except Exception:
            return None
    
    def get_wavelength(self) -> Optional[float]:
        """
        获取当前波长值（nm单位）
        
        返回:
            float: 波长值（nm），失败返回None
        """
        if not self.connected:
            return None
        try:
            wl = wlmData.dll.GetWavelength(0.0)
            if wl > 0:
                return wl
            return None
        except Exception:
            return None
    
    def wait_for_next_event(self, timeout_ms: int = 1000) -> Tuple[bool, Optional[int], Optional[float]]:
        """
        阻塞等待下一个波长计事件（数据更新）
        
        参数:
            timeout_ms (int, optional): 超时时间（毫秒）。默认1000ms。
            
        返回:
            tuple: (是否成功, 事件模式, 频率值THz)
        """
        if not self.connected:
            return False, None, None
        
        if self.stop_flag.is_set():
            return False, None, None
        
        try:
            mode = ctypes.c_int32(0)
            int_val = ctypes.c_int32(0)
            dbl_val = ctypes.c_double(0.0)
            
            result = wlmData.dll.WaitForNextWLMEvent(
                ctypes.byref(mode),
                ctypes.byref(int_val),
                ctypes.byref(dbl_val)
            )
            
            if result > 0:
                return True, mode.value, dbl_val.value
            return False, None, None
        except Exception:
            return False, None, None
    
    def stop(self) -> None:
        """
        停止数据采集
        """
        self.stop_flag.set()
        self.log("[波长计] 数据采集已停止")
    
    def clear_stop(self) -> None:
        """
        清除停止标志
        """
        self.stop_flag.clear()


class WaveLengthTestGUI:
    """
    波长测量与PZT调制测试GUI控制类
    
    提供图形化界面用于配置测试参数、控制测试流程、显示运行日志。
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
            self.root.title("PZT调制")
            self.root.geometry("1200x820")
            self.root.resizable(True, True)
        else:
            self.root = parent
        
        self.params_1um = {
            '信号源IP': '192.168.7.11',
            '输出目录': r"C:\PTS\zhongzi\WaveLength",
            '频率(Hz)': '1',
            '幅值(mVpp)': '2',
            '偏置(mVdc)': '100',
            '变参测试参数': '0.1Hz, 1Vpp, 0.5Vdc\n0.1Hz, 2Vpp, 1Vdc\n0.1Hz, 5Vpp, 2.5Vdc\n0.1Hz, 10Vpp, 5Vdc',
            '调制时间(秒)': '100',
        }
        
        self.params_1_5um = {
            '信号源IP': '192.168.7.11',
            '输出目录': r"C:\PTS\zhongzi\WaveLength",
            '频率(Hz)': '1.5',
            '幅值(mVpp)': '4',
            '偏置(mVdc)': '100',
            '变参测试参数': '0.1Hz, 1Vpp, 0.5Vdc\n0.1Hz, 2Vpp, 1Vdc\n0.1Hz, 5Vpp, 2.5Vdc\n0.1Hz, 10Vpp, 5Vdc',
            '调制时间(秒)': '100',
        }
        
        self.test_type_var = tk.StringVar(value="1μm")
        self.params = self.params_1um.copy()
        
        self.worker: Optional[threading.Thread] = None
        self.signal_gen: Optional[SignalGenerator] = None
        self.wavemeter: Optional[WavemeterController] = None
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
        
        type_labelframe = tk.LabelFrame(left_frame, text='测试项选择', padx=10, pady=10)
        type_labelframe.pack(fill=tk.X, pady=8)
        type_frame = tk.Frame(type_labelframe)
        type_frame.pack(expand=True)
        tk.Label(type_frame, text="测试类型:").pack(side=tk.LEFT, padx=6)
        type_combo = tk.OptionMenu(type_frame, self.test_type_var, "1μm", "1.5μm", command=self._on_test_type_change)
        type_combo.pack(side=tk.LEFT, padx=6)
        
        conn_frame = tk.LabelFrame(left_frame, text='连接设置', padx=8, pady=8)
        conn_frame.pack(fill=tk.X, pady=4)
        conn_frame.columnconfigure(1, weight=1)
        
        init_frame = tk.LabelFrame(left_frame, text='初始测试参数', padx=8, pady=8)
        init_frame.pack(fill=tk.X, pady=4)
        init_frame.columnconfigure(1, weight=1)
        
        sweep_frame = tk.LabelFrame(left_frame, text='变参循环测试', padx=8, pady=8)
        sweep_frame.pack(fill=tk.BOTH, expand=True, pady=4)
        
        self.entries = {}
        self.init_frame = init_frame
        self.conn_frame = conn_frame
        
        conn_params = ['信号源IP', '输出目录']
        for i, k in enumerate(conn_params):
            label = tk.Label(conn_frame, text=k, width=10, anchor='e')
            label.grid(row=i, column=0, sticky='e', padx=2, pady=2)
            e = tk.Entry(conn_frame)
            e.insert(0, str(self.params[k]))
            e.grid(row=i, column=1, padx=2, pady=2, sticky='ew')
            self.entries[k] = e
        
        init_params = ['频率(Hz)', '幅值(mVpp)', '偏置(mVdc)']
        for i, k in enumerate(init_params):
            label = tk.Label(init_frame, text=k, width=10, anchor='e')
            label.grid(row=i, column=0, sticky='e', padx=2, pady=2)
            e = tk.Entry(init_frame)
            e.insert(0, str(self.params[k]))
            e.grid(row=i, column=1, padx=2, pady=2, sticky='ew')
            self.entries[k] = e
        
        sweep_label = tk.Label(sweep_frame, text='测试参数:')
        sweep_label.pack(anchor='w', padx=5, pady=2)
        
        self.sweep_text = tk.Text(sweep_frame, height=4, width=35, font=('Consolas', 10))
        self.sweep_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
        self.sweep_text.insert('1.0', self.params['变参测试参数'])
        
        time_frame = tk.Frame(sweep_frame)
        time_frame.pack(fill=tk.X, padx=5, pady=4)
        
        time_label = tk.Label(time_frame, text='调制时间(秒):', width=10, anchor='w')
        time_label.pack(side=tk.LEFT)
        
        self.time_entry = tk.Entry(time_frame, width=10)
        self.time_entry.insert(0, self.params['调制时间(秒)'])
        self.time_entry.pack(side=tk.LEFT, padx=4)
        self.entries['调制时间(秒)'] = self.time_entry
        
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=8)
        
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')
        
        self.start_btn = tk.Button(
            inner_btn_frame, text='开始测试', 
            bg="#4CAF50", fg="#FFFFFF", 
            command=self.start_test, 
            cursor="hand2",
            width=12
        )
        self.start_btn.pack(side='left', padx=6)
        
        self.stop_btn = tk.Button(
            inner_btn_frame, text='停止测试', 
            bg="#f44336", fg="#FFFFFF", 
            command=self.stop_test, 
            state=tk.DISABLED, 
            cursor="hand2",
            width=12
        )
        self.stop_btn.pack(side='left', padx=6)
        
        log_frame = tk.LabelFrame(right_frame, text='运行日志', padx=6, pady=6)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_box = tk.Text(log_frame, font=('Consolas', 10))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(self.log_box)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_box.yview)
    
    def log(self, msg: str) -> None:
        """
        线程安全地输出日志信息
        
        参数:
            msg (str): 日志消息
        """
        t = time.strftime('[%H:%M:%S]')
        self.root.after(0, lambda: self._safe_log_append(f"{t} {msg}\n"))
    
    def _safe_log_append(self, text: str) -> None:
        """
        线程安全地向日志文本框追加内容
        
        参数:
            text (str): 日志文本
        """
        self.log_box.insert(tk.END, text)
        self.log_box.see(tk.END)
    
    def capture_window(self, window_title_pattern: str, save_path: str, window_name: str = "窗口") -> bool:
        """
        加速版截图：使用 win32 后端和 Desktop 全局查找
        """
        try:
            if not PYW_AVAILABLE:
                self.log(f"[截图] Pywinauto不可用，跳过{window_name}截图")
                return False
            
            # 【优化1】改用 win32 后端，搜索速度极快
            desktop = Desktop(backend="win32")
            
            try:
                matched_windows = desktop.windows(title_re=window_title_pattern, visible_only=True)
                # 缩短等待时间，win32 找窗口基本是瞬间的
                if not matched_windows:
                    self.log(f"[截图] 找不到{window_name}")
                    return False
                
                # 强制选取列表中的第一个可用窗口，完美避开“匹配到多个”的报错
                win = matched_windows[0]
                
            except Exception as e:
                self.log(f"[截图] 查找{window_name}发生异常: {e}")
                return False
            
            if win.is_minimized():
                win.restore()
            
            # 【优化2】直接调用 set_focus，去掉 wait('active') 的 3 秒死等
            try:
                win.set_focus()
            except Exception:
                pass # 忽略焦点失败，继续强行截图
            
            # 截图并保存
            win.capture_as_image().save(save_path)
            self.log(f"[截图] {window_name}截图已保存: {os.path.basename(save_path)}")
            return True
            
        except Exception as e:
            self.log(f"[截图] {window_name}截图失败: {e}")
            return False
    
    def click_reset_button(self, window_title_pattern: str, button_name: str = "按钮", **kwargs) -> bool:
        """
        点击指定窗口中的控件（修正版：通过句柄绑定完整窗口对象）
        """
        try:
            if not PYW_AVAILABLE:
                self.log(f"[按钮] Pywinauto不可用，跳过点击{button_name}")
                return False
            
            desktop = Desktop(backend="win32")
            
            try:
                # 1. 瞬间找到所有匹配的底层 ElementInfo 对象
                matched_windows = desktop.windows(title_re=window_title_pattern, visible_only=True)
                if not matched_windows:
                    self.log(f"[按钮] 找不到目标窗口: {window_title_pattern}")
                    return False
                
                # 2. 拿到第一个匹配窗口的绝对物理句柄 (handle)
                target_handle = matched_windows[0].handle
                
                # 3. 通过句柄直接生成精准的 WindowSpecification 对象
                win = desktop.window(handle=target_handle)
                
            except Exception as e:
                self.log(f"[按钮] 查找目标窗口异常: {e}")
                return False
            
            # 4. 现在 win 是完整的对象，可以极速查找子控件了
            btn = win.child_window(**kwargs)
            
            # 获取控件坐标
            rect = btn.rectangle()
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            
            # 计算点击位置 (最右侧 "Reset now" 的中心)
            click_x = width - 40
            click_y = height // 2
            
            # 使用 click_input 进行物理点击
            btn.click_input(coords=(click_x, click_y))
            
            self.log(f"[按钮] 已点击{button_name} (相对坐标: X={click_x}, Y={click_y})")
            return True
        except Exception as e:
            self.log(f"[按钮] 点击{button_name}失败: {e}")
            return False

    def click_top_right_corner(self, window_title_pattern: str, control_name: str = "控件", **kwargs) -> bool:
        """
        点击指定控件的右上角（专用于 TChart 内部绘制的边缘小按钮）
        """
        try:
            if not PYW_AVAILABLE:
                self.log(f"[控件] Pywinauto不可用，跳过点击{control_name}")
                return False
            
            desktop = Desktop(backend="win32")
            
            try:
                matched_windows = desktop.windows(title_re=window_title_pattern, visible_only=True)
                if not matched_windows:
                    self.log(f"[控件] 找不到目标窗口: {window_title_pattern}")
                    return False
                target_handle = matched_windows[0].handle
                win = desktop.window(handle=target_handle)
            except Exception as e:
                self.log(f"[控件] 查找目标窗口异常: {e}")
                return False
            
            ctrl = win.child_window(**kwargs)
            
            # 获取控件的尺寸大小
            rect = ctrl.rectangle()
            width = rect.right - rect.left
            
            # 核心算法：计算右上角的坐标
            # X轴：整个宽度减去 12 像素（紧贴右边缘）
            # Y轴：从上边缘往下 12 像素
            # （12像素通常是这种16x16或24x24小图标的正中心）
            click_x = width - 12
            click_y = 12
            
            # 先移动鼠标过去，停顿0.1秒（让老软件有时间触发悬停效果）
            ctrl.move_mouse_input(coords=(click_x, click_y))
            import time
            time.sleep(0.1)
            
            # 使用 click_input 进行物理点击
            ctrl.click_input(coords=(click_x, click_y))
            
            self.log(f"[控件] 已点击{control_name} (相对坐标: X={click_x}, Y={click_y})")
            return True
        except Exception as e:
            self.log(f"[控件] 点击{control_name}失败: {e}")
            return False
    
    def read_and_save_pzt_range(self, window_title_pattern: str, output_dir: str) -> bool:
        """
        读取弹出面板中的 THz 数值，计算 Max - Min，并写入 PZT_range.csv
        """
        try:
            if not PYW_AVAILABLE:
                self.log("[数据] Pywinauto不可用，跳过数据读取")
                return False
                
            desktop = Desktop(backend="win32")
            matched_windows = desktop.windows(title_re=window_title_pattern, visible_only=True)
            if not matched_windows:
                self.log(f"[数据] 找不到目标窗口: {window_title_pattern}")
                return False
                
            win = desktop.window(handle=matched_windows[0].handle)
            
            # 获取窗口内所有的 TStaticText (静态文本) 控件
            static_texts = win.descendants(class_name="TStaticText")
            
            thz_values = []
            for ctrl in static_texts:
                text = ctrl.window_text()
                # 过滤出包含 "THz" 的文本
                if text and "THz" in text:
                    # 核心魔法：无视空格！把 "THz" 和空格全部替换为空，"193. 4578 THz" -> "193.4578"
                    clean_str = text.replace("THz", "").replace(" ", "").strip()
                    try:
                        # 尝试转换为浮点数（小数）
                        val = float(clean_str)
                        thz_values.append(val)
                    except ValueError:
                        # 如果是类似于 "Frequency [THz]" 这种不能转成数字的文字，直接忽略
                        pass
            
            # 如果成功抓取到了至少2个数值（Max 和 Min）
            if len(thz_values) >= 2:
                # 数学方法：列表中最大的数就是Max，最小的数就是Min
                max_val = max(thz_values)
                min_val = min(thz_values)
                pzt_range = (max_val - min_val) * 1000
                
                # 将结果保存到 PZT_range.csv
                import csv
                csv_path = os.path.join(output_dir, "PZT_range.csv")
                with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(['最小值(THz)', '最大值(THz)', 'PZT_Range(GHz)'])
                    writer.writerow([f"{min_val:.6f}", f"{max_val:.6f}", f"{pzt_range:.6f}"])
                
                self.log(f"[数据] 成功提取: Min={min_val:.4f}, Max={max_val:.4f}")
                self.log(f"[数据] 极差计算完成: {pzt_range:.1f} GHz，已存入 PZT_range.csv")
                return True
            else:
                self.log(f"[数据] 提取失败: 未能在弹窗中找到足够的 THz 数值。")
                return False
                
        except Exception as e:
            self.log(f"[数据] 读取数值并计算失败: {e}")
            return False

    def close_stats_panel(self, window_title_pattern: str) -> bool:
        """
        利用上下图表的几何边界对齐原理，计算并点击右上角的关闭箭头
        """
        try:
            if not PYW_AVAILABLE:
                self.log("[面板] Pywinauto不可用，跳过关闭弹窗")
                return False
                
            desktop = Desktop(backend="win32")
            matched_windows = desktop.windows(title_re=window_title_pattern, visible_only=True)
            if not matched_windows:
                self.log(f"[面板] 找不到目标窗口: {window_title_pattern}")
                return False
                
            win = desktop.window(handle=matched_windows[0].handle)
            
            # 1. 获取上方主图表（用于提取顶部的 Y 坐标）
            top_chart = win.child_window(class_name="TChart", found_index=1)
            top_rect = top_chart.rectangle()
            
            # 2. 获取下方功率图表（由于它未缩水，它的右边缘即为弹窗的最右边缘，用于提取 X 坐标）
            bottom_chart = win.child_window(class_name="TChart", found_index=0)
            bottom_rect = bottom_chart.rectangle()
            
            # 3. 十字交叉定位：下方最右侧向左退 12 像素，上方最顶端向下移 12 像素
            abs_x = bottom_rect.right - 12
            abs_y = top_rect.top + 12
            
            # 引入全局鼠标控制器，直接进行屏幕绝对坐标的物理点击
            import pywinauto.mouse as mouse
            import time
            
            mouse.move(coords=(abs_x, abs_y))
            time.sleep(0.1) # 停顿让软件响应悬浮状态
            mouse.click(button='left', coords=(abs_x, abs_y))
            
            self.log(f"[面板] 已点击关闭小窗 (屏幕坐标: X={abs_x}, Y={abs_y})")
            return True
            
        except Exception as e:
            self.log(f"[面板] 关闭小窗失败: {e}")
            return False
    
    def _on_test_type_change(self, event=None) -> None:
        """
        测试类型切换回调函数
        
        参数:
            event: 事件对象（OptionMenu回调）
        """
        if self.test_type_var.get() == "1μm":
            self.params = self.params_1um.copy()
        else:
            self.params = self.params_1_5um.copy()
        self._reset_entries()
    
    def _reset_entries(self) -> None:
        """
        刷新参数输入框内容
        """
        for widget in self.init_frame.winfo_children():
            widget.destroy()
        
        init_params = ['频率(Hz)', '幅值(mVpp)', '偏置(mVdc)']
        for i, k in enumerate(init_params):
            label = tk.Label(self.init_frame, text=k, width=10, anchor='e')
            label.grid(row=i, column=0, sticky='e', padx=2, pady=2)
            e = tk.Entry(self.init_frame)
            e.insert(0, str(self.params[k]))
            e.grid(row=i, column=1, padx=2, pady=2, sticky='ew')
            self.entries[k] = e
        
        self.sweep_text.delete('1.0', tk.END)
        self.sweep_text.insert('1.0', self.params['变参测试参数'])
        self.time_entry.delete(0, tk.END)
        self.time_entry.insert(0, self.params['调制时间(秒)'])
    
    def _save_params(self) -> None:
        """
        保存当前界面参数到self.params
        """
        for k, e in self.entries.items():
            self.params[k] = e.get()
        self.params['变参测试参数'] = self.sweep_text.get('1.0', tk.END).strip()
        
        if self.test_type_var.get() == "1μm":
            self.params_1um = self.params.copy()
        else:
            self.params_1_5um = self.params.copy()
    
    def _parse_sweep_params(self) -> List[dict]:
        """
        解析变参测试参数文本
        
        返回:
            list: 参数字典列表，每个字典包含freq, volt, offset
        """
        params_text = self.sweep_text.get('1.0', tk.END).strip()
        lines = [line.strip() for line in params_text.split('\n') if line.strip()]
        
        result = []
        for line in lines:
            try:
                parts = [p.strip() for p in line.split(',')]
                if len(parts) >= 3:
                    freq_str = parts[0].replace('Hz', '').replace('hz', '').strip()
                    volt_str = parts[1].replace('mVpp', '').replace('mvpp', '').replace('Vpp', '').strip()
                    offset_str = parts[2].replace('mVdc', '').replace('mvdc', '').replace('Vdc', '').strip()
                    
                    freq = float(freq_str)
                    volt = float(volt_str)
                    offset = float(offset_str)
                    
                    if 'mVpp' in parts[1].upper() or 'MVPP' in parts[1].upper():
                        volt = volt / 1000.0
                    if 'mVdc' in parts[2].upper() or 'MVDC' in parts[2].upper():
                        offset = offset / 1000.0
                    
                    result.append({
                        'freq': freq,
                        'volt': volt,
                        'offset': offset,
                        'original': line
                    })
            except Exception as e:
                self.log(f"[警告] 解析参数行失败: {line}, 错误: {e}")
        
        return result
    
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
    
    def start_test(self) -> None:
        """
        启动测试流程
        """
        if self.worker and self.worker.is_alive():
            messagebox.showinfo('提示', '测试已在进行中')
            return
        
        self._save_params()
        
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        self.stop_flag.clear()
        
        self.worker = threading.Thread(target=self.test_task, daemon=True)
        self.worker.start()
    
    def stop_test(self) -> None:
        """
        停止测试
        """
        self.stop_flag.set()
        
        self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
        
        if self.signal_gen:
            try:
                self.signal_gen.set_output(False)
            except Exception:
                pass
        
        if self.wavemeter:
            self.wavemeter.stop()
        
        self.log("[停止] 用户请求停止测试")
    
    def test_task(self) -> None:
        """
        后台测试任务线程
        """
        try:
            self.log("=" * 30)
            self.log("[开始] 波长测量与PZT调制测试")
            self.log("=" * 30)
            
            self.signal_gen = SignalGenerator(log_callback=self.log)
            self.wavemeter = WavemeterController(log_callback=self.log)
            
            self.log("\n[步骤1] 连接仪器...")
            
            dll_path = r'C:\Windows\System32\wlmData.dll'
            if not self.wavemeter.load_dll(dll_path):
                raise Exception("加载wlmData.dll失败")
            
            if not self.wavemeter.connect():
                raise Exception("连接波长计失败，请确保波长计软件已启动")
            
            sg_ip = self.params['信号源IP']
            if not self.signal_gen.connect(sg_ip):
                raise Exception("连接信号源失败")
            
            output_dir = self.params['输出目录']
            if not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir)
                    self.log(f"[文件] 已创建输出目录: {output_dir}")
                except Exception as e:
                    raise Exception(f"创建输出目录失败: {e}")
            
            self.log("\n[步骤2] 读取初始波长...")
            
            wavelength = self.wavemeter.get_wavelength()
            if wavelength:
                wavelength_csv_path = os.path.join(output_dir, "wavelength.csv")
                try:
                    with open(wavelength_csv_path, 'w', newline='', encoding='utf-8') as f:
                        writer = csv.writer(f)
                        writer.writerow([f"{wavelength:.0f}"])
                    self.log(f"[波长] 已读取初始波长: {wavelength:.0f} nm")
                    self.log(f"[文件] 波长数据已保存: wavelength.csv")
                except Exception as e:
                    self.log(f"[错误] 保存波长数据失败: {e}")
            else:
                self.log("[警告] 读取波长失败，未保存波长数据")
            
            self.log("\n[步骤3] 窗口一初始截图...")
            window1_title = r".*Wavelength Meter.*"
            window1_screenshot_path = os.path.join(output_dir, "波长.png")
            self.capture_window(window1_title, window1_screenshot_path, "窗口一")
            
            self.log("\n[步骤4] Reset...")
            reset_window_title1 = r".*WLM LongTerm graph.*"
            self.click_reset_button(
                reset_window_title1, 
                "Reset按钮", 
                class_name="TPanel", 
                found_index=1 
            )

            self.log("[步骤4.5] 略微延迟")
            if not self._interruptible_sleep(2):
                self.log("[停止] 用户中断测试")
                return

            self.log("\n[步骤5] 初始调制输出...")
            waveform = "RAMP"
            freq = float(self.params['频率(Hz)'])

            volt = float(self.params['幅值(mVpp)']) / 1000.0
            offset = float(self.params['偏置(mVdc)']) / 1000.0
            
            if not self.signal_gen.configure(waveform=waveform, freq=freq, volt=volt, offset=offset):
                raise Exception("配置信号源失败")
            
            if not self.signal_gen.set_output(True):
                raise Exception("打开信号源输出失败")

            self.log("[步骤5.5] 略微延迟")
            if not self._interruptible_sleep(2):
                self.log("[停止] 用户中断测试")
                return

            self.log("\n[步骤6] 窗口二初始信号截图...")
            window2_title = r".*WLM LongTerm graph.*"
            window2_screenshot_path = os.path.join(output_dir, "初始信号.png")
            self.capture_window(window2_title, window2_screenshot_path, "窗口二")
            
            # self.log("\n[步骤7] Reset...")
            # reset_window_title2 = r".*WLM LongTerm graph.*"
            # self.click_reset_button(
            #     reset_window_title2, 
            #     "Reset按钮", 
            #     class_name="TPanel", 
            #     found_index=1 
            # )

            if not self._interruptible_sleep(1):
                self.log("[停止] 用户中断测试")
                return

            self.log("\n[步骤8] 开始变参循环测试...")
            sweep_params = self._parse_sweep_params()
            if not sweep_params:
                self.log("[警告] 未解析到有效的变参测试参数，将仅执行初始参数测试")
                sweep_params = [{
                    'freq': freq,
                    'volt': volt,
                    'offset': offset,
                    'original': f"{freq}Hz, {volt*1000}mVpp, {offset*1000}mVdc"
                }]
            
            collect_time = float(self.params['调制时间(秒)'])
            
            for i, param in enumerate(sweep_params):
                if self.stop_flag.is_set():
                    self.log(f"[停止] 已停止测试，当前完成到第{i}组参数")
                    break
                
                self.log(f"\n{'='*20}")
                self.log(f"[测试] 第{i+1}/{len(sweep_params)}组参数: {param['original']}")
                self.log(f"{'='*20}")
                
                if not self.signal_gen.configure(
                    waveform=waveform,
                    freq=param['freq'],
                    volt=param['volt'],
                    offset=param['offset']
                ):
                    self.log(f"[错误] 配置第{i+1}组参数失败，跳过")
                    continue
                
                # self.log("[等待] 等待信号稳定（2秒）...")
                # time.sleep(2)
                reset_window_title = r".*WLM LongTerm graph.*"
                self.click_reset_button(
                    reset_window_title, 
                    "Reset按钮", 
                    class_name="TPanel", 
                    found_index=1 
                )
                
                self.log(f"[等待] 等待调制时间（{collect_time}秒）...")
                if not self._interruptible_sleep(collect_time):
                    self.log("[停止] 用户中断测试")
                    break
                
                window2_title = r".*WLM LongTerm graph.*"
                window2_screenshot_path = os.path.join(output_dir, f"{int(param['volt'])}v.png")
                self.capture_window(window2_title, window2_screenshot_path, "窗口二")
            
            if self.stop_flag.is_set():
                self.log("[停止] 测试已停止，跳过后续步骤")
                return
            
            self.log("\n[步骤9] 调制流程结束，计算调制范围...")
            # 使用 class_name="TChart" 和 found_index=1 定位上面的频率主图表
            chart_window_title = r".*WLM LongTerm graph.*"
            self.click_top_right_corner(
                chart_window_title, 
                "图表展开箭头", 
                class_name="TChart", 
                found_index=1 
            )

            # 新增：等待 1.5 秒，确保小弹窗完全展开渲染完毕
            self.log("[等待] 等待稳定...")
            if not self._interruptible_sleep(1.0):
                self.log("[停止] 用户中断测试")
                return
            
            self.log("\n[步骤10] 读取数据并计算调制范围...")
            self.read_and_save_pzt_range(chart_window_title, output_dir)

            if not self._interruptible_sleep(1.0):
                self.log("[停止] 用户中断测试")
                return
            
            self.log("\n[步骤11] 关闭弹窗...")
            self.close_stats_panel(chart_window_title)
            
            self.log("\n[完成] 测试流程结束")
            
        except Exception as e:
            self.log(f"[错误] 测试失败: {e}")
            self.root.after(0, lambda err=str(e): messagebox.showerror('错误', err))
        finally:
            if self.signal_gen:
                self.signal_gen.set_output(False)
                self.signal_gen.close()
            
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))


    def run(self) -> None:
        """
        以独立窗口模式运行
        """
        if self.root.winfo_exists():
            self.root.mainloop()


if __name__ == '__main__':
    gui = WaveLengthTestGUI()
    gui.run()
