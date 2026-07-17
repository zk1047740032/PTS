#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rin_FSV3004 — RIN (Relative Intensity Noise) 测量程序

三层架构（仪器 / 测试 / GUI）：
    FSV3004Instrument(VisaInstrument)   —— 纯 SCPI 通信，不涉及文件 IO / 数据处理 / GUI
    RinTest(BaseTestRunner)             —— RIN 测量流程 + 数据处理 + 可视化
    BackgroundNoiseTest(BaseTestRunner) —— 底噪 / 种子光测量 + 截图
    RinGUI(BaseTestGUI)                 —— 参数面板 + 按钮 + 日志 + 线程调度
"""
from __future__ import annotations  # 启用延迟类型注解
import os  # 文件和目录操作
import time  # 时间相关操作
import csv  # CSV文件读写
import threading  # 多线程支持
import traceback  # 异常追踪
from typing import List, Optional, Any, Dict  # 类型注解
import numpy as np  # 数值计算
import matplotlib  # 绘图库
matplotlib.use('TkAgg')  # 设置matplotlib后端为TkAgg
import matplotlib.pyplot as plt  # 绘图接口
from matplotlib.ticker import MaxNLocator  # 坐标轴刻度定位器
import tkinter as tk  # GUI框架
from tkinter import messagebox, filedialog, simpledialog  # Tkinter对话框
from PIL import Image, ImageTk  # 图片处理
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # matplotlib与tkinter集成

import sys, pathlib
# 确保能 import core（脚本独立运行时不以包形式组织）
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from core import (
    VisaInstrument,
    BaseTestGUI,
    BaseTestRunner,
    visa_address,
    clear_directory,
    append_row_csv,
    write_xy_csv,
    read_instrument_screenshot,
)

# default_logger 保留以兼容历史调用点；新代码直接用 print 或基类 log
def default_logger(msg: str):
    """
    默认日志函数，将消息打印到控制台

    参数:
        msg (str): 要打印的日志消息
    """
    print(msg)


# =========================================================================
# 第一层：仪器控制 —— 纯 SCPI 通信
# =========================================================================
class FSV3004Instrument(VisaInstrument):
    """
    FSV3004 频谱分析仪控制器（纯 SCPI 通信）

    不涉及文件 I/O、数据处理、GUI。同一台仪器可被 RinTest 和
    BackgroundNoiseTest 共用。

    主要属性：
    - instrument: self.inst 的业务别名（供内部 SCPI 方法沿用旧写法）
    - ip_address: 仪器 IP 地址
    """

    def __init__(self, log_func=default_logger):
        super().__init__(log_func=log_func, timeout_ms=60000)
        self.instrument = None  # self.inst 的业务别名
        self.ip_address = "192.168.7.10"

    # ------------------------------------------------------------------
    # 连接 / 释放
    # ------------------------------------------------------------------
    def connect(self, ip_address="192.168.7.10"):
        """
        连接频谱分析仪（VXI-11 协议，可读写二进制数据）

        参数:
            ip_address (str): 仪器IP地址，默认192.168.7.10

        返回:
            bool: 连接成功返回True，失败返回False
        """
        self.ip_address = ip_address
        ok = super().connect(visa_address(ip_address, kind="inst0"), idn=False)
        if ok:
            self.instrument = self.inst  # 别名指向同一资源对象
            self.log("成功连接到频谱分析仪（VXI-11）")
        return ok

    def close(self):
        """
        关闭仪器连接和资源（复用 VisaInstrument.close，并清理 instrument 别名）
        """
        super().close()
        self.instrument = None
        self.log("已关闭仪器连接")

    # ------------------------------------------------------------------
    # RIN 模式配置
    # ------------------------------------------------------------------
    def configure_rin_mode(self):
        """
        配置频谱分析仪为 RIN 测量模式
        设置扫描点数、追踪模式、检测器类型、输入耦合和功率单位
        """
        if not self.instrument:
            self.log("未连接到仪器")
            return
        self.instrument.write(":INST:SEL SAN")  # 切换为频谱分析模式
        # self.instrument.write(":CONF:SAN")  # 配置频谱分析仪
        self.instrument.write("SWE:POIN 2001")  # 设置采样点数
        self.instrument.write("DISPlay:TRACe1:MODE Average")  # Trace -> Trace1 -> Mode: Average
        self.instrument.write("DETector1:FUNCtion RMS")  # Trace -> Trace1 -> Detector Type: RMS
        self.instrument.write("INPut:COUPling DC")  # Amplitude -> Input Coupling: DC
        self.instrument.write("CALCulate:UNIT:POWer V")  # Amplitude -> Reference Level -> Unit: V
        self.log("仪器已配置（扫描点数：2001, 单位: V, 追踪模式: AVERage）")

    # ------------------------------------------------------------------
    # RIN 分段测量
    # ------------------------------------------------------------------
    def measure_rin_segment(self, start_freq, stop_freq, bandwidth, avg_count):
        """
        执行指定频段的 RIN 测量（SCPI + 二进制 trace 读取，不写文件）

        参数:
            start_freq (float): 起始频率(Hz)
            stop_freq (float): 终止频率(Hz)
            bandwidth (float): 带宽(Hz)
            avg_count (int): 平均次数

        返回:
            (freqs, amps): 频率数组和幅值数组；失败返回 (None, None)
        """
        if not self.instrument:
            self.log("未连接到仪器")
            return None, None
        self.instrument.write(f":BAND {bandwidth}")
        self.instrument.write(f":AVER:COUN {avg_count}")
        self.instrument.write(f":FREQ:STAR {start_freq} Hz")
        self.instrument.write(f":FREQ:STOP {stop_freq} Hz")
        self.instrument.query("*OPC?")

        self.instrument.write(":INIT:CONT OFF")
        self.instrument.write(":INIT")
        self.instrument.query("*OPC?")

        # 通过 VXI-11 直读 trace 数据（需先设为二进制格式）
        self.instrument.write("FORM:DATA REAL,32")
        amps = self.instrument.query_binary_values(":TRACe:DATA? TRACE1", datatype='f', is_big_endian=False)
        freqs = np.linspace(start_freq, stop_freq, len(amps))
        return freqs, amps

    # ------------------------------------------------------------------
    # 底噪 / 种子光模式配置
    # ------------------------------------------------------------------
    def configure_noise_mode(self, start_freq, stop_freq, bandwidth, avg_count):
        """
        配置频谱分析仪为底噪 / 种子光测量模式

        参数:
            start_freq (int): 起始频率(Hz)
            stop_freq (int): 终止频率(Hz)
            bandwidth (int): 带宽(Hz)
            avg_count (int): 平均次数
        """
        if not self.instrument:
            self.log("未连接到仪器")
            return
        self.instrument.write(":INST:SEL SA")
        self.instrument.write(":CONF:SAN")
        self.instrument.write("SWE:POIN 2001")
        self.instrument.write("UNIT:POW V")
        self.instrument.write("TRACE1:TYPE AVERage")
        self.instrument.write(f":BAND {bandwidth}")
        self.instrument.write(f":AVER:COUN {avg_count}")
        self.instrument.write(f":FREQ:STAR {start_freq} Hz")
        self.instrument.write(f":FREQ:STOP {stop_freq} Hz")
        self.instrument.query("*OPC?")

    def measure_noise_trace(self, start_freq, stop_freq):
        """
        启动单次扫描并读取 trace 数据（参数已在 configure_noise_mode 中设定）

        参数:
            start_freq (float): 起始频率(Hz) —— 用于构造频率轴
            stop_freq (float): 终止频率(Hz)

        返回:
            (freqs, amps): 频率数组和幅值数组
        """
        self.instrument.write(":INIT:CONT OFF")  # 关闭连续模式
        self.instrument.write(":INIT")
        self.instrument.query("*OPC?")

        self.instrument.write("FORM:DATA REAL,32")
        amps = self.instrument.query_binary_values(":TRACe:DATA? TRACE1", datatype='f', is_big_endian=False)
        freqs = np.linspace(start_freq, stop_freq, len(amps))
        return freqs, amps

    def capture_screenshot(self, instr_temp_path, pc_save_path):
        """
        通过 SCPI 把仪器屏幕截图读回 PC 本地

        参数:
            instr_temp_path (str): 仪器本地临时 png 路径
            pc_save_path (str): PC 本地保存路径（含文件名）
        """
        self.instrument.timeout = 60000
        read_instrument_screenshot(
            self.instrument,
            instr_temp_path,
            pc_save_path,
            log_func=self.log,
        )


# =========================================================================
# 第二层-1：RIN 测试逻辑
# =========================================================================
class RinTest(BaseTestRunner):
    """
    RIN 测量流程编排（继承 BaseTestRunner，填 hook）

    持有 FSV3004Instrument 实例。包含全部数据处理、RIN 换算、
    matplotlib 可视化、驰豫振荡峰检测。

    主要属性：
    - instrument: FSV3004Instrument 实例
    - dc_value / amplification: RIN 换算参数
    - file_paths: 6 段测量数据文件路径
    - dx/dy/ddx/ddy: 原始 / 合并后的频率-RIN 数据
    - RIN_power: 功率积分结果
    - save_path / ui_root / stop_window: GUI 注入属性

    主要方法：
    - run(): 覆写基类模板（匹配原 TestRunner.run_rin 行为）
    - read_data_from_csv / process_files / compute_rin_power: 数据处理
    - visualize_data: matplotlib 可视化 + 弹窗
    """

    def __init__(self, gui=None, log_func=None):
        super().__init__(gui=gui, log_func=log_func)
        self.instrument: Optional[FSV3004Instrument] = None
        self.dc_value = 1.20  # 默认DC值
        self.amplification = 14
        self.file_paths = [
            'C:\\PTS\\zhongzi\\Rin\\FSV3004\\Rin_1.DAT',
            'C:\\PTS\\zhongzi\\Rin\\FSV3004\\Rin_2.DAT',
            'C:\\PTS\\zhongzi\\Rin\\FSV3004\\Rin_3.DAT',
            'C:\\PTS\\zhongzi\\Rin\\FSV3004\\Rin_4.DAT',
            'C:\\PTS\\zhongzi\\Rin\\FSV3004\\Rin_5.DAT',
            'C:\\PTS\\zhongzi\\Rin\\FSV3004\\Rin_6.DAT',
        ]
        self.dx = []
        self.dy = []
        self.ddx = []
        self.ddy = []
        self.RIN_power = []
        self.save_path = None
        self.ui_root = None
        self.stop_window = None
        self.ip_address = "192.168.7.10"

    # ------------------------------------------------------------------
    # 模板方法覆写（匹配原 TestRunner.run_rin 错误处理与流程）
    # ------------------------------------------------------------------
    auto_close_delay_ms = 0  # 不由 runner 自动关窗，交给 GUI 的 finally 块处理

    def run(self):
        """
        执行 RIN 测量流程。
        与 TestRunner.run_rin 的 try/except/finally 结构完全一致，
        hook 仍按 prepare → connect → configure → measure_loop → finalize 编排。
        """
        try:
            try:
                self.prepare()
            except Exception as e:
                self.log(f"[错误] 清理文件夹时出错: {e}")

            if self.connect():
                self.configure()
                self.measure_loop()
                # 关闭连接（匹配原 run_rin 在 measurement_loop 后立即 close 的行为）
                try:
                    self.instrument.close()
                except Exception:
                    pass
            else:
                self.log("[测试] 无法连接到仪器，RIN 测试终止")

            if self.stop_flag.is_set():
                # 如果停止，更新 UI stop_window
                def _notify_stopped():
                    if self.stop_window and self.stop_window.winfo_exists():
                        self.stop_window.destroy()
                        self.stop_window = None
                try:
                    self.ui_root.after(0, _notify_stopped)
                except Exception:
                    pass
                return

            # 原脚本会在这里处理文件与可视化
            if self.stop_window and self.stop_window.winfo_exists():
                try:
                    self.stop_window.destroy()
                except Exception:
                    pass
                self.stop_window = None

            self.finalize()
            self.log("程序执行完毕")
        except Exception as e:
            self.log(f"[Runner Exception] {e}\n{traceback.format_exc()}")
        finally:
            self.cleanup()

    # ------------------------------------------------------------------
    # Hook 实现
    # ------------------------------------------------------------------
    def prepare(self):
        """清空电脑本地数据文件夹"""
        self.log("[初始化] 正在清空电脑本地数据文件夹...")
        local_dir = r"C:\PTS\zhongzi\Rin\FSV3004"
        clear_directory(local_dir, log_func=self.log)
        self.log("[初始化] 文件夹清理完成。")

    def connect(self) -> bool:
        """连接频谱分析仪"""
        return self.instrument.connect(ip_address=self.ip_address)

    def configure(self):
        """配置仪器为 RIN 测量模式"""
        self.instrument.configure_rin_mode()

    def measure_loop(self):
        """6 段测量循环：每段调仪器读 trace → CSV 落盘"""
        measurement_params = [
            (10, 100, 5, 20, "Rin_1.DAT"),
            (100, 1000, 5, 20, "Rin_2.DAT"),
            (1000, 10000, 30, 20, "Rin_3.DAT"),
            (10000, 100000, 30, 20, "Rin_4.DAT"),
            (100000, 1000000, 30, 20, "Rin_5.DAT"),
            (1000000, 10000000, 30, 20, "Rin_6.DAT"),
        ]
        dest_dir = r"C:\PTS\zhongzi\Rin\FSV3004"
        for start, stop, bw, avg, fname in measurement_params:
            if self.stop_flag.is_set():
                self.log("[测试] RIN 测试已被终止")
                break
            self.log(f"[测试] 正在测量: {start}Hz - {stop}Hz, 带宽: {bw}Hz")
            freqs, amps = self.instrument.measure_rin_segment(start, stop, bw, avg)
            if freqs is None:
                self.log(f"测量失败: {start}Hz - {stop}Hz")
                continue
            os.makedirs(dest_dir, exist_ok=True)
            filepath = os.path.join(dest_dir, fname)
            with open(filepath, 'w', newline='') as f:
                writer = csv.writer(f)
                for freq, amp in zip(freqs, amps):
                    writer.writerow([freq, amp])
            self.log(f"数据已保存到: {filepath}")

    def finalize(self):
        """数据处理 → 可视化"""
        self.log("正在处理数据...")
        self.process_files()
        self.log("正在显示可视化结果...")
        self.visualize_data()

    def cleanup(self):
        """仪器已在 measure_loop 后关闭，此处兜底"""
        if self.instrument and self.instrument.is_connected:
            try:
                self.instrument.close()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # 数据处理（从 RinAnalyzer 逐行保留）
    # ------------------------------------------------------------------
    def read_data_from_csv(self, file_path):
        """
        从CSV文件读取频率和幅度数据

        参数:
            file_path (str): CSV文件路径

        返回:
            bool: 读取成功返回True，失败返回False
        """
        try:
            with open(file_path, 'r') as f:
                sample = f.read(1024)
                f.seek(0)
                dialect = csv.Sniffer().sniff(sample)
                reader = csv.reader(f, dialect)

                file_dx, file_dy = [], []
                for row in reader:
                    if len(row) >= 2:
                        try:
                            # 兼容逗号作为小数分隔符
                            x = float(row[0].replace(',', '.'))
                            y = float(row[1].replace(',', '.'))
                            file_dx.append(x)
                            file_dy.append(y)
                        except ValueError:
                            continue
            if len(file_dy) < 2001:
                self.log(f"警告: 数据点数非2001，实际 {len(file_dy)}")
                return False
            self.dx.append(file_dx)
            self.dy.append(file_dy)
            return True
        except Exception as e:
            self.log(f"读取文件失败 {file_path}: {e}")
            return False

    def process_files(self):
        """
        处理所有测量数据文件
        读取 CSV 数据，计算 RIN 值，合并数据并计算功率积分
        """
        # 检测文件夹是否为空
        data_dir = r"C:\PTS\zhongzi\Rin\FSV3004"
        if os.path.exists(data_dir):
            files = [f for f in os.listdir(data_dir) if os.path.isfile(os.path.join(data_dir, f))]
            if not files:
                self.log(f"[警告] 数据文件夹为空：{data_dir}")
            else:
                self.log(f"数据文件夹包含 {len(files)} 个文件：{', '.join(files[:5])}{'...' if len(files) > 5 else ''}")
        else:
            self.log(f"[警告] 数据文件夹不存在：{data_dir}")

        self.dx = []
        self.dy = []
        self.ddx = []
        self.ddy = []

        for file_path in self.file_paths:
            # 数据已通过 SCPI 直写本地，直接读取即可
            if not (os.path.exists(file_path) and os.path.getsize(file_path) > 0):
                self.log(f"文件不存在: {file_path}")
                self.dx.append([])
                self.dy.append([])
                continue

            read_ok = False
            for attempt in range(3):
                try:
                    if self.read_data_from_csv(file_path):
                        self.log(f"成功读取: {file_path}")
                        read_ok = True
                        break
                except Exception as e:
                    self.log(f"读取异常（尝试{attempt+1}）: {file_path} -> {e}")
                time.sleep(0.5)

            if not read_ok:
                self.log(f"最终读取失败: {file_path}")
                self.dx.append([])
                self.dy.append([])

        rows_per_file = 2001
        for j in range(len(self.dx)):
            if not self.dx[j]:
                self.log(f"文件{j}数据为空，跳过处理")
                continue

            if len(self.dy[j]) != rows_per_file:
                self.log(f"警告: 文件{j}数据点不足，期望{rows_per_file}个，实际 {len(self.dy[j])}")

            for i in range(min(rows_per_file, len(self.dx[j]))):
                self.ddx.append(self.dx[j][i])
                scale_factor = np.sqrt(5) if j < 2 else np.sqrt(30)
                v_noise = self.dy[j][i]
                if v_noise <= 0:
                    self.ddy.append(float('-inf'))  # 无效数据填充为 -inf
                else:
                    rin_value = 20 * np.log10(v_noise / (self.dc_value * self.amplification * scale_factor))
                    self.ddy.append(rin_value)

        if self.ddx and self.ddy:
            self.RIN_power = self.compute_rin_power(self.ddx, self.ddy)
        else:
            self.log("错误: 无有效数据可处理")
            self.RIN_power = []

    def compute_rin_power(self, x, y):
        """
        计算RIN功率积分

        参数:
            x (list): 频率数据列表
            y (list): RIN值数据列表

        返回:
            list: 功率积分结果列表
        """
        power = []
        segment_length = 6
        for k in range(1, len(x) // segment_length + 1):
            sub_x = x[:k*segment_length]
            sub_y_exp = [np.power(10, val / 10.0) if np.isfinite(val) else 0 for val in y[:k*segment_length]]
            integral = sum((sub_x[i] - sub_x[i-1]) * (sub_y_exp[i] + sub_y_exp[i-1]) / 2.0
                           for i in range(1, len(sub_x)))
            power.append(np.sqrt(integral))
        return power

    # ------------------------------------------------------------------
    # 可视化（从 RinAnalyzer.visualize_data 逐行保留）
    # ------------------------------------------------------------------
    def visualize_data(self):
        """
        可视化RIN数据和积分结果
        创建包含RIN曲线和积分曲线的图表窗口
        支持手动保存和自动保存功能
        """
        if not self.ddx or not self.ddy or not self.RIN_power:
            self.log("没有可视化的数据")
            return
        root = tk.Toplevel()
        root.title("测Rin数据可视化")
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8), gridspec_kw={'height_ratios': [3, 2]})

        def save_figure():
            file_path = filedialog.asksaveasfilename(defaultextension=".png",filetypes=[("PNG files", "*.png"),("JPEG files", "*.jpg"),("All files", "*.*")])
            if file_path:
                fig.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("成功", f"图像已保存到 {file_path}")
        # 创建一个框架用于放置顶部按钮
        top_frame = tk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=10)
        # 在框架中间放置保存按钮
        tk.Button(top_frame, text="保存", command=save_figure, font=('SimHei', 20), cursor="hand2").pack(side=tk.TOP)

        """图1: RIN曲线"""
        ax1.plot(self.ddx, self.ddy, color="#085cab", linewidth=2) # 曲线
        ax1.set_xscale('log')
        ax1.margins(x=0) # 边距
        ax1.tick_params(axis='both', which='major', labelsize=20, pad=5, length=12, width=3, direction='in') # 刻度线
        #ax1.set_xlabel('Frequency(Hz)', fontsize=35, fontweight='bold', fontstyle='normal') # 子图1不要x轴单位
        ax1.set_ylabel('RIN (dBc/Hz)', fontsize=18, fontweight='bold', fontstyle='normal') # 子图1y轴单位
        ax1.grid(True, which='both', lw = 2, linestyle='--', alpha=1) # 网格线
        # 边框加粗
        for spine in ax1.spines.values():
            spine.set_linewidth(2.5)
        # 刻度坐标
        for label in ax1.get_xticklabels() + ax1.get_yticklabels():
            label.set_fontname('Times New Roman')
            label.set_fontsize(20)
            label.set_fontweight('bold')
        # 设置x轴范围
        finite_ddy = [v for v in self.ddy if np.isfinite(v)]
        if finite_ddy:
            ax1.set_ylim(np.floor(min(finite_ddy)/10)*10, np.ceil(max(finite_ddy)/10)*10)
        ax1.yaxis.set_major_locator(MaxNLocator(nbins=6, integer=True))
        # 设置x轴主刻度
        adjusted_power = [p * 100 for p in self.RIN_power]

        """图2: RMS积分曲线"""
        plot_len = len(adjusted_power)
        ax2.plot(self.ddx[::6][:plot_len], adjusted_power, color="#085cab", linewidth=2)
        ax2.set_xscale('log')
        ax2.margins(x=0)
       # y轴只显示最大值、最小值和中间值
        y2_min, y2_max = np.min(adjusted_power), np.max(adjusted_power)
        if np.isclose(y2_min, y2_max):
            y2_min = y2_min - 1
            y2_max = y2_max + 1
        y2_mid = (y2_min + y2_max) / 2
        ax2.set_yticks([y2_min, y2_mid, y2_max])
        ax2.set_yticklabels([f"{y2_min:.3f}%", f"{y2_mid:.3f}%", f"{y2_max:.3f}%"])

        # 显示效果设置
        ax2.tick_params(axis='both', which='major', labelsize=15, pad=5, length=12, width=3, direction='in') # 刻度设置
        ax2.set_xlabel('Frequency(Hz)', fontsize=18, fontweight='bold', fontstyle='normal')
        ax2.set_ylabel('Integrated RMS', fontsize=18, fontweight='bold', fontstyle='normal') # %显示在y轴刻度上
        ax2.grid(True, which='both', lw=2, linestyle='--', alpha=1)
        # 边框加粗
        for spine in ax2.spines.values():
            spine.set_linewidth(2.5)
        # 设置坐标刻度字体为 Times New Roman，字体加粗
        for label in ax2.get_xticklabels() + ax2.get_yticklabels():
            label.set_fontname('Times New Roman')
            label.set_fontsize(20)
            label.set_fontweight('bold')

        plt.tight_layout()
        plt.subplots_adjust(hspace=0.15)  # 两图间距

        canvas = FigureCanvasTkAgg(fig, master=root)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 自动保存图像为 Rin.png：优先使用 GUI 提供的保存目录（save_path），否则回退到第一个 data 路径所在目录或 cwd
        try:
            save_dir = None
            # 优先使用外部设置的 save_path（由 GUI 在启动时注入到对象上）
            if hasattr(self, 'save_path') and self.save_path:
                save_dir = str(self.save_path)
            # 否则使用第一个 data 路径所在目录
            if not save_dir and self.file_paths and len(self.file_paths) > 0:
                save_dir = os.path.dirname(self.file_paths[0])
            if not save_dir:
                save_dir = os.getcwd()
            os.makedirs(save_dir, exist_ok=True)
            auto_path = os.path.join(save_dir, "Rin.png")

            # 确保 Tk 布局完成后再读取 widget 的像素尺寸，以生成与弹窗中显示一致的图片
            try:
                root.update_idletasks()
                widget = canvas.get_tk_widget()
                w_px = widget.winfo_width()
                h_px = widget.winfo_height()
            except Exception:
                w_px = h_px = 0

            # 使用固定的 DPI=300 进行保存，与手动保存保持一致
            try:
                fig.savefig(auto_path, dpi=300, bbox_inches='tight')
                self.log(f"[保存] 自动保存Rin图片: {auto_path}")

                # 保存图二y轴最大值到CSV文件
                try:
                    max_value_path = os.path.join(save_dir, "rin_figure2_max.csv")
                    with open(max_value_path, "w", newline="", encoding="utf-8") as f:
                        writer = csv.writer(f)
                        writer.writerow([f"{y2_max:.3f}"])
                    self.log(f"[保存] 自动保存图二y轴最大值: {max_value_path}")
                except Exception as e:
                    self.log(f"[保存] 自动保存图二y轴最大值失败: {e}")

            except Exception as e:
                self.log(f"[保存] 自动保存Rin图片失败: {e}")
        except Exception as e:
            self.log(f"[保存] 生成自动保存路径失败: {e}")

        target_xs = [1000, 10000, 100000, 1000000]

        """驰誉振荡峰检测"""
        # 1. 修改检测范围为 1e5 到 3e6 Hz
        relax_start = 1e5
        relax_stop = 3e6

        freqs = np.array(self.ddx)
        ys = np.array(self.ddy)

        # 筛选出范围内的有效数据
        mask = (freqs >= relax_start) & (freqs <= relax_stop) & np.isfinite(ys)

        highest_rin = float('nan')
        peak_freq = float('nan')

        if np.any(mask):
            masked_freqs = freqs[mask]
            masked_ys = ys[mask]

            # 2. 寻找"凸起"部分（局部最大值）
            # 只有数据点足够时才进行形态判断
            if len(masked_ys) >= 3:
                # 比较每个点是否比左右邻居都大
                center_vals = masked_ys[1:-1]
                left_vals = masked_ys[:-2]
                right_vals = masked_ys[2:]

                # 找到所有局部峰值的布尔掩码
                is_peak = (center_vals > left_vals) & (center_vals > right_vals)

                if np.any(is_peak):
                    # 获取局部峰值在 masked_ys 中的索引 (+1是因为切片从1开始)
                    peak_indices = np.where(is_peak)[0] + 1

                    # 在所有局部峰值中，找出 RIN 值最大的那个
                    best_peak_idx = peak_indices[np.argmax(masked_ys[peak_indices])]
                    highest_rin = float(masked_ys[best_peak_idx])
                    peak_freq = float(masked_freqs[best_peak_idx])
                else:
                    # 如果区间内没有"凸起"（例如单调上升或下降），回退到取最大值
                    max_idx = np.argmax(masked_ys)
                    highest_rin = float(masked_ys[max_idx])
                    peak_freq = float(masked_freqs[max_idx])
            else:
                # 点数太少无法判断形态，直接取最大值
                max_idx = np.argmax(masked_ys)
                highest_rin = float(masked_ys[max_idx])
                peak_freq = float(masked_freqs[max_idx])

        # 拼接弹窗文本（显示峰值及若干指定频点的值）
        result_text = f"驰豫振荡峰 ({int(relax_start):d} - {int(relax_stop):d} Hz): {highest_rin:.3f} dBc/Hz @ {peak_freq:.0f} Hz\n\n"
        for tx in target_xs:
            if len(self.ddx) == 0:
                continue
            idx = np.argmin(np.abs(np.array(self.ddx) - tx))
            x_val = self.ddx[idx]
            y_val = self.ddy[idx]
            # 处理无效值显示
            if not np.isfinite(y_val):
                result_text += f"x={x_val:.0f} Hz 时, y=无效数据\n"
            else:
                result_text += f"x={x_val:.0f} Hz 时, y={y_val:.3f} dBc/Hz\n"

        # 弹窗显示结果
        messagebox.showinfo("指定点的RIN值", result_text, parent=root)


# =========================================================================
# 第二层-2：底噪 / 种子光测试逻辑
# =========================================================================
class BackgroundNoiseTest(BaseTestRunner):
    """
    底噪 / 种子光测量流程编排（继承 BaseTestRunner，填 hook）

    持有 FSV3004Instrument 实例。负责配置仪器 → 单次测量 → 截图 → 显示。

    主要属性：
    - instrument: FSV3004Instrument 实例
    - is_seedlight: True=种子光, False=底噪（控制文件名与窗口标题）
    """

    def __init__(self, gui=None, log_func=None):
        super().__init__(gui=gui, log_func=log_func)
        self.instrument: Optional[FSV3004Instrument] = None
        self.ip_address = "192.168.7.10"
        self.start_freq = 10
        self.stop_freq = 100_000_000
        self.bandwidth = 30
        self.avg_count = 1
        self.screenshot_name = "BackgroundNoise_Screen.png"
        self.dat_filename = "BackgroundNoise.DAT"
        self.is_seedlight = False
        self.dest_dir = r"C:\PTS\zhongzi\Rin\FSV3004"

    # ------------------------------------------------------------------
    # 模板方法覆写（匹配原 TestRunner.run_background 结构）
    # ------------------------------------------------------------------
    auto_close_delay_ms = 0

    def run(self):
        """执行底噪 / 种子光测量"""
        try:
            if self.connect():
                self.configure()
                self.measure_loop()
                self.finalize()
                try:
                    self.instrument.close()
                except Exception:
                    pass
            else:
                self.log("错误", "无法连接到仪器")
        except Exception as e:
            self.log(f"[Background Exception] {e}\n{traceback.format_exc()}")
        finally:
            self.cleanup()

    # ------------------------------------------------------------------
    # Hook 实现
    # ------------------------------------------------------------------
    def connect(self) -> bool:
        """连接频谱分析仪"""
        return self.instrument.connect(ip_address=self.ip_address)

    def configure(self):
        """配置仪器为底噪 / 种子光测量模式"""
        self.instrument.configure_noise_mode(
            self.start_freq, self.stop_freq, self.bandwidth, self.avg_count
        )

    def measure_loop(self):
        """单次测量：读 trace → 保存 DAT → 截图"""
        freqs, amps = self.instrument.measure_noise_trace(self.start_freq, self.stop_freq)

        # 根据类型显示不同的日志
        if self.is_seedlight:
            self.log("种子光测量完成，开始截图和保存数据...")
        else:
            self.log("底噪测量完成，开始截图和保存数据...")

        # 保存 trace 数据
        os.makedirs(self.dest_dir, exist_ok=True)
        dat_path = os.path.join(self.dest_dir, self.dat_filename)
        write_xy_csv(dat_path, freqs, amps)
        self.log(f"数据已保存到: {dat_path}")

        # 截图
        local_png = os.path.join(self.dest_dir, self.screenshot_name)
        self.instrument.capture_screenshot(
            r"C:\PTS\Rin\_temp.png",
            local_png,
        )

        # 显示截图
        self._show_screenshot(self.is_seedlight)

    def finalize(self):
        """截图显示已在 measure_loop 中完成，此处为空"""
        pass

    def cleanup(self):
        """兜底关闭仪器"""
        if self.instrument and self.instrument.is_connected:
            try:
                self.instrument.close()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # 截图显示（从 BackgroundNoiseAnalyzer.show_screenshot 逐行保留）
    # ------------------------------------------------------------------
    def _show_screenshot(self, is_seedlight=False):
        """
        显示仪器截图和数据文件

        参数:
            is_seedlight (bool): 是否为种子光测量，影响窗口标题
        """
        local_img_path = os.path.join(self.dest_dir, self.screenshot_name)
        # 数据文件路径
        local_dat_path = os.path.join(self.dest_dir, self.dat_filename)

        win = tk.Toplevel()  # 不要用Tk()
        # 根据类型设置窗口标题
        if is_seedlight:
            win.title("种子光仪器截图")
        else:
            win.title("底噪仪器截图")

        # 添加顶部框架和居中保存按钮
        top_frame = tk.Frame(win)
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=10)

        def save_image():
            save_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg"), ("All files", "*.*")]
            )
            if save_path:
                try:
                    img = Image.open(local_img_path)
                    img.save(save_path)
                    messagebox.showinfo("成功", f"图像已保存到 {save_path}")
                except Exception as e:
                    self.log("保存失败", f"保存图片时出错: {e}")

        def save_data():
            if not os.path.exists(local_dat_path):
                self.log("错误", "数据文件不存在")
                return

            # 根据类型设置默认文件名
            default_filename = "SeedLight.dat" if is_seedlight else "BackgroundNoise.dat"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".dat",
                filetypes=[("DAT files", "*.dat"), ("All files", "*.*")],
                initialfile=default_filename
            )
            if save_path:
                try:
                    # 复制文件
                    import shutil
                    shutil.copy2(local_dat_path, save_path)
                    messagebox.showinfo("成功", f"数据已保存到 {save_path}")
                except Exception as e:
                    self.log("保存失败", f"保存数据时出错: {e}")

        # 创建按钮框架来放置两个按钮
        btn_frame = tk.Frame(top_frame)
        btn_frame.pack(side=tk.TOP)

        # 保存图片按钮
        tk.Button(btn_frame, text="保存图片", command=save_image, font=('SimHei', 16), cursor="hand2").pack(side=tk.LEFT, padx=10)
        # 保存数据按钮
        tk.Button(btn_frame, text="保存数据", command=save_data, font=('SimHei', 16), cursor="hand2").pack(side=tk.LEFT, padx=10)

        # 截图已通过 SCPI 直写本地，无需网络同步等待
        if not os.path.exists(local_img_path) or os.path.getsize(local_img_path) == 0:
            msg = f"图片未找到：{local_img_path}"
            tk.Label(win, text=msg, fg="red", wraplength=700, justify='left').pack(padx=8, pady=8)
            self.log(f"[显示] {msg}")
        else:
            try:
                img = Image.open(local_img_path)
                img = img.resize((800, 600))
                photo = ImageTk.PhotoImage(img, master=win)
                label = tk.Label(win, image=photo)
                label.image = photo  # 防止被回收
                label.pack()
            except Exception as e:
                tk.Label(win, text=f"图片加载失败: {e}", fg="red").pack()


# =========================================================================
# 第三层：GUI
# =========================================================================
class RinGUI(BaseTestGUI):
    """
    RIN测量图形用户界面类

    继承 BaseTestGUI：复用窗口构造、线程安全日志(log)、后台线程、
    文件/目录浏览对话框、auto_close、run 等通用能力，子类只保留自身布局与业务按钮。

    功能：
    - 提供参数设置界面（IP地址、DC值、保存路径）
    - 控制测量任务（测RIN、测底噪、种子光）
    - 显示运行日志
    - 支持独立运行和集成模式
    - 提供文件重命名功能

    主要属性：
    - root: 主窗口或父容器（来自基类）
    - stop_flag: 停止事件（来自基类）
    - params: 参数字典
    - entries: 输入框字典
    - worker_thread: 工作线程
    - running_task: 当前运行任务标识
    - _active_test: 当前运行的测试实例（用于传递停止信号）

    主要方法：
    - create_widgets(): 创建界面组件
    - start_rin/start_background/start_seedlight: 启动不同测量任务
    - connect_instrument: 测试仪器连接
    - stop_running: 停止当前任务
    - rename_files: 重命名测量文件
    """
    def __init__(self, parent=None):
        super().__init__(parent, title="Rin_FSV3004 - 独立模式",
                         geometry="1170x630", icon="PreciLasers.ico")
        # 独立模式下窗口居中（集成模式由父容器决定布局，不居中）
        if parent is None:
            self.set_center(1170, 330)

        # 默认参数（保留原脚本默认路径/IP）
        self.params = {
            "osa_ip": "192.168.7.20",
            #"osa_port": 5025,
            "save_path": r"C:\PTS\zhongzi\Rin\FSV3004",
            "dc_initial": 2.40
        }
        self.entries: Dict[str, tk.Entry] = {}
        self.worker_thread: Optional[threading.Thread] = None
        self.running_task: Optional[str] = None
        self._active_test: Optional[BaseTestRunner] = None  # 当前运行的测试实例，供停止按钮使用

        self.create_widgets()

    def set_center(self, width: int, height: int):
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        posx = (screenwidth - width) // 2
        posy = (screenheight - height) // 2
        self.root.geometry(f'{width}x{height}+{posx}+{posy}')

    def create_widgets(self):
        """
        创建GUI界面组件
        包括参数设置区、按钮区域和日志显示区
        """
        # 创建主容器，使用grid布局
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 左侧容器 - 用于容纳参数设置框和按钮，形成一个整体
        left_frame = tk.Frame(main_container)
        left_frame.grid(row=0, column=0, sticky="n", padx=(0, 5))

        # --- 参数设置 --- (左侧容器内) - 固定大小，不随窗口拉伸
        param_frame = tk.LabelFrame(left_frame, text="参数设置", padx=10, pady=10)
        param_frame.pack(fill=tk.X, padx=0, pady=0)

        # IP / port / 保存路径
        self._add_param_entry(param_frame, "osa_ip", "IP地址:", self.params["osa_ip"], row=0)
        self._add_param_entry(param_frame, "dc_value", "DC值:", "2.4", row=1)
        #self._add_param_entry(param_frame, "osa_port", "端口:", str(self.params["osa_port"]), row=1)
        self._add_param_entry(param_frame, "save_path", "保存路径:", self.params["save_path"], row=2)

        # --- 按钮区域 --- (左侧容器内，参数设置框下方，居中显示)
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=0, pady=8)

        # 创建一个内部框架来容纳按钮，实现居中
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')

        # 第一行按钮框架（测RIN、测底噪和种子光）
        first_row_frame = tk.Frame(inner_btn_frame)
        first_row_frame.pack(fill=tk.X, pady=(0, 6))  # 第一行与第二行之间有间距

        # 第二行按钮框架（连接和停止）
        second_row_frame = tk.Frame(inner_btn_frame)
        second_row_frame.pack(fill=tk.X)

        # 添加按钮
        self.btn_rin = tk.Button(first_row_frame, text="测RIN", command=self.start_rin, bg="#28862B", fg="#FFFFFF", width=10, cursor="hand2")
        self.btn_bg = tk.Button(first_row_frame, text="测底噪", command=self.start_background, bg="#28862B", fg="#FFFFFF", width=10, cursor="hand2")
        self.btn_seed = tk.Button(first_row_frame, text="种子光", command=self.start_seedlight, bg="#28862B", fg="#FFFFFF", width=10, cursor="hand2")
        self.btn_connect = tk.Button(second_row_frame, text="连接", command=self.connect_instrument, bg="#1D74C0", fg="#FFFFFF", width=10, cursor="hand2")
        self.btn_stop = tk.Button(second_row_frame, text="停止", command=self.stop_running, bg="#f44336", fg="#FFFFFF", width=10, cursor="hand2")
        self.btn_rename = tk.Button(second_row_frame, text="改名", command=self.rename_files, bg="#FF9800", fg="#FFFFFF", width=10, cursor="hand2")

        # 排列按钮
        # 第一行按钮居中
        first_row_spacer = tk.Label(first_row_frame)
        first_row_spacer.pack(side=tk.LEFT, expand=True)  # 左侧填充
        self.btn_rin.pack(side=tk.LEFT, padx=6)
        self.btn_bg.pack(side=tk.LEFT, padx=6)
        self.btn_seed.pack(side=tk.LEFT, padx=6)
        first_row_spacer2 = tk.Label(first_row_frame)
        first_row_spacer2.pack(side=tk.LEFT, expand=True)  # 右侧填充

        # 第二行按钮居中
        second_row_spacer = tk.Label(second_row_frame)
        second_row_spacer.pack(side=tk.LEFT, expand=True)  # 左侧填充
        self.btn_connect.pack(side=tk.LEFT, padx=6)
        self.btn_stop.pack(side=tk.LEFT, padx=6)
        self.btn_rename.pack(side=tk.LEFT, padx=6)
        second_row_spacer2 = tk.Label(second_row_frame)
        second_row_spacer2.pack(side=tk.LEFT, expand=True)  # 右侧填充

        # --- 日志显示区域 - 右侧 --- 占据整个右侧区域
        log_frame = tk.LabelFrame(main_container, text="运行日志", padx=5, pady=5)
        log_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.log_box = tk.Text(log_frame, wrap=tk.WORD)
        self.log_box.pack(fill=tk.BOTH, expand=True)

        # 设置grid权重，确保参数设置列固定，日志框列可以扩展
        main_container.grid_columnconfigure(0, weight=0)  # 参数设置列固定大小
        main_container.grid_columnconfigure(1, weight=1)  # 日志框列可以扩展
        main_container.grid_rowconfigure(0, weight=1)     # 第一行可以扩展

    def _add_param_entry(self, parent, key, label, default="", row=0, browse=None):
        tk.Label(parent, text=label, anchor="e", width=10).grid(row=row, column=0, sticky="e", padx=4, pady=4)
        ent = tk.Entry(parent, width=24)
        ent.insert(0, str(default))
        ent.grid(row=row, column=1, padx=4, pady=4)
        self.entries[key] = ent
        if browse == "file":
            tk.Button(parent, text="浏览", command=lambda k=key: self.browse_file(k), cursor="hand2").grid(row=row, column=2, padx=4, pady=4)
        if browse == "dir":
            tk.Button(parent, text="保存路径", command=lambda k=key: self.browse_savefile(k), cursor="hand2").grid(row=row, column=2, padx=4, pady=4)
        return ent

    # log() 复用基类 BaseTestGUI.log（线程安全，root.after 异步写入 log_box）

    def get_params(self) -> Dict[str, Any]:
        p = {}
        try:
            p["osa_ip"] = self.entries["osa_ip"].get().strip()
            #p["osa_port"] = int(self.entries["osa_port"].get().strip())
            p["save_path"] = self.entries["save_path"].get().strip() or self.params["save_path"]
        except Exception:
            p = self.params.copy()
        return p

    def browse_savefile(self, param_key: str):
        if messagebox.askyesno("选择", "选择保存目录？(否 = 选择具体文件名)"):
            dirname = filedialog.askdirectory(title="选择保存目录")
            if dirname:
                self.entries[param_key].delete(0, tk.END)
                self.entries[param_key].insert(0, dirname)
        else:
            filename = filedialog.asksaveasfilename(title="选择保存 文件", defaultextension=".csv", filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
            if filename:
                self.entries[param_key].delete(0, tk.END)
                self.entries[param_key].insert(0, filename)

    def browse_file(self, param_key: str):
        filename = filedialog.askopenfilename(title="选择文件", filetypes=[("所有文件", "*.*")])
        if filename:
            self.entries[param_key].delete(0, tk.END)
            self.entries[param_key].insert(0, filename)

    def set_dc_value(self):
        try:
            new_dc = simpledialog.askfloat("输入DC值", "请输入新的DC值:", minvalue=0, maxvalue=100, initialvalue=self.params["dc_initial"], parent=self.root)
            if new_dc is None:
                messagebox.showinfo("信息", "使用默认DC值 1.20")
                self.params["dc_initial"] = 2.40
            else:
                self.params["dc_initial"] = new_dc
                messagebox.showinfo("信息", f"已设置 DC 初始为 {new_dc}")
        except Exception as e:
            self.log("错误", f"设置 DC 值失败: {e}")

    # 诊断连接（快速尝试连接，不会改变任何测量逻辑）
    def connect_instrument(self):
        """
        诊断仪器连接
        快速尝试连接仪器以验证网络连通性
        """
        try:
            self.log("[连接] 正在尝试连接仪器...")
            inst = FSV3004Instrument(log_func=self.log)  # 临时创建一个测试连接实例
            ip = self.entries["osa_ip"].get().strip()
            success = inst.connect(ip)
            if success:
                self.log("[连接] 成功连接到 FSV3004 频谱仪。")
            else:
                self.log("[连接] 无法连接仪器，请检查地址或网络。")
        except Exception as e:
            self.log(f"[连接] 失败: {e}")

    def start_rin(self):
        """
        开始RIN测量任务
        在后台线程中运行RIN测量序列
        """
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()
        # DC 值在主线程弹窗输入（保留原行为）
        try:
            dc_input = float(self.entries["dc_value"].get())
            dc_for_ra = dc_input / 2.0
            self.log(f"[参数] DC 输入值 = {dc_input:.2f}，内部使用值 = {dc_for_ra:.2f}")
        except Exception:
            messagebox.showwarning("警告", "DC 值输入无效，将使用默认 2.40V")
            dc_for_ra = 1.20

        # 创建仪器 + 测试实例
        inst = FSV3004Instrument(log_func=self.log)
        test = RinTest(gui=self, log_func=self.log)
        test.instrument = inst
        test.dc_value = dc_for_ra
        # 把 GUI 中的保存目录传给测试实例，供 visualize_data 使用
        try:
            test.save_path = p.get("save_path") or self.params.get("save_path")
        except Exception:
            test.save_path = None
        test.ui_root = self.root

        self._active_test = test

        # run in background thread
        def target():
            try:
                self.running_task = "rin"
                self.btn_rin.config(state=tk.DISABLED)
                self.btn_bg.config(state=tk.DISABLED)
                self.btn_connect.config(state=tk.DISABLED)
                self.btn_stop.config(state=tk.NORMAL)
                test.stop_flag.clear()
                test.run()
            except Exception as e:
                self.log(f"[线程异常] {e}\n{traceback.format_exc()}")
            finally:
                try:
                    self.btn_rin.config(state=tk.NORMAL)
                    self.btn_bg.config(state=tk.NORMAL)
                    self.btn_connect.config(state=tk.NORMAL)
                    self.btn_stop.config(state=tk.DISABLED)
                except Exception:
                    pass
                self.running_task = None
                self._active_test = None
                # 一键测试模式：自动关闭窗口，触发进程退出
                self.auto_close(2000)

        self.worker_thread = threading.Thread(target=target, daemon=True)
        self.worker_thread.start()
        self.log("[主] RIN 测试线程已启动")

    # 开始底噪（线程）
    def start_background(self):
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()

        inst = FSV3004Instrument(log_func=self.log)
        test = BackgroundNoiseTest(gui=self, log_func=self.log)
        test.instrument = inst
        test.ip_address = p.get("osa_ip", "192.168.7.10")
        test.is_seedlight = False

        self._active_test = test

        def target_bg():
            try:
                self.running_task = "bg"
                self.btn_rin.config(state=tk.DISABLED)
                self.btn_bg.config(state=tk.DISABLED)
                self.btn_seed.config(state=tk.DISABLED)
                self.btn_connect.config(state=tk.DISABLED)
                self.btn_stop.config(state=tk.NORMAL)
                test.stop_flag.clear()
                test.run()
            except Exception as e:
                self.log(f"[线程异常] {e}\n{traceback.format_exc()}")
            finally:
                try:
                    self.btn_rin.config(state=tk.NORMAL)
                    self.btn_bg.config(state=tk.NORMAL)
                    self.btn_seed.config(state=tk.NORMAL)
                    self.btn_connect.config(state=tk.NORMAL)
                    self.btn_stop.config(state=tk.DISABLED)
                except Exception:
                    pass
                self.running_task = None
                self._active_test = None

        self.worker_thread = threading.Thread(target=target_bg, daemon=True)
        self.worker_thread.start()
        self.log("[主] 底噪测试线程已启动")

    def start_seedlight(self):
        """
        开始种子光测量任务
        在后台线程中运行种子光测量，功能与底噪测量相同但使用不同的文件名
        测量完成后会显示截图和数据文件，支持保存功能

        功能：
        - 检查是否有任务在运行，防止重复启动
        - 获取用户参数设置
        - 创建 BackgroundNoiseTest 实例
        - 在后台线程运行种子光测量
        - 管理按钮状态（禁用/启用）
        - 处理异常情况并记录日志

        注意：
        - 种子光测量与底噪测量使用相同的测量逻辑
        - 区别在于保存的文件名不同（SeedLight vs BackgroundNoise）
        - 使用is_seedlight=True参数标识种子光测量类型
        """
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()

        inst = FSV3004Instrument(log_func=self.log)
        test = BackgroundNoiseTest(gui=self, log_func=self.log)
        test.instrument = inst
        test.ip_address = p.get("osa_ip", "192.168.7.10")
        # 切换为种子光文件名
        test.screenshot_name = "SeedLight_Screen.png"
        test.dat_filename = "SeedLight.DAT"
        test.is_seedlight = True

        self._active_test = test

        def target_seed():
            try:
                self.running_task = "seed"
                self.btn_rin.config(state=tk.DISABLED)
                self.btn_bg.config(state=tk.DISABLED)
                self.btn_seed.config(state=tk.DISABLED)
                self.btn_connect.config(state=tk.DISABLED)
                self.btn_stop.config(state=tk.NORMAL)
                test.stop_flag.clear()
                test.run()
            except Exception as e:
                self.log(f"[线程异常] {e}\n{traceback.format_exc()}")
            finally:
                try:
                    self.btn_rin.config(state=tk.NORMAL)
                    self.btn_bg.config(state=tk.NORMAL)
                    self.btn_seed.config(state=tk.NORMAL)
                    self.btn_connect.config(state=tk.NORMAL)
                    self.btn_stop.config(state=tk.DISABLED)
                except Exception:
                    pass
                self.running_task = None
                self._active_test = None

        self.worker_thread = threading.Thread(target=target_seed, daemon=True)
        self.worker_thread.start()
        self.log("[主] 种子光测试线程已启动")

    def stop_running(self):
        """
        停止当前运行的测量任务
        发送停止信号给后台工作线程
        """
        # 通知当前活跃的测试实例停止
        if self._active_test is not None:
            self._active_test.stop_flag.set()
        self.log("[主] 停止命令已发送给后台任务")
        # 禁用停止按钮直到线程响应
        self.btn_stop.config(state=tk.DISABLED)

    def rename_files(self):
        """
        重命名保存目录中的文件
        将BackgroundNoise相关文件改名为'底噪'，SeedLight相关文件改名为'种子光'
        如果目标文件已存在，则添加时间戳避免覆盖
        """
        try:
            p = self.get_params()
            save_dir = p.get("save_path") or self.params.get("save_path")
            if not save_dir:
                self.log("错误", "未配置保存路径")
                return
            if not os.path.isdir(save_dir):
                self.log("错误", f"保存目录不存在: {save_dir}")
                return

            renamed = []
            for fname in os.listdir(save_dir):
                low = fname.lower()
                src = os.path.join(save_dir, fname)
                if not os.path.isfile(src):
                    continue
                base_cn = None
                # 匹配包含关键字的文件
                if 'backgroundnoise' in low or 'background' in low:
                    base_cn = '底噪'
                elif 'seedlight' in low or 'seed' in low:
                    base_cn = '种子光'
                if base_cn is None:
                    continue

                _, ext = os.path.splitext(fname)
                target_name = f"{base_cn}{ext}"
                dst = os.path.join(save_dir, target_name)
                if os.path.exists(dst):
                    ts = time.strftime("%Y%m%d_%H%M%S")
                    dst = os.path.join(save_dir, f"{base_cn}_{ts}{ext}")
                try:
                    os.rename(src, dst)
                    renamed.append((src, dst))
                    self.log(f"[改名] {src} -> {dst}")
                except Exception as e:
                    self.log(f"[改名] 重命名失败: {src} -> {dst} : {e}")
        except Exception as e:
            self.log(f"[改名] 出现异常: {e}")

    # run() 复用基类 BaseTestGUI.run（启动 mainloop）


# -------------------------
# Entry point
# -------------------------
if __name__ == "__main__":
    """
    程序主入口
    创建并运行RIN测量GUI应用程序
    """
    gui = RinGUI()
    gui.run()
