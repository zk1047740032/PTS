#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rin_FSV3004_CTStyle.py
改造自 Rin_FSV3004.py，界面与结构风格统一到 CT_W 风格：
- 保留原有测量逻辑、数据传输逻辑、绘图效果（未改动计算/命令/画图细节）
- 增加统一日志输出、参数输入区、线程化运行、开始/停止按钮
- 测试运行在后台线程，不阻塞 GUI
"""

from __future__ import annotations
import os
import time
import csv
import threading
import traceback
from typing import List, Optional, Any, Dict
import ctypes

import pyvisa
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from io import StringIO
from PIL import Image, ImageTk

# 启用DPI感知，解决高DPI屏幕下界面模糊问题
if os.name == 'nt':
    try:
        # 设置进程DPI感知
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        # 获取系统DPI
        dpi = ctypes.windll.user32.GetDpiForSystem()
        # 设置缩放因子
        scaling_factor = dpi / 96.0
    except Exception:
        scaling_factor = 1.0
else:
    scaling_factor = 1.0

# -------------------------
# Helpers
# -------------------------
def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)
    return path

def default_logger(msg: str):
    print(msg)

# -------------------------
# RinAnalyzer (原样逻辑，增加 log_func 支持)
# -------------------------
class RinAnalyzer:
    def __init__(self, log_func=default_logger):
        self.rm = None
        self.instrument = None
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
        self.stop_flag = False
        self.stop_window = None

        # logging
        self.log = log_func
        # 等待文件同步的默认超时（秒）及轮询间隔
        self.file_wait_timeout_s = 30.0
        self.file_wait_poll_s = 0.5

    # 连接仪器（保持原命令）
    def connect(self, ip_address="192.168.7.10", port=5025):
        try:
            self.rm = pyvisa.ResourceManager()
            # 使用 SOCKET 地址（与原脚本一致）
            self.instrument = self.rm.open_resource(f"TCPIP0::{ip_address}::{port}::SOCKET")
            self.instrument.timeout = 60000
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.log("成功连接到频谱分析仪")
            return True
        except Exception as e:
            self.log(f"连接失败: {e}")
            return False

    # 配置仪器（与原样）
    def configure_instrument(self):
        if not self.instrument:
            self.log("未连接到仪器")
            return
        self.instrument.write(":INST:SEL SA") # 选择频谱分析仪
        self.instrument.write(":CONF:SAN") # 配置频谱分析仪
        self.instrument.write("SWE:POIN 2001") # 设置采样点数
        self.instrument.write("UNIT:POW V") # 设置单位为伏特
        self.instrument.write("TRACE1:TYPE AVERage") # 设置追踪类型为平均
        self.log("仪器已配置（SWE:POIN 2001, UNIT: V, TRACE: AVERage）")

    # 测量函数（与原样）
    def measure_segment(self, start_freq, stop_freq, bandwidth, avg_count, filename):
        if not self.instrument:
            self.log("未连接到仪器")
            return False
        try:
            self.instrument.write(f":BAND {bandwidth}")
            self.instrument.write(f":AVER:COUN {avg_count}")
            self.instrument.write(f":FREQ:STAR {start_freq} Hz")
            self.instrument.write(f":FREQ:STOP {stop_freq} Hz")
            self.instrument.query("*OPC?")

            self.instrument.write(":INIT:CONT OFF")  # 关闭连续模式
            self.instrument.write(":INIT")
            self.instrument.query("*OPC?")
            
            instrument_path = f"C:\\PTS\\Rin\\{filename}"
            self.instrument.write("MMEM:MDIR 'C:\\PTS\\Rin'")
            self.instrument.query("*OPC?")

            self.instrument.write(f":MMEM:STOR:TRAC 1,'{instrument_path}'")
            self.instrument.query("*OPC?")
            self.log(f"数据已存储在仪器内部: {instrument_path}")

            # 复制到共享目录（保留原逻辑）
            instrument_ip = "192.168.7.10"
            source_path = "C:\\PTS\\Rin"
            dest_path = r"\\192.168.7.7\PTS\zhongzi\Rin\FSV3004"
            rm = pyvisa.ResourceManager()
            instr = rm.open_resource(f"TCPIP0::{instrument_ip}::inst0::INSTR")
            instr.write(f"MMEM:COPY '{source_path}\\*.*','{dest_path}'")
            instr.close()
            self.log(f"文件已从仪器复制到电脑共享文件夹：{dest_path}")

        except Exception as e:
            self.log(f"测量失败: {e}")
            return False
        return True

    def _parse_and_save_data(self, raw_data, filename):
        # 保留原解析逻辑（未改动核心解析）
        try:
            hash_pos = raw_data.find(b'#')
            if hash_pos == -1:
                self._parse_fallback_data(raw_data, filename)
                return

            digit_count = int(chr(raw_data[hash_pos+1]))
            data_length = int(raw_data[hash_pos+2:hash_pos+2+digit_count])
            data_start = hash_pos + 2 + digit_count
            data_block = raw_data[data_start:data_start+data_length]

            if not data_block:
                self.log(f"错误: 数据块为空，无法解析")
                return

            local_path = next((p for p in self.file_paths if filename.lower() in p.lower()), None)
            if not local_path:
                self.log(f"未找到本地保存路径: {filename}")
                return

            os.makedirs(os.path.dirname(local_path), exist_ok=True)

            with open(local_path, 'wb') as f:
                f.write(data_block)
            self.log(f"原始数据已保存到: {local_path}")

        except Exception as e:
            self.log(f"保存数据失败: {e}")

    # 备用解析（原脚本没有实现细节，这里保留占位以兼容）
    def _parse_fallback_data(self, raw_data, filename):
        # 如果没有 # 标记，保留原始行为：尝试以文本方式保存
        try:
            local_path = next((p for p in self.file_paths if filename.lower() in p.lower()), None)
            if not local_path:
                self.log(f"_parse_fallback_data: 未找到本地保存路径: {filename}")
                return
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            with open(local_path, 'wb') as f:
                f.write(raw_data)
            self.log(f"_parse_fallback_data: 原始数据已保存到: {local_path}")
        except Exception as e:
            self.log(f"_parse_fallback_data 保存失败: {e}")

    # DC值输入（保留原弹窗逻辑）
    def read_dc_value(self, parent=None):
        parent = parent or getattr(self, "ui_root", None)
        new_dc_value = simpledialog.askfloat("输入DC值", "请输入新的DC值:",
                                            minvalue=0, maxvalue=100, initialvalue=2.40,
                                            parent=parent)
        if new_dc_value is not None:
            self.dc_value = new_dc_value / 2
            self.log(f"DC值已更新为: {self.dc_value}")
        else:
            messagebox.showinfo("信息", "使用默认DC值.", parent=parent)
            self.dc_value = 1.20

    # 读取 CSV 数据（保持原逻辑）
    def read_data_from_csv(self, file_path):
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
                            x = float(row[0])
                            y = float(row[1])
                            file_dx.append(x)
                            file_dy.append(y)
                        except ValueError:
                            continue
            if len(file_dy) != 2001:
                self.log(f"警告: 数据点数非2001，实际 {len(file_dy)}")
            self.dx.append(file_dx)
            self.dy.append(file_dy)
            return True
        except Exception as e:
            self.log(f"读取文件失败 {file_path}: {e}")
            return False

    # 处理文件（保留原逻辑，稍作 logger 替换）
    def process_files(self):
        self.dx = []
        self.dy = []
        self.ddx = []
        self.ddy = []

        for file_path in self.file_paths:
            # 等待文件被复制/同步到本地（存在且大小>0）
            waited = 0.0
            file_ready = False
            while waited < getattr(self, 'file_wait_timeout_s', 30.0):
                if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                    file_ready = True
                    break
                time.sleep(getattr(self, 'file_wait_poll_s', 0.5))
                waited += getattr(self, 'file_wait_poll_s', 0.5)

            if not file_ready:
                self.log(f"文件不存在或未同步（等待{getattr(self,'file_wait_timeout_s',30.0)}s）: {file_path}")
                # 保留原行为：把占位空列表加入以维持索引
                self.dx.append([])
                self.dy.append([])
                continue

            # 尝试读取文件，若失败则重试几次（读取可能因文件正在被写入而瞬时失败）
            read_ok = False
            read_attempts = 0
            max_read_attempts = 3
            while read_attempts < max_read_attempts and not read_ok:
                try:
                    if self.read_data_from_csv(file_path):
                        self.log(f"成功读取: {file_path}")
                        read_ok = True
                        break
                    else:
                        self.log(f"读取失败（尝试{read_attempts+1}）: {file_path}")
                except Exception as e:
                    self.log(f"读取异常（尝试{read_attempts+1}）: {file_path} -> {e}")
                read_attempts += 1
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

    # compute_rin_power 保持原实现
    def compute_rin_power(self, x, y):
        power = []
        segment_length = 6
        for k in range(1, len(x) // segment_length + 1):
            sub_x = x[:k*segment_length]
            sub_y_exp = [np.power(10, val / 10.0) if np.isfinite(val) else 0 for val in y[:k*segment_length]]
            integral = sum((sub_x[i] - sub_x[i-1]) * (sub_y_exp[i] + sub_y_exp[i-1]) / 2.0
                           for i in range(1, len(sub_x)))
            power.append(np.sqrt(integral))
        return power

    # visualize_data 完整保留（仅把 print 改为 self.log）
    def visualize_data(self):
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
        tk.Button(top_frame, text="保存", command=save_figure, font=('SimHei', 20)).pack(side=tk.TOP)

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
        ax2.plot(self.ddx[::6], adjusted_power, color="#085cab", linewidth=2)
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

        # 自动保存图像为 Rin.png：优先使用 GUI 提供的保存目录（ra.save_path），否则回退到第一个 data 路径所在目录或 cwd
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
            except Exception as e:
                self.log(f"[保存] 自动保存Rin图片失败: {e}")
        except Exception as e:
            self.log(f"[保存] 生成自动保存路径失败: {e}")

        # 获取三段曲线最高点 + 指定频点的RIN值
        target_xs = [1000, 10000, 100000, 1000000]

        # 将驰豫振荡峰的检测范围限制在 1e5 - 8e6 Hz
        relax_start = 1e5
        relax_stop = 8e6

        # 根据整个频率数组筛选出位于检测范围内的索引
        freqs = np.array(self.ddx)
        ys = np.array(self.ddy)
        mask = (freqs >= relax_start) & (freqs <= relax_stop) & np.isfinite(ys)
        if np.any(mask):
            masked_freqs = freqs[mask]
            masked_ys = ys[mask]
            max_idx = np.argmax(masked_ys)
            highest_rin = float(masked_ys[max_idx])
            peak_freq = float(masked_freqs[max_idx])
        else:
            highest_rin = float('nan')
            peak_freq = float('nan')

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

    def request_stop(self):
        self.log("[用户操作] 请求停止 RIN 测试")
        self.stop_flag = True

    def close(self):
        if self.instrument:
            try:
                self.instrument.close()
            except Exception:
                pass
        if self.rm:
            try:
                self.rm.close()
            except Exception:
                pass
        self.log("已关闭仪器连接")


# -------------------------
# BackgroundNoiseAnalyzer（保留原逻辑，增加 log）
# -------------------------
class BackgroundNoiseAnalyzer:
    def __init__(self, log_func=default_logger):
        self.rm = None
        self.instrument = None
        self.log = log_func

    def connect(self, ip_address="192.168.7.10", port=5025):
        try:
            self.rm = pyvisa.ResourceManager()
            self.instrument = self.rm.open_resource(f"TCPIP0::{ip_address}::{port}::SOCKET")
            self.instrument.timeout = 60000
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.log("成功连接到频谱分析仪")
            return True
        except Exception as e:
            self.log(f"连接失败: {e}")
            return False

    def measure_and_screenshot(self, start_freq=10, stop_freq=100_000_000, bandwidth=30, avg_count=1, screenshot_name="BackgroundNoise_Screen.png", dat_filename="BackgroundNoise.DAT", is_seedlight=False):
        if not self.instrument:
            self.log("未连接到仪器")
            return False
        try:
            # 配置参数
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

            # 启动测量
            self.instrument.write(":INIT:CONT OFF")  # 关闭连续模式
            self.instrument.write(":INIT")
            self.instrument.query("*OPC?")
            
            # 根据类型显示不同的日志
            if is_seedlight:
                self.log("种子光测量完成，开始截图和保存数据...")
            else:
                self.log("底噪测量完成，开始截图和保存数据...")

            # 保存数据到仪器
            instrument_path = f"C:\\PTS\\Rin\\{dat_filename}"
            self.instrument.write("MMEM:MDIR 'C:\\PTS\\Rin'")
            self.instrument.query("*OPC?")
            self.instrument.write(f":MMEM:STOR:TRAC 1,'{instrument_path}'")
            self.instrument.query("*OPC?")
            
            # 根据类型显示不同的日志
            if is_seedlight:
                self.log(f"种子光数据已存储在仪器内部: {instrument_path}")
            else:
                self.log(f"底噪数据已存储在仪器内部: {instrument_path}")

            # 截图并保存到仪器
            self.instrument.write("HCOPy:DEST 'MMEM'")
            self.instrument.write(f"MMEM:NAME 'C:\\PTS\\Rin\\{screenshot_name}'")
            self.instrument.write("HCOPy:IMM")
            self.instrument.query("*OPC?")
            self.log("仪器已截图并保存。")

            # 复制到共享目录（按 Rin 的简化实现：一次性复制整个仪器目录到目标）
            instrument_ip = "192.168.7.10"
            source_path = "C:\\PTS\\Rin"
            dest_path = r"\\192.168.7.7\PTS\zhongzi\Rin\FSV3004"
            rm = pyvisa.ResourceManager()
            instr = rm.open_resource(f"TCPIP0::{instrument_ip}::inst0::INSTR")
            # 使用通配符一次性复制（与 Rin 的实现保持一致，注意：健壮性较低，但与用户要求一致）
            instr.write(f"MMEM:COPY '{source_path}\\*.*','{dest_path}'")
            instr.close()
            self.log(f"文件已从仪器复制到电脑共享文件夹：{dest_path}")

            # 直接尝试显示截图（不等待文件同步，行为与 Rin 保持一致）
            self.show_screenshot(dest_path, screenshot_name, dat_filename, is_seedlight)
            self.log("已发送复制命令并尝试显示图片。")

        except Exception as e:
            self.log(f"底噪测量或截图失败: {e}")
            return False
        return True

    def show_screenshot(self, dest_path, screenshot_name, dat_filename="BackgroundNoise.DAT", is_seedlight=False):
        local_img_path = os.path.join(dest_path, screenshot_name)
        # 数据文件路径
        local_dat_path = os.path.join(dest_path, dat_filename)
        
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
                    messagebox.showerror("保存失败", f"保存图片时出错: {e}")
        
        def save_data():
            if not os.path.exists(local_dat_path):
                messagebox.showerror("错误", "数据文件不存在")
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
                    messagebox.showerror("保存失败", f"保存数据时出错: {e}")
        
        # 创建按钮框架来放置两个按钮
        btn_frame = tk.Frame(top_frame)
        btn_frame.pack(side=tk.TOP)
        
        # 保存图片按钮
        tk.Button(btn_frame, text="保存图片", command=save_image, font=('SimHei', 16)).pack(side=tk.LEFT, padx=10)
        # 保存数据按钮
        tk.Button(btn_frame, text="保存数据", command=save_data, font=('SimHei', 16)).pack(side=tk.LEFT, padx=10)
        
        # 等待图片文件同步到本机（最多等待 timeout 秒）
        timeout = 10.0  # seconds
        poll_interval = 0.5
        waited = 0.0
        while not (os.path.exists(local_img_path) and os.path.getsize(local_img_path) > 0) and waited < timeout:
            time.sleep(poll_interval)
            waited += poll_interval

        if not os.path.exists(local_img_path) or os.path.getsize(local_img_path) == 0:
            # 未同步到本地，显示文字提示而不是直接抛错
            msg = f"图片尚未同步到电脑（等待{timeout}s未出现）：{local_img_path}"
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

# -------------------------
# TestRunner: 负责在后台运行 Rin / BackgroundNoise 流程（保持原测量逻辑）
# -------------------------
class TestRunner:
    def __init__(self, log_func=default_logger):
        self.log = log_func
        self._stop = False

    def stop(self):
        self._stop = True
        self.log("[Runner] 停止信号已设置")

    def run_rin(self, ra: RinAnalyzer, ui_root: tk.Tk):
        """
        Run RIN sequence - this mirrors the original Rin(ra) function behavior but routed through log_func.
        保持原来测量段、顺序、文件拷贝、process_files、visualize_data 等逻辑不变。
        """
        try:
            try:
                self.log("[初始化] 正在清空共享文件夹和仪器内部文件夹...")

                # 电脑共享目录
                local_dir = r"\\192.168.7.7\\PTS\\zhongzi\\Rin\\FSV3004"
                if os.path.exists(local_dir):
                    for f in os.listdir(local_dir):
                        fp = os.path.join(local_dir, f)
                        try:
                            if os.path.isfile(fp) or os.path.islink(fp):
                                os.remove(fp)
                            elif os.path.isdir(fp):
                                import shutil
                                shutil.rmtree(fp)
                        except Exception as e:
                            self.log(f"[警告] 删除 {fp} 失败: {e}")

                # 仪器目录清空
                try:
                    ip = "192.168.7.10"
                    rm = pyvisa.ResourceManager()
                    inst = rm.open_resource(f"TCPIP0::{ip}::5025::SOCKET")
                    inst.write("MMEM:MDIR 'C:\\PTS\\Rin'")  # 确保路径存在
                    inst.write("MMEM:DEL 'C:\\PTS\\Rin\\*.*'")
                    #inst.query("*OPC?")
                    inst.close()
                    rm.close()
                except Exception as e:
                    self.log(f"[警告] 仪器文件夹清理失败: {e}")

                self.log("[初始化] 文件夹清理完成。")
            except Exception as e:
                self.log(f"[错误] 清理文件夹时出错: {e}")
            ra.ui_root = ui_root
            ra.log = self.log  # route analyzer logs to gui

            if ra.connect():
                ra.configure_instrument()
                measurement_params = [
                    (10, 100, 5, 20, "Rin_1.DAT"),
                    (100, 1000, 5, 20, "Rin_2.DAT"),
                    (1000, 10000, 30, 20, "Rin_3.DAT"),
                    (10000, 100000, 30, 20, "Rin_4.DAT"),
                    (100000, 1000000, 30, 20, "Rin_5.DAT"),
                    (1000000, 10000000, 30, 20, "Rin_6.DAT")
                ]
                for start, stop, bw, avg, fname in measurement_params:
                    if self._stop or ra.stop_flag:
                        self.log("[测试] RIN 测试已被终止")
                        break
                    self.log(f"[测试] 正在测量: {start}Hz - {stop}Hz, 带宽: {bw}Hz")
                    ra.measure_segment(start, stop, bw, avg, fname)
                # 关闭连接
                try:
                    ra.close()
                except Exception:
                    pass
            else:
                self.log("[测试] 无法连接到仪器，RIN 测试终止")

            if self._stop or ra.stop_flag:
                # 如果停止，更新 UI stop_window
                def _notify_stopped():
                    if ra.stop_window and ra.stop_window.winfo_exists():
                        ra.stop_window.destroy()
                        ra.stop_window = None
                try:
                    ra.ui_root.after(0, _notify_stopped)
                except Exception:
                    pass
                return

            # 原脚本会在这里处理文件与可视化
            if ra.stop_window and ra.stop_window.winfo_exists():
                try:
                    ra.stop_window.destroy()
                except Exception:
                    pass
                ra.stop_window = None

            self.log("正在处理数据...")
            ra.process_files()
            self.log("正在显示可视化结果...")
            ra.visualize_data()
            self.log("程序执行完毕")

        except Exception as e:
            self.log(f"[Runner Exception] {e}\n{traceback.format_exc()}")

    def run_background(self, bna: BackgroundNoiseAnalyzer, ui_root: tk.Tk, is_seedlight=False):
        try:
            bna.log = self.log
            if bna.connect():
                # 根据是否为种子光设置不同的文件名
                if is_seedlight:
                    screenshot_name = "SeedLight_Screen.png"
                    dat_filename = "SeedLight.DAT"
                else:
                    screenshot_name = "BackgroundNoise_Screen.png"
                    dat_filename = "BackgroundNoise.DAT"
                
                bna.measure_and_screenshot(screenshot_name=screenshot_name, 
                                          dat_filename=dat_filename, 
                                          is_seedlight=is_seedlight)
                try:
                    bna.instrument.close()
                except Exception:
                    pass
                try:
                    bna.rm.close()
                except Exception:
                    pass
            else:
                messagebox.showerror("错误", "无法连接到仪器")
        except Exception as e:
            self.log(f"[Background Exception] {e}\n{traceback.format_exc()}")

# -------------------------
# GUI: CT 风格（参数区 + 日志区）
# -------------------------
class RinGUI:
    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
        else:
            self.root = tk.Toplevel(parent)
        self.root.title("Rin_FSV3004")
        self.root.resizable(True, True)
        self.set_center(1170, 330)

        # 默认参数（保留原脚本默认路径/IP）
        self.params = {
            "osa_ip": "192.168.7.10",
            #"osa_port": 5025,
            "save_path": r"C:\PTS\zhongzi\Rin\FSV3004",
            "dc_initial": 2.40
        }
        self.entries: Dict[str, tk.Entry] = {}
        self.runner = TestRunner(log_func=self.log)
        self.worker_thread: Optional[threading.Thread] = None
        self.running_task: Optional[str] = None

        self.create_widgets()

    def set_center(self, width: int, height: int):
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        posx = (screenwidth - width) // 2
        posy = (screenheight - height) // 2
        self.root.geometry(f'{width}x{height}+{posx}+{posy}')

    def create_widgets(self):
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
        self.btn_rin = tk.Button(first_row_frame, text="测RIN", command=self.start_rin, bg="#28862B", fg="#FFFFFF", width=10)
        self.btn_bg = tk.Button(first_row_frame, text="测底噪", command=self.start_background, bg="#28862B", fg="#FFFFFF", width=10)
        self.btn_seed = tk.Button(first_row_frame, text="种子光", command=self.start_seedlight, bg="#28862B", fg="#FFFFFF", width=10)
        self.btn_connect = tk.Button(second_row_frame, text="连接", command=self.connect_instrument, bg="#1D74C0", fg="#FFFFFF", width=10)
        self.btn_stop = tk.Button(second_row_frame, text="停止", command=self.stop_running, bg="#f44336", fg="#FFFFFF", width=10)
        self.btn_rename = tk.Button(second_row_frame, text="改名", command=self.rename_files, bg="#FF9800", fg="#FFFFFF", width=10)
        
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
            tk.Button(parent, text="浏览", command=lambda k=key: self.browse_file(k)).grid(row=row, column=2, padx=4, pady=4)
        if browse == "dir":
            tk.Button(parent, text="保存路径", command=lambda k=key: self.browse_savefile(k)).grid(row=row, column=2, padx=4, pady=4)
        return ent

    def log(self, msg: str):
        t = time.strftime("[%H:%M:%S]")
        try:
            self.log_box.insert(tk.END, f"{t} {msg}\n")
            self.log_box.see(tk.END)
            self.root.update_idletasks()
        except Exception:
            pass
        # 也打印到 stdout，方便日志文件或控制台查看
        print(f"{t} {msg}")

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
            messagebox.showerror("错误", f"设置 DC 值失败: {e}")

    # 诊断连接（快速尝试连接，不会改变任何测量逻辑）
    def connect_instrument(self):
        try:
            self.log("[连接] 正在尝试连接仪器...")
            ra = RinAnalyzer(log_func=self.log)  # 临时创建一个测试连接实例
            ip = self.entries["osa_ip"].get().strip()
            port = 5025
            success = ra.connect(ip, port)
            if success:
                self.log("[连接] 成功连接到 FSV3004 频谱仪。")
            else:
                self.log("[连接] 无法连接仪器，请检查地址或网络。")
        except Exception as e:
            self.log(f"[连接] 失败: {e}")

    # 开始 RIN（线程）
    def start_rin(self):
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

        # 创建 RinAnalyzer 并设置 dc
        ra = RinAnalyzer(log_func=self.log)
        ra.dc_value = dc_for_ra
        # 把 GUI 中的保存目录传给 RinAnalyzer，供 visualize_data 使用
        try:
            ra.save_path = p.get("save_path") or self.params.get("save_path")
        except Exception:
            ra.save_path = None
        ra.ui_root = self.root
        ra.stop_flag = False

        # run in background thread using TestRunner
        def target():
            try:
                self.running_task = "rin"
                self.btn_rin.config(state=tk.DISABLED)
                self.btn_bg.config(state=tk.DISABLED)
                self.btn_connect.config(state=tk.DISABLED)
                self.btn_stop.config(state=tk.NORMAL)
                self.runner._stop = False
                self.runner.run_rin(ra, self.root)
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

        self.worker_thread = threading.Thread(target=target, daemon=True)
        self.worker_thread.start()
        self.log("[主] RIN 测试线程已启动")

    # 开始底噪（线程）
    def start_background(self):
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()
        bna = BackgroundNoiseAnalyzer(log_func=self.log)

        def target_bg():
            try:
                self.running_task = "bg"
                self.btn_rin.config(state=tk.DISABLED)
                self.btn_bg.config(state=tk.DISABLED)
                self.btn_seed.config(state=tk.DISABLED)
                self.btn_connect.config(state=tk.DISABLED)
                self.btn_stop.config(state=tk.NORMAL)
                self.runner._stop = False
                self.runner.run_background(bna, self.root, is_seedlight=False)
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

        self.worker_thread = threading.Thread(target=target_bg, daemon=True)
        self.worker_thread.start()
        self.log("[主] 底噪测试线程已启动")
    
    # 开始种子光（线程）- 功能与底噪相同但使用不同的文件名
    def start_seedlight(self):
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()
        bna = BackgroundNoiseAnalyzer(log_func=self.log)

        def target_seed():
            try:
                self.running_task = "seed"
                self.btn_rin.config(state=tk.DISABLED)
                self.btn_bg.config(state=tk.DISABLED)
                self.btn_seed.config(state=tk.DISABLED)
                self.btn_connect.config(state=tk.DISABLED)
                self.btn_stop.config(state=tk.NORMAL)
                self.runner._stop = False
                self.runner.run_background(bna, self.root, is_seedlight=True)
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

        self.worker_thread = threading.Thread(target=target_seed, daemon=True)
        self.worker_thread.start()
        self.log("[主] 种子光测试线程已启动")

    def stop_running(self):
        # 通知 runner 停止
        self.runner.stop()
        self.log("[主] 停止命令已发送给后台任务")
        # also try to set stop_flag on analyzer if available via thread local variables:
        # (original analyzers set stop_flag themselves via their stop window "请求停止")
        # 禁用停止按钮直到线程响应
        self.btn_stop.config(state=tk.DISABLED)

    def rename_files(self):
        """把保存目录中包含 BackgroundNoise 的文件改名为 中文 '底噪'，包含 SeedLight 的文件改名为 '种子光'。
        如果目标文件名已存在，则追加时间戳以避免覆盖。
        """
        try:
            p = self.get_params()
            save_dir = p.get("save_path") or self.params.get("save_path")
            if not save_dir:
                messagebox.showerror("错误", "未配置保存路径")
                return
            if not os.path.isdir(save_dir):
                messagebox.showerror("错误", f"保存目录不存在: {save_dir}")
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

    def run(self):
        if self.root.winfo_exists():
            self.root.mainloop()

# -------------------------
# Entry point
# -------------------------
if __name__ == "__main__":
    gui = RinGUI()
    gui.run()