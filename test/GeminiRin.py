#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rin_FSV3004_CTStyle.py
改造自 Rin_FSV3004.py，界面与结构风格统一到 CT_W 风格：
- 修复：IP地址不再硬编码，而是使用界面输入的参数
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
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        dpi = ctypes.windll.user32.GetDpiForSystem()
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
# RinAnalyzer
# -------------------------
class RinAnalyzer:
    def __init__(self, log_func=default_logger):
        self.rm = None
        self.instrument = None
        self.target_ip = "192.168.7.10"  # 默认值，会被connect覆盖
        self.dc_value = 1.20
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

        self.log = log_func
        self.file_wait_timeout_s = 30.0
        self.file_wait_poll_s = 0.5

    # [修改] 增加 self.target_ip 的赋值
    def connect(self, ip_address="192.168.7.10", port=5025):
        try:
            self.target_ip = ip_address  # 保存传入的IP
            self.rm = pyvisa.ResourceManager()
            self.instrument = self.rm.open_resource(f"TCPIP0::{ip_address}::{port}::SOCKET")
            self.instrument.timeout = 60000
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.log(f"成功连接到频谱分析仪 ({ip_address})")
            return True
        except Exception as e:
            self.log(f"连接失败: {e}")
            return False

    def configure_instrument(self):
        if not self.instrument:
            self.log("未连接到仪器")
            return
        self.instrument.write(":INST:SEL SAN")
        self.instrument.write("SWE:POIN 2001")
        self.instrument.write("DISPlay:TRACe1:MODE Average")
        self.instrument.write("DETector1:FUNCtion RMS")
        self.instrument.write("INPut:COUPling DC")
        self.instrument.write("CALCulate:UNIT:POWer V")
        self.log("仪器已配置（扫描点数：2001, 单位: V, 追踪模式: AVERage）")

    # [修改] 使用 self.target_ip 替代硬编码
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

            self.instrument.write(":INIT:CONT OFF")
            self.instrument.write(":INIT")
            self.instrument.query("*OPC?")
            
            instrument_path = f"C:\\PTS\\Rin\\{filename}"
            self.instrument.write("MMEM:MDIR 'C:\\PTS\\Rin'")
            self.instrument.query("*OPC?")

            self.instrument.write(f":MMEM:STOR:TRAC 1,'{instrument_path}'")
            self.instrument.query("*OPC?")
            self.log(f"数据已存储在仪器内部: {instrument_path}")

            # [修改] 使用动态 IP
            instrument_ip = self.target_ip 
            source_path = "C:\\PTS\\Rin"
            dest_path = r"\\192.168.7.7\PTS\zhongzi\Rin\FSV3004" # 注意：这里是目标电脑共享路径，保持原样还是也需要改？通常这是本机IP，暂时保持原样。
            
            # 使用临时 resource manager 连接仪器文件服务
            rm = pyvisa.ResourceManager()
            instr = rm.open_resource(f"TCPIP0::{instrument_ip}::inst0::INSTR")
            instr.write(f"MMEM:COPY '{source_path}\\*.*','{dest_path}'")
            instr.close()
            self.log(f"文件已从仪器({instrument_ip})复制到电脑共享文件夹：{dest_path}")

        except Exception as e:
            self.log(f"测量失败: {e}")
            return False
        return True

    def _parse_and_save_data(self, raw_data, filename):
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

    def _parse_fallback_data(self, raw_data, filename):
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
            if len(file_dy) < 2001:
                self.log(f"警告: 数据点数非2001，实际 {len(file_dy)}")
                return False # Fixed typo: false -> False
            self.dx.append(file_dx)
            self.dy.append(file_dy)
            return True
        except Exception as e:
            self.log(f"读取文件失败 {file_path}: {e}")
            return False

    def process_files(self):
        self.dx = []
        self.dy = []
        self.ddx = []
        self.ddy = []

        for file_path in self.file_paths:
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
                self.dx.append([])
                self.dy.append([])
                continue

            read_ok = False
            read_attempts = 0
            max_read_attempts = 10
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
                time.sleep(1.0)

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
                    self.ddy.append(float('-inf'))
                else:
                    rin_value = 20 * np.log10(v_noise / (self.dc_value * self.amplification * scale_factor))
                    self.ddy.append(rin_value)

        if self.ddx and self.ddy:
            self.RIN_power = self.compute_rin_power(self.ddx, self.ddy)
        else:
            self.log("错误: 无有效数据可处理")
            self.RIN_power = []

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
        
        top_frame = tk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=10)
        tk.Button(top_frame, text="保存", command=save_figure, font=('SimHei', 20)).pack(side=tk.TOP)

        """图1: RIN曲线"""
        ax1.plot(self.ddx, self.ddy, color="#085cab", linewidth=2)
        ax1.set_xscale('log')
        ax1.margins(x=0)
        ax1.tick_params(axis='both', which='major', labelsize=20, pad=5, length=12, width=3, direction='in')
        ax1.set_ylabel('RIN (dBc/Hz)', fontsize=18, fontweight='bold', fontstyle='normal')
        ax1.grid(True, which='both', lw = 2, linestyle='--', alpha=1)
        for spine in ax1.spines.values():
            spine.set_linewidth(2.5)
        for label in ax1.get_xticklabels() + ax1.get_yticklabels():
            label.set_fontname('Times New Roman')
            label.set_fontsize(20)
            label.set_fontweight('bold')
        
        finite_ddy = [v for v in self.ddy if np.isfinite(v)]
        if finite_ddy:
            ax1.set_ylim(np.floor(min(finite_ddy)/10)*10, np.ceil(max(finite_ddy)/10)*10)
        ax1.yaxis.set_major_locator(MaxNLocator(nbins=6, integer=True))
        
        adjusted_power = [p * 100 for p in self.RIN_power]

        """图2: RMS积分曲线"""
        plot_len = len(adjusted_power)
        ax2.plot(self.ddx[::6][:plot_len], adjusted_power, color="#085cab", linewidth=2)
        ax2.set_xscale('log')
        ax2.margins(x=0)
        y2_min, y2_max = np.min(adjusted_power), np.max(adjusted_power)
        if np.isclose(y2_min, y2_max):
            y2_min = y2_min - 1
            y2_max = y2_max + 1
        y2_mid = (y2_min + y2_max) / 2
        ax2.set_yticks([y2_min, y2_mid, y2_max])
        ax2.set_yticklabels([f"{y2_min:.3f}%", f"{y2_mid:.3f}%", f"{y2_max:.3f}%"])   

        ax2.tick_params(axis='both', which='major', labelsize=15, pad=5, length=12, width=3, direction='in')
        ax2.set_xlabel('Frequency(Hz)', fontsize=18, fontweight='bold', fontstyle='normal')
        ax2.set_ylabel('Integrated RMS', fontsize=18, fontweight='bold', fontstyle='normal')
        ax2.grid(True, which='both', lw=2, linestyle='--', alpha=1)
        for spine in ax2.spines.values():
            spine.set_linewidth(2.5)
        for label in ax2.get_xticklabels() + ax2.get_yticklabels():
            label.set_fontname('Times New Roman')
            label.set_fontsize(20)
            label.set_fontweight('bold')

        plt.tight_layout()
        plt.subplots_adjust(hspace=0.15)

        canvas = FigureCanvasTkAgg(fig, master=root)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        try:
            save_dir = None
            if hasattr(self, 'save_path') and self.save_path:
                save_dir = str(self.save_path)
            if not save_dir and self.file_paths and len(self.file_paths) > 0:
                save_dir = os.path.dirname(self.file_paths[0])
            if not save_dir:
                save_dir = os.getcwd()
            os.makedirs(save_dir, exist_ok=True)
            auto_path = os.path.join(save_dir, "Rin.png")

            try:
                root.update_idletasks()
            except Exception:
                pass

            try:
                fig.savefig(auto_path, dpi=300, bbox_inches='tight')
                self.log(f"[保存] 自动保存Rin图片: {auto_path}")
            except Exception as e:
                self.log(f"[保存] 自动保存Rin图片失败: {e}")
        except Exception as e:
            self.log(f"[保存] 生成自动保存路径失败: {e}")

        target_xs = [1000, 10000, 100000, 1000000]
        relax_start = 1e5
        relax_stop = 1e7

        freqs = np.array(self.ddx)
        ys = np.array(self.ddy)
        
        mask = (freqs >= relax_start) & (freqs <= relax_stop) & np.isfinite(ys)
        
        highest_rin = float('nan')
        peak_freq = float('nan')

        if np.any(mask):
            masked_freqs = freqs[mask]
            masked_ys = ys[mask]
            if len(masked_ys) >= 3:
                center_vals = masked_ys[1:-1]
                left_vals = masked_ys[:-2]
                right_vals = masked_ys[2:]
                is_peak = (center_vals > left_vals) & (center_vals > right_vals)
                if np.any(is_peak):
                    peak_indices = np.where(is_peak)[0] + 1
                    best_peak_idx = peak_indices[np.argmax(masked_ys[peak_indices])]
                    highest_rin = float(masked_ys[best_peak_idx])
                    peak_freq = float(masked_freqs[best_peak_idx])
                else:
                    max_idx = np.argmax(masked_ys)
                    highest_rin = float(masked_ys[max_idx])
                    peak_freq = float(masked_freqs[max_idx])
            else:
                max_idx = np.argmax(masked_ys)
                highest_rin = float(masked_ys[max_idx])
                peak_freq = float(masked_freqs[max_idx])

        result_text = f"驰豫振荡峰 ({int(relax_start):d} - {int(relax_stop):d} Hz): {highest_rin:.3f} dBc/Hz @ {peak_freq:.0f} Hz\n\n"
        for tx in target_xs:
            if len(self.ddx) == 0:
                continue
            idx = np.argmin(np.abs(np.array(self.ddx) - tx))
            x_val = self.ddx[idx]
            y_val = self.ddy[idx]
            if not np.isfinite(y_val):
                result_text += f"x={x_val:.0f} Hz 时, y=无效数据\n"
            else:
                result_text += f"x={x_val:.0f} Hz 时, y={y_val:.3f} dBc/Hz\n"

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
# BackgroundNoiseAnalyzer
# -------------------------
class BackgroundNoiseAnalyzer:
    def __init__(self, log_func=default_logger):
        self.rm = None
        self.instrument = None
        self.log = log_func
        self.target_ip = "192.168.7.10" # 默认值

    # [修改] 增加 self.target_ip 的赋值
    def connect(self, ip_address="192.168.7.10", port=5025):
        try:
            self.target_ip = ip_address
            self.rm = pyvisa.ResourceManager()
            self.instrument = self.rm.open_resource(f"TCPIP0::{ip_address}::{port}::SOCKET")
            self.instrument.timeout = 60000
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.log(f"成功连接到频谱分析仪 ({ip_address})")
            return True
        except Exception as e:
            self.log(f"连接失败: {e}")
            return False

    # [修改] 使用 self.target_ip
    def measure_and_screenshot(self, start_freq=10, stop_freq=100_000_000, bandwidth=30, avg_count=1, screenshot_name="BackgroundNoise_Screen.png", dat_filename="BackgroundNoise.DAT", is_seedlight=False):
        if not self.instrument:
            self.log("未连接到仪器")
            return False
        try:
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

            self.instrument.write(":INIT:CONT OFF")
            self.instrument.write(":INIT")
            self.instrument.query("*OPC?")
            
            if is_seedlight:
                self.log("种子光测量完成，开始截图和保存数据...")
            else:
                self.log("底噪测量完成，开始截图和保存数据...")

            instrument_path = f"C:\\PTS\\Rin\\{dat_filename}"
            self.instrument.write("MMEM:MDIR 'C:\\PTS\\Rin'")
            self.instrument.query("*OPC?")
            self.instrument.write(f":MMEM:STOR:TRAC 1,'{instrument_path}'")
            self.instrument.query("*OPC?")
            
            if is_seedlight:
                self.log(f"种子光数据已存储在仪器内部: {instrument_path}")
            else:
                self.log(f"底噪数据已存储在仪器内部: {instrument_path}")

            self.instrument.write("HCOPy:DEST 'MMEM'")
            self.instrument.write(f"MMEM:NAME 'C:\\PTS\\Rin\\{screenshot_name}'")
            self.instrument.write("HCOPy:IMM")
            self.instrument.query("*OPC?")
            self.log("仪器已截图并保存。")

            # [修改] 使用动态 IP
            instrument_ip = self.target_ip
            source_path = "C:\\PTS\\Rin"
            dest_path = r"\\192.168.7.7\PTS\zhongzi\Rin\FSV3004"
            
            rm = pyvisa.ResourceManager()
            instr = rm.open_resource(f"TCPIP0::{instrument_ip}::inst0::INSTR")
            instr.write(f"MMEM:COPY '{source_path}\\*.*','{dest_path}'")
            instr.close()
            self.log(f"文件已从仪器({instrument_ip})复制到电脑共享文件夹：{dest_path}")

            self.show_screenshot(dest_path, screenshot_name, dat_filename, is_seedlight)
            self.log("已发送复制命令并尝试显示图片。")

        except Exception as e:
            self.log(f"底噪测量或截图失败: {e}")
            return False
        return True

    def show_screenshot(self, dest_path, screenshot_name, dat_filename="BackgroundNoise.DAT", is_seedlight=False):
        local_img_path = os.path.join(dest_path, screenshot_name)
        local_dat_path = os.path.join(dest_path, dat_filename)
        
        win = tk.Toplevel()
        if is_seedlight:
            win.title("种子光仪器截图")
        else:
            win.title("底噪仪器截图")
        
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
            
            default_filename = "SeedLight.dat" if is_seedlight else "BackgroundNoise.dat"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".dat",
                filetypes=[("DAT files", "*.dat"), ("All files", "*.*")],
                initialfile=default_filename
            )
            if save_path:
                try:
                    import shutil
                    shutil.copy2(local_dat_path, save_path)
                    messagebox.showinfo("成功", f"数据已保存到 {save_path}")
                except Exception as e:
                    messagebox.showerror("保存失败", f"保存数据时出错: {e}")
        
        btn_frame = tk.Frame(top_frame)
        btn_frame.pack(side=tk.TOP)
        
        tk.Button(btn_frame, text="保存图片", command=save_image, font=('SimHei', 16)).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="保存数据", command=save_data, font=('SimHei', 16)).pack(side=tk.LEFT, padx=10)
        
        timeout = 10.0
        poll_interval = 0.5
        waited = 0.0
        while not (os.path.exists(local_img_path) and os.path.getsize(local_img_path) > 0) and waited < timeout:
            time.sleep(poll_interval)
            waited += poll_interval

        if not os.path.exists(local_img_path) or os.path.getsize(local_img_path) == 0:
            msg = f"图片尚未同步到电脑（等待{timeout}s未出现）：{local_img_path}"
            tk.Label(win, text=msg, fg="red", wraplength=700, justify='left').pack(padx=8, pady=8)
            self.log(f"[显示] {msg}")
        else:
            try:
                img = Image.open(local_img_path)
                img = img.resize((800, 600))
                photo = ImageTk.PhotoImage(img, master=win)
                label = tk.Label(win, image=photo)
                label.image = photo
                label.pack()
            except Exception as e:
                tk.Label(win, text=f"图片加载失败: {e}", fg="red").pack()

# -------------------------
# TestRunner
# -------------------------
class TestRunner:
    def __init__(self, log_func=default_logger):
        self.log = log_func
        self._stop = False

    def stop(self):
        self._stop = True
        self.log("[Runner] 停止信号已设置")

    # [修改] 增加 ip_address 参数
    def run_rin(self, ra: RinAnalyzer, ui_root: tk.Tk, ip_address: str):
        try:
            try:
                self.log("[初始化] 正在清空共享文件夹和仪器内部文件夹...")

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

                try:
                    # [修改] 使用传入的 ip_address 进行清理连接
                    rm = pyvisa.ResourceManager()
                    inst = rm.open_resource(f"TCPIP0::{ip_address}::5025::SOCKET")
                    inst.write("MMEM:MDIR 'C:\\PTS\\Rin'")
                    inst.write("MMEM:DEL 'C:\\PTS\\Rin\\*.*'")
                    inst.close()
                    rm.close()
                except Exception as e:
                    self.log(f"[警告] 仪器文件夹清理失败: {e}")

                self.log("[初始化] 文件夹清理完成。")
            except Exception as e:
                self.log(f"[错误] 清理文件夹时出错: {e}")
            ra.ui_root = ui_root
            ra.log = self.log

            # [修改] 传递 IP 给 connect
            if ra.connect(ip_address=ip_address):
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
                try:
                    ra.close()
                except Exception:
                    pass
            else:
                self.log("[测试] 无法连接到仪器，RIN 测试终止")

            if self._stop or ra.stop_flag:
                def _notify_stopped():
                    if ra.stop_window and ra.stop_window.winfo_exists():
                        ra.stop_window.destroy()
                        ra.stop_window = None
                try:
                    ra.ui_root.after(0, _notify_stopped)
                except Exception:
                    pass
                return

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

    # [修改] 增加 ip_address 参数
    def run_background(self, bna: BackgroundNoiseAnalyzer, ui_root: tk.Tk, is_seedlight: bool, ip_address: str):
        try:
            bna.log = self.log
            # [修改] 传递 IP 给 connect
            if bna.connect(ip_address=ip_address):
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
# GUI
# -------------------------
class RinGUI:
    def __init__(self, parent=None):
        self.parent = parent
        
        if parent is None:
            self.root = tk.Tk()
            self.root.title("Rin_FSV3004 - 独立模式")
            self.root.geometry("1170x630")
            self.root.resizable(True, True)
            try:
                self.root.iconbitmap(r'PreciLasers.ico')
            except:
                pass
            if hasattr(self, 'set_center'):
                self.set_center(1170, 330) 
        else:
            self.root = parent

        self.params = {
            "osa_ip": "192.168.7.10", # 这里只是界面默认显示值，不再硬编码到逻辑中
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
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        left_frame = tk.Frame(main_container)
        left_frame.grid(row=0, column=0, sticky="n", padx=(0, 5))
        
        param_frame = tk.LabelFrame(left_frame, text="参数设置", padx=10, pady=10)
        param_frame.pack(fill=tk.X, padx=0, pady=0)

        self._add_param_entry(param_frame, "osa_ip", "IP地址:", self.params["osa_ip"], row=0)
        self._add_param_entry(param_frame, "dc_value", "DC值:", "2.4", row=1)
        self._add_param_entry(param_frame, "save_path", "保存路径:", self.params["save_path"], row=2)

        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=0, pady=8)
        
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')
        
        first_row_frame = tk.Frame(inner_btn_frame)
        first_row_frame.pack(fill=tk.X, pady=(0, 6))
        
        second_row_frame = tk.Frame(inner_btn_frame)
        second_row_frame.pack(fill=tk.X)
        
        self.btn_rin = tk.Button(first_row_frame, text="测RIN", command=self.start_rin, bg="#28862B", fg="#FFFFFF", width=10)
        self.btn_bg = tk.Button(first_row_frame, text="测底噪", command=self.start_background, bg="#28862B", fg="#FFFFFF", width=10)
        self.btn_seed = tk.Button(first_row_frame, text="种子光", command=self.start_seedlight, bg="#28862B", fg="#FFFFFF", width=10)
        self.btn_connect = tk.Button(second_row_frame, text="连接", command=self.connect_instrument, bg="#1D74C0", fg="#FFFFFF", width=10)
        self.btn_stop = tk.Button(second_row_frame, text="停止", command=self.stop_running, bg="#f44336", fg="#FFFFFF", width=10)
        self.btn_rename = tk.Button(second_row_frame, text="改名", command=self.rename_files, bg="#FF9800", fg="#FFFFFF", width=10)
        
        first_row_spacer = tk.Label(first_row_frame)
        first_row_spacer.pack(side=tk.LEFT, expand=True)
        self.btn_rin.pack(side=tk.LEFT, padx=6)
        self.btn_bg.pack(side=tk.LEFT, padx=6)
        self.btn_seed.pack(side=tk.LEFT, padx=6)
        first_row_spacer2 = tk.Label(first_row_frame)
        first_row_spacer2.pack(side=tk.LEFT, expand=True)
        
        second_row_spacer = tk.Label(second_row_frame)
        second_row_spacer.pack(side=tk.LEFT, expand=True)
        self.btn_connect.pack(side=tk.LEFT, padx=6)
        self.btn_stop.pack(side=tk.LEFT, padx=6)
        self.btn_rename.pack(side=tk.LEFT, padx=6)
        second_row_spacer2 = tk.Label(second_row_frame)
        second_row_spacer2.pack(side=tk.LEFT, expand=True)

        log_frame = tk.LabelFrame(main_container, text="运行日志", padx=5, pady=5)
        log_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.log_box = tk.Text(log_frame, wrap=tk.WORD)
        self.log_box.pack(fill=tk.BOTH, expand=True)
        
        main_container.grid_columnconfigure(0, weight=0)
        main_container.grid_columnconfigure(1, weight=1)
        main_container.grid_rowconfigure(0, weight=1)

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
        print(f"{t} {msg}")

    def get_params(self) -> Dict[str, Any]:
        p = {}
        try:
            p["osa_ip"] = self.entries["osa_ip"].get().strip()
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

    # [修改] 使用界面获取的 IP
    def connect_instrument(self):
        try:
            self.log("[连接] 正在尝试连接仪器...")
            ra = RinAnalyzer(log_func=self.log)
            # 从界面获取 IP
            ip = self.entries["osa_ip"].get().strip()
            port = 5025
            success = ra.connect(ip, port)
            if success:
                self.log(f"[连接] 成功连接到频谱仪: {ip}")
                # 连接测试只需开一下即关
                ra.close()
            else:
                self.log(f"[连接] 无法连接仪器: {ip}")
        except Exception as e:
            self.log(f"[连接] 失败: {e}")

    # [修改] 传递 IP 到 run_rin
    def start_rin(self):
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()
        # 获取 IP
        target_ip = p.get("osa_ip", "192.168.7.10")

        try:
            dc_input = float(self.entries["dc_value"].get())
            dc_for_ra = dc_input / 2.0
            self.log(f"[参数] DC 输入值 = {dc_input:.2f}，内部使用值 = {dc_for_ra:.2f}，目标IP = {target_ip}")
        except Exception:
            messagebox.showwarning("警告", "DC 值输入无效，将使用默认 2.40V")
            dc_for_ra = 1.20

        ra = RinAnalyzer(log_func=self.log)
        ra.dc_value = dc_for_ra
        try:
            ra.save_path = p.get("save_path") or self.params.get("save_path")
        except Exception:
            ra.save_path = None
        ra.ui_root = self.root
        ra.stop_flag = False

        def target():
            try:
                self.running_task = "rin"
                self.btn_rin.config(state=tk.DISABLED)
                self.btn_bg.config(state=tk.DISABLED)
                self.btn_connect.config(state=tk.DISABLED)
                self.btn_stop.config(state=tk.NORMAL)
                self.runner._stop = False
                # 传递 ip_address
                self.runner.run_rin(ra, self.root, ip_address=target_ip)
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

    # [修改] 传递 IP 到 run_background
    def start_background(self):
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()
        target_ip = p.get("osa_ip", "192.168.7.10")
        
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
                # 传递 ip_address
                self.runner.run_background(bna, self.root, is_seedlight=False, ip_address=target_ip)
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
    
    # [修改] 传递 IP 到 run_background
    def start_seedlight(self):
        if self.running_task:
            messagebox.showwarning("警告", "已有任务在运行")
            return
        p = self.get_params()
        target_ip = p.get("osa_ip", "192.168.7.10")
        
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
                # 传递 ip_address
                self.runner.run_background(bna, self.root, is_seedlight=True, ip_address=target_ip)
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
        self.runner.stop()
        self.log("[主] 停止命令已发送给后台任务")
        self.btn_stop.config(state=tk.DISABLED)

    def rename_files(self):
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

if __name__ == "__main__":
    gui = RinGUI()
    gui.run()