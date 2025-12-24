#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CT Tuning GUI - Robust version with second-group current sweep (user request)
- Keeps GUI layout and structure
- First group: temperature sweep at specified current (as before)
- Second group: current sweep from group2_start_mA down to group2_stop_mA by group2_step_mA,
  recording main wavelength at each current, saving spectra and summary, and plotting wl vs current
Requirements:
    pip install pyvisa numpy matplotlib pillow pywinauto
"""
from __future__ import annotations

import os
import time
import threading
import csv
import struct
import traceback
from typing import Tuple, Optional, Any, Dict, List

import pyvisa
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog
import matplotlib
import matplotlib.ticker as mticker

# 新增导入PIL库
from PIL import Image, ImageDraw, ImageFont, ImageTk


matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
# 设置中文字体
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题
# pywinauto guarded import
try:
    from pywinauto.application import Application
    from pywinauto import timings
    PYW_AVAILABLE = True
except Exception:
    Application = None
    timings = None
    PYW_AVAILABLE = False

# -------------------------
# Helpers
# -------------------------
def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)
    return path

# -------------------------
# LaserController (same)
# -------------------------
class LaserController:
    def __init__(self, exe_path: str = r"C:\PTS\CT_W\Preci_Semi\Preci-Seed.exe",
                 window_title: str = r"Preci-Semi-Seed", log_func=print):
        self.exe_path = exe_path
        self.window_title = window_title
        self.app = None
        self.win = None
        self.log = log_func
        self.connected = False

    def connect(self, timeout: float = 10.0, attach_timeout: float = 3.0) -> bool:
        if not PYW_AVAILABLE:
            raise RuntimeError("pywinauto 未安装或不可用，无法控制激光器上位机。")
        if self.connected and self.win is not None:
            try:
                if self.win.exists() and self.win.is_visible():
                    self.log("[Laser] 已复用现有窗口句柄")
                    return True
            except Exception:
                self.log("[Laser] 现有句柄无效，重新连接")
                self.app = None
                self.win = None
                self.connected = False
        try:
            self.log("[Laser] 尝试附加到运行中的 Preci-Semi-Seed 窗口...")
            self.app = Application(backend="uia").connect(title_re=self.window_title, timeout=attach_timeout)
            self.win = self.app.window(title_re=self.window_title)
            timings.wait_until_passes(3, 0.5, lambda: self.win.exists() and self.win.is_visible())
            self.win.set_focus()
            self.connected = True
            self.log("[Laser] 附加成功")
            return True
        except Exception as e_attach:
            self.log(f"[Laser] 未找到运行实例: {e_attach}，尝试启动 exe：{self.exe_path}")
        try:
            self.log("[Laser] 启动 exe...")
            self.app = Application(backend="uia").start(cmd_line=f'"{self.exe_path}"')
            start_time = time.time()
            deadline = start_time + timeout
            while time.time() < deadline:
                try:
                    self.app.connect(title_re=self.window_title, timeout=1)
                    self.win = self.app.window(title_re=self.window_title)
                    timings.wait_until_passes(5, 0.5, lambda: self.win.exists() and self.win.is_visible())
                    self.win.set_focus()
                    self.connected = True
                    self.log("[Laser] 启动并连接成功")
                    return True
                except Exception:
                    time.sleep(0.3)
            raise RuntimeError("启动后未检测到窗口（超时）")
        except Exception as e_start:
            self.log(f"[Laser][错误] 启动或连接失败: {e_start}")
            self.connected = False
            raise

    def get_current_mA(self) -> Optional[float]:
        try:
            ctrl = self.win.child_window(auto_id="Label_current", control_type="Text")
            txt = ctrl.window_text()
            return float(txt)
        except Exception as e:
            self.log(f"[Laser] 读取电流失败: {e}")
            return None

    def set_current_mA(self, val_mA: float):
        try:
            edit = self.win.child_window(auto_id="textBox_Current", control_type="Edit")
            edit.set_edit_text(f"{val_mA:.2f}")
            btn = self.win.child_window(title="Set", control_type="Button")
            btn.click()
            self.log(f"[Laser] 已设置电流: {val_mA:.2f} mA")
        except Exception as e:
            self.log(f"[Laser] 设置电流失败: {e}")
            raise

    def get_temperature_C(self) -> Optional[float]:
        try:
            ctrl = self.win.child_window(auto_id="Label_Temperature", control_type="Text")
            txt = ctrl.window_text()
            return float(txt)
        except Exception as e:
            self.log(f"[Laser] 读取温度失败: {e}")
            return None

    def set_temperature_C(self, val_C: float):
        try:
            edit = self.win.child_window(auto_id="TextBox_Temperature", control_type="Edit")
            edit.set_edit_text(f"{val_C:.2f}")
            btn = self.win.child_window(title="Set", control_type="Button")
            btn.click()
            self.log(f"[Laser] 已设置温度: {val_C:.2f} °C")
        except Exception as e:
            self.log(f"[Laser] 设置温度失败: {e}")
            raise

# -------------------------
# OSAController (robust)
# -------------------------
class OSAController:
    def __init__(self, resource: str, log_func=print):
        self.rm = pyvisa.ResourceManager()
        self.inst = None
        self.resource = resource
        self.log = log_func
        self.timeout = 20000
        self.retries = 2

    def connect(self):
        try:
            self.inst = self.rm.open_resource(self.resource)
            self.inst.timeout = max(self.timeout, 30000)
            self.log(f"[OSA] 已连接: {self.resource}")
        except Exception as e:
            self.log(f"[OSA] 连接失败: {e}")
            raise

    def query_idn(self) -> str:
        try:
            return self.inst.query("*IDN?").strip()
        except Exception as e:
            self.log(f"[OSA] *IDN? 失败: {e}")
            return ""

    def query_format(self) -> str:
        try:
            return self.inst.query(":FORMat:DATA?").strip().upper()
        except Exception as e:
            self.log(f"[OSA] :FORMat:DATA? 失败: {e}")
            return ""

    def query_x_axis(self, trace: Optional[str] = None) -> Optional[np.ndarray]:
        """
        尝试从仪器读取 X 轴（波长轴）。不同仪器命令不同，按顺序尝试几种常见命令。
        返回 np.ndarray 或 None。
        """
        if self.inst is None:
            self.log("[OSA] 未连接，无法读取 X 轴")
            return None
        t = trace or self.query_active_trace()
        cmds = [
            f":TRACe:DATA:X? {t}",
            f":TRACe:X? {t}",
            ":TRACe:DATA:X?",
            ":TRACe:X?",
            ":SENSE:WAVELENGTH:DATA?",
            ":SENSE:WAV:DATA?",
        ]
        last_errs = []
        for cmd in cmds:
            try:
                self.log(f"[OSA] 尝试读取 X 轴 (cmd='{cmd}')")
                # 优先使用 query_ascii_values（返回数值列表）
                try:
                    vals = self.inst.query_ascii_values(cmd)
                    if vals and len(vals) > 0:
                        arr = np.array(vals, dtype=float)
                        self.log(f"[OSA] X 轴 ASCII 返回, pts={len(arr)} (cmd='{cmd}')")
                        return arr
                except Exception as e_ascii:
                    last_errs.append((cmd, str(e_ascii)))
                    # 继续尝试以纯文本方式读取
                # 退回到文本读取并解析
                try:
                    resp = self.inst.query(cmd).strip()
                    if resp:
                        tokens = [tok.strip() for tok in resp.replace('\r', '').replace('\n', ',').split(',') if tok.strip() != ""]
                        vals = [float(tok) for tok in tokens]
                        arr = np.array(vals, dtype=float)
                        self.log(f"[OSA] X 轴 raw ascii 返回, pts={len(arr)} (cmd='{cmd}')")
                        return arr
                except Exception as e_txt:
                    last_errs.append((cmd, str(e_txt)))
                    continue
            except Exception as e:
                last_errs.append((cmd, str(e)))
                continue
        self.log(f"[OSA] 未能从仪器读取 X 轴，尝试的命令返回错误: {last_errs}")
        return None

    def query_active_trace(self) -> str:
        try:
            t = self.inst.query(":TRACe:ACTive?").strip()
            return t if t else "TRA"
        except Exception as e:
            self.log(f"[OSA] :TRACe:ACTive? 失败: {e}")
            return "TRA"

    def query_trace_sample_count(self, trace: Optional[str] = None) -> Optional[int]:
        try:
            t = trace or self.query_active_trace()
            resp = self.inst.query(f":TRACe:DATA:SNUMber? {t}").strip()
            return int(float(resp))
        except Exception:
            try:
                resp = self.inst.query(":TRACe:DATA:SNUMber?").strip()
                return int(float(resp))
            except Exception as e:
                self.log(f"[OSA] :TRACe:DATA:SNUMber? 失败: {e}")
                return None

    def _try_query_float(self, cmd_list: List[str]) -> Optional[float]:
        for cmd in cmd_list:
            try:
                resp = self.inst.query(cmd).strip()
                if resp == "":
                    continue
                token = resp.split()[0].replace(",", "")
                return float(token)
            except Exception:
                continue
        return None

    def _build_wavelength_axis(self, npoints: int) -> np.ndarray:
        start_cmds = [":SENSE:WAVELENGTH:START?", ":SENSE:WAV:STAR?", ":SENSE:WAV:START?"]
        stop_cmds = [":SENSE:WAVELENGTH:STOP?", ":SENSE:WAV:STOP?"]
        start = self._try_query_float(start_cmds)
        stop = self._try_query_float(stop_cmds)
        if start is not None and stop is not None and npoints > 1:
            if abs(start) < 1.0 and abs(stop) < 1.0:
                start_nm = start * 1e9
                stop_nm = stop * 1e9
            else:
                start_nm = start
                stop_nm = stop
            return np.linspace(start_nm, stop_nm, npoints)
        center_cmds = [":SENSE:WAVELENGTH:CENTER?", ":SENSE:WAV:CENTER?"]
        span_cmds = [":SENSE:WAVELENGTH:SPAN?", ":SENSE:WAV:SPAN?"]
        center = self._try_query_float(center_cmds)
        span = self._try_query_float(span_cmds)
        if center is not None and span is not None and npoints > 1:
            if abs(center) < 1.0:
                center_nm = center * 1e9
                span_nm = span * 1e9
            else:
                center_nm = center
                span_nm = span
            half = span_nm / 2.0
            return np.linspace(center_nm - half, center_nm + half, npoints)
        try:
            pts = self.query_trace_sample_count()
            if pts and pts == npoints:
                return np.linspace(0.0, float(npoints - 1), npoints)
        except Exception:
            pass
        return np.arange(npoints).astype(float)

    def fetch_trace(self) -> Tuple[np.ndarray, np.ndarray]:
        if self.inst is None:
            raise RuntimeError("OSA 未连接")
        trace = self.query_active_trace() or "TRA"
        fmt = self.query_format() or ""
        cmd = f":TRACe:DATA:Y? {trace}"
        last_errs = []

        def try_ascii():
            try:
                vals = self.inst.query_ascii_values(cmd)
                if vals and len(vals) > 0:
                    self.log(f"[OSA] ASCII 读取成功 {len(vals)} 点 (cmd='{cmd}')")
                    return np.array(vals, dtype=float)
            except Exception as e:
                last_errs.append(("ascii", str(e)))
                self.log(f"[OSA] ASCII 读取失败: {e}")
            return None

        def try_binary(is_big_endian: bool, datatype: str = 'f'):
            try:
                orig_to = self.inst.timeout
                self.inst.timeout = max(orig_to, self.timeout * 2)
                vals = self.inst.query_binary_values(cmd, datatype=datatype, is_big_endian=is_big_endian)
                self.inst.timeout = orig_to
                if vals and len(vals) > 0:
                    self.log(f"[OSA] Binary 读取成功 {len(vals)} 点 (big_endian={is_big_endian})")
                    return np.array(vals, dtype=float)
            except Exception as e:
                last_errs.append((f"bin_be={is_big_endian}", str(e)))
                self.log(f"[OSA] Binary 读取失败 (big_endian={is_big_endian}): {e}")
            return None

        if "ASCII" in fmt or fmt == "":
            arr = try_ascii()
            if arr is not None:
                w = self._build_wavelength_axis(len(arr))
                if np.max(np.abs(w)) < 1.0:
                    w = w * 1e9
                return w, arr

        arr = try_binary(False, 'f') or try_binary(True, 'f')
        if arr is not None:
            w = self._build_wavelength_axis(len(arr))
            if np.max(np.abs(w)) < 1.0:
                w = w * 1e9
            return w, arr

        try:
            raw = self.inst.read_raw()
            if raw is None or len(raw) == 0:
                raise RuntimeError("read_raw returned empty")
            if raw.startswith(b'#'):
                ndig = int(chr(raw[1]))
                length_bytes = raw[2:2 + ndig]
                length = int(length_bytes.decode())
                data_bytes = raw[2 + ndig:2 + ndig + length]
                if length % 4 == 0:
                    count = length // 4
                    try:
                        vals = struct.unpack('<' + 'f' * count, data_bytes)
                        arr = np.array(vals, dtype=float)
                        w = self._build_wavelength_axis(len(arr))
                        if np.max(np.abs(w)) < 1.0:
                            w = w * 1e9
                        self.log(f"[OSA] raw '#' 解析成功 (little-endian), pts={len(arr)}")
                        return w, arr
                    except Exception:
                        try:
                            vals = struct.unpack('>' + 'f' * count, data_bytes)
                            arr = np.array(vals, dtype=float)
                            w = self._build_wavelength_axis(len(arr))
                            if np.max(np.abs(w)) < 1.0:
                                w = w * 1e9
                            self.log(f"[OSA] raw '#' 解析成功 (big-endian), pts={len(arr)}")
                            return w, arr
                        except Exception as e2:
                            raise RuntimeError(f"raw '#' 数据解析失败: {e2}")
                else:
                    raise RuntimeError("raw '#' 数据长度不是 float32 的整数倍")
            else:
                txt = raw.decode(errors='ignore').strip()
                tokens = [t.strip() for t in txt.replace('\r', '').replace('\n', ',').split(',') if t.strip() != ""]
                vals = [float(t) for t in tokens]
                arr = np.array(vals, dtype=float)
                w = self._build_wavelength_axis(len(arr))
                if np.max(np.abs(w)) < 1.0:
                    w = w * 1e9
                self.log(f"[OSA] raw ascii 解析成功, pts={len(arr)}")
                return w, arr
        except Exception as e:
            last_errs.append(("raw", str(e)))
            self.log(f"[OSA] raw 读取解析失败: {e}")

        raise RuntimeError(f"无法读取 OSA trace。尝试记录: {last_errs}")

    def sweep_and_fetch(self) -> Tuple[np.ndarray, np.ndarray]:
        try:
            try:
                self.inst.write(":INIT:CONT OFF")
            except Exception:
                pass
            self.inst.write(":INIT")
            self.inst.query("*OPC?")
        except Exception as e:
            self.log(f"[OSA] 触发扫描失败: {e}")
            raise
        return self.fetch_trace()

# -------------------------
# TestRunner
# -------------------------
class TestRunner:
    def __init__(self, laser: Optional[LaserController], osa: OSAController, log_func=print):
        self.laser = laser
        self.osa = osa
        self.log = log_func
        self._stop = False

    def stop(self):
        self._stop = True
        self.log("[Runner] 停止信号已设置")

    def _float_range(self, start: float, stop: float, step: float) -> List[float]:
        if step == 0:
            raise ValueError("step cannot be 0")
        out = []
        t = start
        step_magnitude = abs(step)
        # 根据start和stop的关系决定是递增还是递减
        if start < stop:
            # 递增：从start到stop，每次加step_magnitude
            while t <= stop + 1e-9:
                out.append(round(t, 6))
                t += step_magnitude
        else:
            # 递减：从start到stop，每次减step_magnitude
            while t >= stop - 1e-9:
                out.append(round(t, 6))
                t -= step_magnitude
        return out

    def _save_spectrum(self, wavelengths: np.ndarray, powers: np.ndarray, save_path: str, prefix: str) -> str:
        if os.path.isdir(save_path) or save_path.endswith(os.sep):
            out_dir = save_path
        else:
            out_dir = os.path.dirname(save_path) or "."
        ensure_dir(out_dir)
        filename = os.path.join(out_dir, f"{prefix}_{time.strftime('%Y%m%d_%H%M%S')}.csv")
        with open(filename, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Wavelength_nm", "Power"])
            for x, y in zip(wavelengths, powers):
                # 波长保留小数点后 4 位，功率格式保持原样
                w.writerow([f"{float(x):.4f}", f"{float(y):.6f}"])
        self.log(f"[Runner] 保存光谱: {filename}")
        return filename

    def _append_summary(self, save_path: str, current_mA: float, temperature: Optional[float], main_wl: float, spectrum_file: str, test_group: int = 0, summary_filename: str = None):
        if os.path.isdir(save_path) or save_path.endswith(os.sep):
            out_dir = save_path
        else:
            out_dir = os.path.dirname(save_path) or "."
        ensure_dir(out_dir)
        # 确定汇总文件名的优先级：传入的文件名 > 默认的组文件名 > 通用文件名
        if summary_filename:
            # 添加自动追加.csv后缀的逻辑
            if not summary_filename.lower().endswith('.csv'):
                summary_filename += '.csv'
            summary_fn = os.path.join(out_dir, summary_filename)
        elif test_group == 1:
            summary_fn = os.path.join(out_dir, "Test1_summary.csv")
        elif test_group == 2:
            summary_fn = os.path.join(out_dir, "Test2_summary.csv")
        else:
            # 保持原有命名逻辑作为默认
            summary_fn = os.path.join(out_dir, f"ct_tuning_summary_{time.strftime('%Y%m%d')}.csv")
            
        header_needed = not os.path.exists(summary_fn)
        with open(summary_fn, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if header_needed:
                # 测试1和测试2的CSV都去掉Timestamp和SpectrumFile两列
                w.writerow(["Current_mA", "Temperature_C", "MainWavelength_nm"])
            temp_str = f"{temperature:.2f}" if temperature is not None else "N/A"
            # 测试1和测试2的数据行都只保留3列（主波长保留 4 位）
            w.writerow([f"{current_mA:.2f}", temp_str, f"{main_wl:.4f}"])
        
    def _compute_peak_wavelength(self, wavelengths: np.ndarray, powers: np.ndarray) -> float:
        """
        改进版主波长计算：
        使用二次插值法拟合峰值，提升波长精度（原方法仅取采样点）
        """
        if len(powers) == 0:
            return float("nan")

        # 找到最大功率点索引
        idx = int(np.nanargmax(powers))
        if idx <= 0 or idx >= len(powers) - 1:
            # 边界情况直接返回采样点
            return float(wavelengths[idx])

        # 三点抛物线拟合（x: wavelength, y: power）
        x1, x2, x3 = wavelengths[idx - 1], wavelengths[idx], wavelengths[idx + 1]
        y1, y2, y3 = powers[idx - 1], powers[idx], powers[idx + 1]

        # 抛物线顶点位置计算（参考二次曲线拟合公式）
        denom = (y1 - 2 * y2 + y3)
        if abs(denom) < 1e-15:
            return float(x2)  # 避免除零

        delta = 0.5 * (y1 - y3) / denom
        wl_peak = x2 + delta * (x3 - x1) / 2

        return float(wl_peak)

    
    def _plot_xy_curve(self, x, y, xlabel, ylabel, title, out_dir, prefix, invert_x=False, save_csv=False, extra_cols=None):
        """
        通用绘图函数
        """
        ensure_dir(out_dir)
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        fig_path = os.path.join(out_dir, f"{prefix}_{timestamp}.png")

        # 绘制曲线
        plt.figure(figsize=(20, 10))
        plt.plot(x, y, marker='o', linestyle='-', linewidth=2)
        if invert_x:
            plt.gca().invert_xaxis()
        plt.xlabel(xlabel, fontsize=20)
        plt.ylabel(ylabel, fontsize=20)
        plt.title(title, fontsize=22)

        # 强制y轴不用科学计数法
        ax = plt.gca()
        ax.ticklabel_format(style='plain', axis='y')
        ax.yaxis.get_major_formatter().set_scientific(False)
        ax.yaxis.get_major_formatter().set_useOffset(False)
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.3f'))
        
        ax.xaxis.get_major_formatter().set_scientific(False)
        ax.xaxis.get_major_formatter().set_useOffset(False)
        # 设置刻度字体大小
        plt.xticks(fontsize=16)
        plt.yticks(fontsize=16)
        # 设置网格线 (可选)
        plt.grid(True, linestyle='--', alpha=0.7, which='major')
        ax.minorticks_on()
        ax.grid(True, axis='x', linestyle=':', alpha=0.5, which='minor')
        # 设置x轴刻度步进
        if "Temperature" in xlabel or "group1" in prefix:
            # 图一：温度步进为1
            x_min, x_max = min(x), max(x)
            plt.xticks(np.arange(round(x_min), round(x_max) + 1, 1))
        elif "Current" in xlabel or "group2" in prefix:
            # 图二：电流步进为10
            x_min, x_max = min(x), max(x)
            plt.xticks(np.arange(round(x_min), round(x_max) + 5, 5))

        # 设置 y 轴刻度为数据点（或按数据间距）
        #plt.yticks(sorted(set(np.round(y, 3))))
        plt.tight_layout()
        plt.savefig(fig_path, dpi=300)
        plt.close()
        self.log(f"[Runner] 图像保存到 {fig_path}")

        # 返回保存的图像路径
        return fig_path

    # 保留原有的run_manual_two_groups方法，但不自动连续执行两组测试
    def run_manual_two_groups(self, start_temp: float, end_temp: float, step: float, save_path: str = "./data", 
                           group2_start_mA: float = 400.0, group2_step_mA: float = 5.0, 
                           group2_stop_mA: float = 0.5, group2_temp_C: float = 25.0):
        """
        注意：此方法已不连续执行两组测试，仅作为兼容性保留
        请使用单独的run_group1和run_group2方法
        """
        self.log("[Runner] 注意：run_manual_two_groups方法已不连续执行两组测试")
        self.log("[Runner] 请使用单独的开始按钮控制每组测试")

    def run_group1(self, start_temp: float, end_temp: float, step: float, save_path: str = "./data", delay_s: float = 0.8, summary_filename: str = None, current_mA: float = None):
        """
        Group1: temperature sweep at current = GUI current_mA
        """
        self._stop = False

        try:
            # 确定保存目录
            if os.path.isdir(save_path) or save_path.endswith(os.sep):
                out_dir = save_path
            else:
                out_dir = os.path.dirname(save_path) or "."
            ensure_dir(out_dir)
            
            # 确保文件名包含.csv扩展名
            if summary_filename:
                if not summary_filename.lower().endswith('.csv'):
                    summary_filename += '.csv'
                file_path = os.path.join(out_dir, summary_filename)
            else:
                file_path = os.path.join(out_dir, "Test1_summary.csv")
            
            # 检查文件是否存在，如果存在则删除
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    self.log(f"[Runner] 已删除同名文件: {file_path}")
                except Exception as e:
                    self.log(f"[Runner] 删除文件失败: {e}")
            
            # 优先使用传入的电流值，否则读取当前电流
            current_for_temp = 360.0
            if current_mA is not None:
                current_for_temp = current_mA
                # 如果有激光控制器，尝试设置电流
                if self.laser:
                    try:
                        self.laser.set_current_mA(current_for_temp)
                        self.log(f"[Runner] 已设置电流为 {current_for_temp} mA")
                        # 等待电流稳定
                        time.sleep(1.0)
                    except Exception as e:
                        self.log(f"[Runner] 设置电流失败: {e}")
                        # 设置失败时读取当前电流
                        val = self.laser.get_current_mA()
                        if val is not None:
                            current_for_temp = val
            elif self.laser:
                val = self.laser.get_current_mA()
                if val is not None:
                    current_for_temp = val
            temps = self._float_range(start_temp, end_temp, step)
            self.log(f"[Runner] 组1: 电流 {current_for_temp} mA 温度扫描 {start_temp}->{end_temp} step {step} 共 {len(temps)} 步，稳定时间 {delay_s} 秒")
            # 添加温度稳定检测参数
            stability_threshold = 0.1  # 稳定阈值，摄氏度
            max_wait_time = delay_s * 5  # 最大等待时间
            check_interval = 0.5  # 检查间隔
            
            for t in temps:
                if self._stop:
                    self.log("[Runner] 收到停止信号，结束组1")
                    break
                if self.laser:
                    try:
                        self.laser.set_temperature_C(t)
                        # 新增：等待温度稳定
                        self.log(f"[Runner] 设置温度为 {t}°C，等待稳定...")
                        wait_time = 0
                        stable = False
                        
                        # 先等待一段时间让温度开始变化
                        time.sleep(delay_s * 0.5)
                        
                        # 循环检查温度是否稳定
                        while wait_time < max_wait_time and not stable and not self._stop:
                            current_temp = self.laser.get_temperature_C()
                            if current_temp is not None:
                                temp_diff = abs(current_temp - t)
                                self.log(f"[Runner] 当前温度: {current_temp:.2f}°C, 目标: {t:.2f}°C, 差值: {temp_diff:.2f}°C")
                                
                                if temp_diff <= stability_threshold:
                                    stable = True
                                    self.log(f"[Runner] 温度已稳定在 {t}°C")
                                else:
                                    time.sleep(check_interval)
                                    wait_time += check_interval
                            else:
                                # 无法读取温度时，退化为简单延时
                                time.sleep(check_interval)
                                wait_time += check_interval
                        
                        if not stable and not self._stop:
                            self.log(f"[Runner] 温度在 {max_wait_time}s 内未完全稳定，继续测量")
                    except Exception as e:
                        self.log(f"[Runner] 设置温度失败: {e}")
                        # 设置失败时也等待一段时间
                        time.sleep(delay_s)
                else:
                    # 未连接激光控制器时，使用简单延时
                    time.sleep(delay_s)
                try:
                    wavelengths, powers = self.osa.sweep_and_fetch()
                except Exception as e:
                    self.log(f"[Runner] 组1 OSA 读取失败 (temp {t}°C): {e}")
                    continue
                main_wl = self._compute_peak_wavelength(wavelengths, powers)
                try:
                    self._append_summary(save_path, current_for_temp, t, main_wl, "", test_group=1, summary_filename=summary_filename)
                    self.log(f"[Runner] 组1 {current_for_temp}mA, {t:.2f}°C -> 主波长 {main_wl:.4f} nm")
                except Exception as e:
                    self.log(f"[Runner] 组1 写入汇总失败: {e}")
        except Exception as e:
            self.log(f"[Runner] 组1 出错: {e}")

        self.log("[Runner] 组1流程完成")

    def plot_group1_wavelength_vs_temperature(self, out_dir, summary_filename=None):
        try:
            filename = summary_filename if summary_filename else "Test1_summary.csv"
            # 修复：自动处理没有.csv扩展名的情况
            if not filename.endswith('.csv'):
                filename += '.csv'
            file_path = os.path.join(out_dir, filename)
            if not os.path.exists(file_path):
                self.log(f"[Runner] {filename} 文件不存在: {file_path}")
                return

            temps, wavelengths = [], []
            with open(file_path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader)
                self.log(f"[Runner] 读取到文件头: {header}")

                # 🚀 新格式：3列 [Current_mA, Temperature_C, MainWavelength_nm]
                for row in reader:
                    try:
                        temp = float(row[1])
                        wl = float(row[2])
                        if wl > 200:   # 波长大于200nm才算有效
                            temps.append(temp)
                            wavelengths.append(wl)
                    except Exception as e:
                        self.log(f"[Runner] 跳过无效行 {row}: {e}")
                        continue

            if temps:
                uniq = {}
                for t, wl in zip(temps, wavelengths):
                    uniq[t] = wl  # 保留最后一次的测量结果
                temps = sorted(uniq.keys(), reverse=True)
                wavelengths = [uniq[t] for t in temps]

                return self._plot_xy_curve(
                    temps, wavelengths,
                    xlabel="温度(°C)", ylabel="波长(nm)",
                    title=f"{self.laser.get_current_mA() if self.laser else 360:.2f} mA下温度-波长关系",
                    out_dir=out_dir, prefix="温度波长关系图",
                    invert_x=True, save_csv=False
                )
            else:
                self.log("[Runner] 组1 没有采集到有效波长数据，请检查 CSV 内容")
                return None
        except Exception as e:
            self.log(f"[Runner] 组1 绘制失败: {e}")
            return None

    

    # 新增：单独运行第二组测试
    def run_group2(self, start_mA: float, step_mA: float, stop_mA: float, temp_C: float,
               save_path: str = "./data", delay_s: float = 0.6, summary_filename: str = None):
        """
        Group2: current sweep from start_mA down by step_mA to stop_mA,
                with temperature fixed at temp_C
        """
        self._stop = False
        
        try:
            # 确定保存目录
            if os.path.isdir(save_path) or save_path.endswith(os.sep):
                out_dir = save_path
            else:
                out_dir = os.path.dirname(save_path) or "."
            ensure_dir(out_dir)
            
            # 确保文件名包含.csv扩展名
            if summary_filename:
                if not summary_filename.lower().endswith('.csv'):
                    summary_filename += '.csv'
                file_path = os.path.join(out_dir, summary_filename)
            else:
                file_path = os.path.join(out_dir, "Test2_summary.csv")
            
            # 检查文件是否存在，如果存在则删除
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    self.log(f"[Runner] 已删除同名文件: {file_path}")
                except Exception as e:
                    self.log(f"[Runner] 删除文件失败: {e}")
            
            # 固定组2测试温度
            if self.laser:
                self.laser.set_temperature_C(temp_C)
                self.log(f"[Runner] 组2: 设置温度为 {temp_C:.2f} °C")
                
                # 新增：等待温度稳定
                # 添加温度稳定检测参数
                temp_stability_threshold = 0.1  # 温度稳定阈值，摄氏度
                temp_max_wait_time = delay_s * 5  # 最大等待时间
                temp_check_interval = 0.5  # 检查间隔
                
                self.log(f"[Runner] 等待温度稳定在 {temp_C:.2f}°C...")
                temp_wait_time = 0
                temp_stable = False
                
                # 先等待一段时间让温度开始变化
                time.sleep(delay_s * 0.5)
                
                # 循环检查温度是否稳定
                while temp_wait_time < temp_max_wait_time and not temp_stable and not self._stop:
                    current_temp = self.laser.get_temperature_C()
                    if current_temp is not None:
                        temp_diff = abs(current_temp - temp_C)
                        self.log(f"[Runner] 当前温度: {current_temp:.2f}°C, 目标: {temp_C:.2f}°C, 差值: {temp_diff:.2f}°C")
                        
                        if temp_diff <= temp_stability_threshold:
                            temp_stable = True
                            self.log(f"[Runner] 温度已稳定在 {temp_C:.2f}°C")
                        else:
                            time.sleep(temp_check_interval)
                            temp_wait_time += temp_check_interval
                    else:
                        # 无法读取温度时，退化为简单延时
                        time.sleep(temp_check_interval)
                        temp_wait_time += temp_check_interval
                
                if not temp_stable and not self._stop:
                    self.log(f"[Runner] 温度在 {temp_max_wait_time}s 内未完全稳定，继续测量")
        except Exception as e:
            self.log(f"[Runner] 组2: 设置温度失败 {e}")

        # 构造递减电流序列
        start_curr = float(start_mA)
        step_mag = abs(float(step_mA))
        stop_curr = float(stop_mA)
        if step_mag == 0:
            self.log("[Runner] group2_step_mA 不能为 0，已跳过组2")
            return
        currents = []
        c = start_curr
        while c >= stop_curr - 1e-9:
            currents.append(round(c, 6))
            c -= step_mag

        self.log(f"[Runner] 组2: 电流从 {start_curr}mA 每次 -{step_mag}mA 到 {stop_curr}mA，共 {len(currents)} 步，稳定时间 {delay_s} 秒")

        peaks_curr = []
        peaks_wl = []

        # 添加电流稳定检测相关参数
        stability_threshold = 1.0  # 电流稳定阈值，mA
        max_wait_time = delay_s * 3  # 最大等待时间
        check_interval = 0.3  # 检查间隔

        for cur in currents:
            if self._stop:
                self.log("[Runner] 收到停止信号，提前结束组2")
                break
            try:
                if self.laser:
                    try:
                        self.laser.set_current_mA(cur)
                        # 新增：等待电流稳定
                        self.log(f"[Runner] 设置电流为 {cur}mA，等待稳定...")
                        wait_time = 0
                        stable = False
                         
                        # 循环检查电流是否稳定
                        while wait_time < max_wait_time and not stable and not self._stop:
                            current_current = self.laser.get_current_mA()
                            if current_current is not None:
                                curr_diff = abs(current_current - cur)
                                self.log(f"[Runner] 当前电流: {current_current:.2f}mA, 目标: {cur:.2f}mA, 差值: {curr_diff:.2f}mA")
                                 
                                if curr_diff <= stability_threshold:
                                    stable = True
                                    self.log(f"[Runner] 电流已稳定在 {cur}mA")
                                else:
                                    time.sleep(check_interval)
                                    wait_time += check_interval
                            else:
                                # 无法读取电流时，退化为简单延时
                                time.sleep(check_interval)
                                wait_time += check_interval
                         
                        if not stable and not self._stop:
                            self.log(f"[Runner] 电流在 {max_wait_time}s 内未完全稳定，继续测量")
                    except Exception as e:
                        self.log(f"[Runner] 设置电流 {cur} mA 失败: {e}")
                        time.sleep(delay_s)  # 设置失败时也等待一段时间
                else:
                    self.log(f"[Runner] 未配置 LaserController，跳过设置电流 {cur} mA (仍会采集 OSA)")
                    time.sleep(delay_s)  # 未配置时使用简单延时

                time.sleep(delay_s * 0.5)  # 额外小延时，确保系统稳定

                try:
                    wavelengths, powers = self.osa.sweep_and_fetch()
                except Exception as e:
                    self.log(f"[Runner] 组2 OSA 读取失败 (current {cur} mA): {e}")
                    continue

                main_wl = self._compute_peak_wavelength(wavelengths, powers)
                try:
                    self._append_summary(save_path, cur, temp_C, main_wl, "",
                                        test_group=2, summary_filename=summary_filename)
                except Exception as e:
                    self.log(f"[Runner] 组2 写入汇总失败: {e}")

                peaks_curr.append(cur)
                peaks_wl.append(main_wl)
                self.log(f"[Runner] 组2 {int(cur)}mA @ {temp_C:.2f}°C -> 主波长 {main_wl:.4f} nm")

            except Exception as e:
                self.log(f"[Runner] 组2 电流 {cur} mA 处理失败: {e}")
                continue

        if peaks_curr:
            self._plot_xy_curve(
                peaks_curr, peaks_wl,
                xlabel="电流(mA)", ylabel="波长(nm)",
                title=f"{temp_C:.2f}°C下电流-波长关系",
                out_dir=save_path, prefix="电流波长关系图",
                invert_x=False, save_csv=False,
                extra_cols={"Temperature_C": [f"{temp_C:.2f}"] * len(peaks_curr)}
            )
        else:
            self.log("[Runner] 组2 没有采集到峰值数据，跳过作图")

# -------------------------
# GUI (with new group2 params)
# -------------------------
class CT_W_GUI:
    def __init__(self, parent=None):
        self.parent = parent
        
        # --- 核心修改：如果是集成模式，直接使用父控件作为 root ---
        if parent is None:
            self.root = tk.Tk()
            self.root.title("CT_P - 独立模式")
            # 假设 set_center() 只有在独立模式下需要
            if hasattr(self, 'set_center'):
                self.set_center(1510, 1090) 
            self.root.resizable(True, True)
            try:
                self.root.iconbitmap(r'PreciLasers.ico')
            except:
                pass
        else:
            self.root = parent # <--- 修改点：直接使用父 Frame

        # defaults (added group2 params)
        self.params = {
            "osa_ip": "192.168.7.21",
            "current_mA": 360.0,
            "t_start": 36.0,
            "t_stop": 15.0,
            "t_step": 1.0,
            "center_nm": 1550.0,
            "span_nm": 5.0,
            "laser_exe_path": r"C:\PTS\qijian\上位机软件\CT_W\Preci_Semi\Preci-Seed.exe",
            "save_path": r"C:\PTS\qijian\CT_W",
            # group2 specific
            "group2_temp_C": 25.0,           # 新增：第二组测试前设置的温度
            "group2_start_mA": 400.0,
            "group2_stop_mA": 0,
            "group2_step_mA": 5.0,
            # 新增：时延参数
            "group1_delay_s": 5,            # 组1温度步进后的等待时间(秒)
            "group2_delay_s": 2,            # 组2电流步进后的等待时间(秒)
            # 新增：文件名参数
            "group1_summary_filename": "Test1_summary",
            "group2_summary_filename": "Test2_summary"
        }
        self.param_labels = {
            "laser_exe_path": "软件路径",
            "osa_ip": "IP地址",
            "current_mA": "电流 (mA)",
            "t_start": "初始温度 (℃)",
            "t_stop": "终止温度 (℃)",
            "t_step": "温度温度 (℃)",
            #"center_nm": "中心波长 (nm)",
            #"span_nm": "波长范围 (nm)",
            "save_path": "保存路径",
            # group2
            "group2_temp_C": "组2 固定温度 (℃)",   # ✅ 新增：GUI 显示名称
            "group2_start_mA": "组2 初始电流 (mA)",
            "group2_stop_mA": "组2 终止电流 (mA)",
            "group2_step_mA": "组2 步进电流 (mA)",
            # 新增：时延参数标签
            "group1_delay_s": "组1 温度稳定时间 (秒)",
            "group2_delay_s": "组2 电流稳定时间 (秒)",
            # 新增：文件名参数标签
            "group1_summary_filename": "组1文件名",
            "group2_summary_filename": "组2文件名"
        }

        self.create_widgets()
        self.laser: Optional[LaserController] = None
        self.osa: Optional[OSAController] = None
        self.runner: Optional[TestRunner] = None
        self.runner_thread: Optional[threading.Thread] = None
        # 添加组1和组2的运行状态标志
        self.group1_running = False
        self.group2_running = False

    def set_center(self, width: int, height: int):
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        posx = (screenwidth - width) // 2
        posy = (screenheight - height) // 2
        self.root.geometry(f'{width}x{height}+{posx}+{posy}')
    
    def create_widgets(self):
        # 创建主容器框架，用于放置参数设置和日志框
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # ================= 参数设置大框 ================= #
        param_frame = tk.LabelFrame(main_container, text="参数设置", padx=8, pady=8)
        param_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.entries: Dict[str, tk.Entry] = {}
        
        # ================= 连接 ================= #
        connect_frame = tk.LabelFrame(param_frame, text="连接与地址", padx=8, pady=8)
        connect_frame.pack(fill=tk.X, padx=6, pady=4)

        # 将VISA地址标签改为IP地址
        self._add_param_entry(
            connect_frame, "osa_ip", "IP 地址:", 
            self.params.get("osa_ip", ""), row=0
        )
        self._add_param_entry(
            connect_frame, "save_path", "保存路径:", 
            self.params.get("save_path", "./data"), row=1
        )
        self._add_param_entry(
            connect_frame, "laser_exe_path", "软件路径:", 
            self.params.get("laser_exe_path", ""), row=2
        )
        
        # 按钮
        connect_buttons = tk.Frame(connect_frame)
        connect_buttons.grid(row=6, column=0, columnspan=3, pady=4)
        self.btn_connect = tk.Button(
            connect_buttons, text="连接", command=self.diag_connect_and_query, 
            bg="#1D74C0", fg="#FFFFFF", width=12
        )
        self.btn_connect.pack(side=tk.LEFT, padx=4)

        self.btn_connect = tk.Button(
            connect_buttons, text="上位机", command=self.open_laser_software, 
            bg="#1D74C0", fg="#FFFFFF", width=12
        )
        self.btn_connect.pack(side=tk.RIGHT, padx=4)

        # -------- 第一组测试 -------- #
        group1_frame = tk.LabelFrame(param_frame, text="第一组测试", padx=6, pady=6)
        group1_frame.pack(fill="x", padx=6, pady=4)

        self._add_param_entry(group1_frame, "t_start", "初始温度:", self.params.get("t_start", 20.0), row=0)
        self._add_param_entry(group1_frame, "t_stop", "终止温度:", self.params.get("t_stop", 40.0), row=1)
        self._add_param_entry(group1_frame, "t_step", "步进温度:", self.params.get("t_step", 0.5), row=2)
        self._add_param_entry(group1_frame, "current_mA", "固定电流:", self.params.get("current_mA", 360.0), row=5)
        # 新增：组1时延参数输入框
        self._add_param_entry(group1_frame, "group1_delay_s", "稳定时间:", self.params.get("group1_delay_s", 5), row=6)
        # 新增：组1文件名输入框
        self._add_param_entry(group1_frame, "group1_summary_filename", "保存文件名", self.params.get("group1_summary_filename", "Test1_summary.csv"), row=7)
        # 为第一组添加开始和停止按钮
        group1_buttons = tk.Frame(group1_frame)
        group1_buttons.grid(row=8, column=0, columnspan=3, pady=4)
        self.btn_group1_start = tk.Button(
            group1_buttons, text="开始测试", command=self.start_group1, 
            bg="#4CAF50", fg="#FFFFFF", width=12
        )
        self.btn_group1_start.pack(side=tk.LEFT, padx=4)
        self.btn_group1_stop = tk.Button(
            group1_buttons, text="停止测试", command=self.stop_group1, 
            bg="#f44336", fg="#FFFFFF", width=12,
        )
        self.btn_group1_stop.pack(side=tk.LEFT, padx=4)
    
        # -------- 第二组测试 -------- #
        group2_frame = tk.LabelFrame(param_frame, text="第二组测试", padx=6, pady=6)
        group2_frame.pack(fill="x", padx=6, pady=4)
    
        self._add_param_entry(group2_frame, "group2_start_mA", "初始电流:", self.params.get("group2_start_mA", 400.0), row=0)
        self._add_param_entry(group2_frame, "group2_stop_mA", "终止电流:", self.params.get("group2_stop_mA", 0.5), row=1)
        self._add_param_entry(group2_frame, "group2_step_mA", "步进电流:", self.params.get("group2_step_mA", 5.0), row=2)
        self._add_param_entry(group2_frame, "group2_temp_C", "测试温度:", self.params.get("group2_temp_C", 25.0), row=3)
        # 新增：组2时延参数输入框
        self._add_param_entry(group2_frame, "group2_delay_s", "稳定时间:", self.params.get("group2_delay_s", 2), row=4)
        # 新增：组2文件名输入框
        self._add_param_entry(group2_frame, "group2_summary_filename", "保存文件名:", self.params.get("group2_summary_filename", "Test2_summary.csv"), row=5)
        # 为第二组添加开始和停止按钮
        group2_buttons = tk.Frame(group2_frame)
        group2_buttons.grid(row=6, column=0, columnspan=3, pady=4)
        self.btn_group2_start = tk.Button(
            group2_buttons, text="开始测试", command=self.start_group2, 
            bg="#4CAF50", fg="#FFFFFF", width=12
        )
        self.btn_group2_start.pack(side=tk.LEFT, padx=4)
        self.btn_group2_stop = tk.Button(
            group2_buttons, text="停止测试", command=self.stop_group2, 
            bg="#f44336", fg="#FFFFFF", width=12
        )
        self.btn_group2_stop.pack(side=tk.LEFT, padx=4)

        # ================= 日志 ================= #
        log_frame = tk.LabelFrame(main_container, text="运行日志", padx=6, pady=6)
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _add_param_entry(self, parent, key, label, default="", row=0, browse=None):
        tk.Label(parent, text=label, anchor="e", width=14).grid(row=row, column=0, sticky="e", padx=4, pady=4)
        ent = tk.Entry(parent, width=30)
        ent.insert(0, str(self.params.get(key, default)))
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

    def open_laser_software(self):
        p = self.get_params()
        try:
            # 获取软件路径
            exe_path = p["laser_exe_path"]
            if not exe_path:
                messagebox.showerror("错误", "请先设置软件路径")
                return
            
            # 在单独的线程中执行打开上位机的操作
            def _open_laser_thread():
                try:
                    # 创建LaserController实例并连接
                    self.laser = LaserController(exe_path=exe_path, window_title=r"Preci-Semi-Seed", log_func=self.log)
                    self.laser.connect()
                    self.log("[上位机] 已成功打开或连接到上位机软件")
                    # 确保UI更新在主线程中进行
                    
                except Exception as e:
                    error_msg = f"[错误] 打开上位机软件失败: {e}"
                    self.log(error_msg)
                    # 确保UI更新在主线程中进行
                    self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                    self.laser = None
            
            # 启动线程
            thread = threading.Thread(target=_open_laser_thread, daemon=True)
            thread.start()
            
        except Exception as e:
            self.log(f"[错误] 准备打开上位机软件失败: {e}")
            messagebox.showerror("错误", f"准备打开上位机软件失败: {e}")

    def browse_file(self, param_key: str):
        filename = filedialog.askopenfilename(title="选择激光控制软件 (exe)", filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")])
        if filename:
            self.entries[param_key].delete(0, tk.END)
            self.entries[param_key].insert(0, filename)

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

    def get_params(self) -> Dict[str, Any]:
        p = {}  
        for k in self.params.keys():
            try:
                if k in self.entries:
                    val = self.entries[k].get()
                    if k in ("laser_exe_path", "osa_ip", "save_path", "group1_summary_filename", "group2_summary_filename"):
                        p[k] = val
                    else:
                        p[k] = float(val)
                else:
                    # 如果条目不存在，使用默认值
                    p[k] = self.params[k]
            except Exception:
                p[k] = float(self.params[k]) if k not in ("laser_exe_path", "osa_ip", "save_path", "group1_summary_filename", "group2_summary_filename") else self.params[k]
        return p

    def show_image_popup(self, img_path, title="测试完成 - 截图预览"):
        win = tk.Toplevel(self.root)
        win.title(title)

        # 读取原始图片
        try:
            img = Image.open(img_path)
        except Exception as e:
            self.log(f"[错误] 无法打开图片: {e}")
            messagebox.showerror("错误", f"无法打开图片: {e}")
            return

        # 获取屏幕尺寸
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        max_w, max_h = int(sw * 0.8), int(sh * 0.8)

        disp_img = img
        if img.width > max_w or img.height > max_h:
            scale = min(max_w / img.width, max_h / img.height)
            new_size = (int(img.width * scale), int(img.height * scale))
            disp_img = img.resize(new_size, Image.LANCZOS)

        img_tk = ImageTk.PhotoImage(disp_img)

        # 挂载引用，避免被回收
        win.img = img
        win.disp_img = disp_img
        win.img_tk = img_tk

        # 顶部按钮区
        btn_frame = tk.Frame(win)
        btn_frame.pack(side=tk.TOP, pady=8)

        def save_img():
            save_path = filedialog.asksaveasfilename(
                defaultextension=".bmp",
                filetypes=[("BMP 文件", "*.bmp"), ("PNG 文件", "*.png"), ("所有文件", "*.*")],
                title="保存图片"
            )
            if save_path:
                # 保存的就是原始图
                win.img.save(save_path)
                messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")

        save_btn = tk.Button(btn_frame, text="保存图片", command=save_img)
        save_btn.pack()

        # 显示图片
        lbl = tk.Label(win, image=win.img_tk)
        lbl.pack(padx=6, pady=6)

    # Diagnostics
    # 修改诊断连接方法，内部构建VISA地址格式
    def diag_connect_and_query(self):
        # 获取IP地址输入
        ip_addr = self.entries["osa_ip"].get().strip()
        if not ip_addr:
            messagebox.showerror("错误", "请在诊断面板填写 IP 地址")
            return
        # 在内部构建完整的VISA地址格式
        visa_addr = f"TCPIP0::{ip_addr}::INSTR"
        try:
            osa = OSAController(resource=visa_addr, log_func=self.log)
            osa.connect()
            idn = osa.query_idn()

            # 自动设置为ASCII格式
            osa.inst.write(":FORMat:DATA ASCII")
            time.sleep(0.2)  # 等待设置生效
            fmt = osa.query_format()

            self.log(f"[Diag] 连接成功, 已自动设置FORMAT={fmt}")
            self.osa = osa
        except Exception as e:
            self.log(f"[Diag] 连接/查询失败: {e}")
            messagebox.showerror("错误", f"诊断失败: {e}")

    # 修改CT_W_GUI类的start_group1方法，在测试完成后调用绘图函数
    def start_group1(self):
        p = self.get_params()
        self.btn_group1_start.config(state=tk.DISABLED)
        self.btn_group1_stop.config(state=tk.NORMAL)
        self.group1_running = True
        
        try:
            # 初始化激光器和OSA控制器
            if not self.laser:
                self.laser = LaserController(exe_path=p["laser_exe_path"], window_title=r"Preci-Semi-Seed", log_func=self.log)
                try:
                    self.laser.connect()
                except Exception as e:
                    self.log(f"[错误] 激光控制软件连接失败: {e}")
                    if not messagebox.askyesno("警告", "激光控制软件连接失败，是否继续仅使用 OSA?"):
                        self.btn_group1_start.config(state=tk.NORMAL)
                        self.btn_group1_stop.config(state=tk.DISABLED)
                        self.group1_running = False
                        return
                    else:
                        self.laser = None

            if not self.osa:
                visa_address = f"TCPIP0::{p['osa_ip']}::INSTR"
                self.osa = OSAController(resource=visa_address, log_func=self.log)
                self.osa.connect()

            if not self.runner:
                self.runner = TestRunner(self.laser, self.osa, log_func=self.log)
            else:
                # 重置停止标志
                self.runner._stop = False

            def target():
                try:
                    self.runner.run_group1(
                        start_temp=p["t_start"],
                        end_temp=p["t_stop"],
                        step=p["t_step"],
                        save_path=p["save_path"],
                        # 新增：传递组1时延参数
                        delay_s=p["group1_delay_s"],
                        # 新增：传递文件名参数
                        summary_filename=p["group1_summary_filename"],
                        # 新增：传递电流参数
                        current_mA=p["current_mA"]
                    )
                    # 在测试完成后调用绘图函数，并传递文件名参数
                    img_path = self.runner.plot_group1_wavelength_vs_temperature(
                        p["save_path"], 
                        summary_filename=p["group1_summary_filename"]
                    )
                    # 如果成功保存了图像，显示弹窗
                    if img_path and os.path.exists(img_path):
                        self.root.after(0, lambda: self.show_image_popup(img_path, "第一组测试完成 - 截图预览"))
                except Exception as e:
                    self.log(f"[线程异常] {e}\n{traceback.format_exc()}")
                finally:
                    try:
                        self.btn_group1_start.config(state=tk.NORMAL)
                        self.btn_group1_stop.config(state=tk.DISABLED)
                        self.group1_running = False
                    except Exception:
                        pass

            self.runner_thread = threading.Thread(target=target, daemon=True)
            self.runner_thread.start()
            self.log("[主] 第一组测试线程已启动")
        except Exception as e:
            self.log(f"[错误] 启动第一组测试失败: {e}")
            messagebox.showerror("错误", f"启动第一组测试失败: {e}")
            self.btn_group1_start.config(state=tk.NORMAL)
            self.btn_group1_stop.config(state=tk.DISABLED)
            self.group1_running = False

    # 新增：停止第一组测试
    def stop_group1(self):
        if self.runner and self.group1_running:
            try:
                self.runner.stop()
                self.log("[主] 第一组测试停止信号已发送")
            except Exception as e:
                self.log(f"[错误] 停止第一组测试失败: {e}")
        else:
            self.log("[主] 没有正在运行的第一组测试")

    # 新增：启动第二组测试
    def start_group2(self):
        p = self.get_params()
        self.btn_group2_start.config(state=tk.DISABLED)
        self.btn_group2_stop.config(state=tk.NORMAL)
        self.group2_running = True
        
        try:
            # 初始化激光器和OSA控制器
            if not self.laser:
                self.laser = LaserController(exe_path=p["laser_exe_path"], window_title=r"Preci-Semi-Seed", log_func=self.log)
                try:
                    self.laser.connect()
                except Exception as e:
                    self.log(f"[错误] 激光控制软件连接失败: {e}")
                    if not messagebox.askyesno("警告", "激光控制软件连接失败，是否继续仅使用 OSA?"):
                        self.btn_group2_start.config(state=tk.NORMAL)
                        self.btn_group2_stop.config(state=tk.DISABLED)
                        self.group2_running = False
                        return
                    else:
                        self.laser = None

            if not self.osa:
                visa_address = f"TCPIP0::{p['osa_ip']}::INSTR"
                self.osa = OSAController(resource=visa_address, log_func=self.log)
                self.osa.connect()

            if not self.runner:
                self.runner = TestRunner(self.laser, self.osa, log_func=self.log)
            else:
                # 重置停止标志
                self.runner._stop = False

            def target():
                try:
                    # 先创建一个保存图像路径的变量
                    img_path = None
                    self.runner.run_group2(
                        start_mA=p["group2_start_mA"],
                        step_mA=p["group2_step_mA"],
                        stop_mA=p["group2_stop_mA"],
                        temp_C=p["group2_temp_C"],
                        save_path=p["save_path"],
                        # 新增：传递组2时延参数
                        delay_s=p["group2_delay_s"],
                        # 新增：传递文件名参数
                        summary_filename=p["group2_summary_filename"]
                    )
                    import glob
                    
                    # 匹配由 _plot_xy_curve 保存的第二组图片（前缀是“电流波长关系图”）
                    pattern = os.path.join(p["save_path"], "电流波长关系图_*.png")
                    group2_files = glob.glob(pattern)

                    # 按修改时间排序，获取最新的文件
                    if group2_files:
                        group2_files.sort(key=os.path.getmtime, reverse=True)
                        img_path = group2_files[0]
                        self.log(f"[Runner] 找到最新的第二组测试图像: {img_path}")

                        # 显示自动弹窗
                        if img_path and os.path.exists(img_path):
                            self.root.after(0, lambda: self.show_image_popup(img_path, "第二组测试完成 - 截图预览"))
                    else:
                        self.log("[Runner] 未找到第二组测试图像，请检查保存路径或命名。")
                except Exception as e:
                    self.log(f"[线程异常] {e}\n{traceback.format_exc()}")
                finally:
                    try:
                        self.btn_group2_start.config(state=tk.NORMAL)
                        self.btn_group2_stop.config(state=tk.DISABLED)
                        self.group2_running = False
                    except Exception:
                        pass

            self.runner_thread = threading.Thread(target=target, daemon=True)
            self.runner_thread.start()
            self.log("[主] 第二组测试线程已启动")
        except Exception as e:
            self.log(f"[错误] 启动第二组测试失败: {e}")
            messagebox.showerror("错误", f"启动第二组测试失败: {e}")
            self.btn_group2_start.config(state=tk.NORMAL)
            self.btn_group2_stop.config(state=tk.DISABLED)
            self.group2_running = False

    # 新增：停止第二组测试
    def stop_group2(self):
        if self.runner and self.group2_running:
            try:
                self.runner.stop()
                self.log("[主] 第二组测试停止信号已发送")
            except Exception as e:
                self.log(f"[错误] 停止第二组测试失败: {e}")
        else:
            self.log("[主] 没有正在运行的第二组测试")

    def single_scan(self):
        p = self.get_params()
        try:
            if not self.osa:
                visa_address = f"TCPIP0::{p['osa_ip']}::INSTR"
                self.osa = OSAController(resource=visa_address, log_func=self.log)
                self.osa.connect()

            try:
                self.osa.sweep_and_fetch()
            except Exception:
                pass

            wavelengths, powers = self.osa.fetch_trace()
            npoints = len(powers)
            self.log(f"[单次] 读取到 {npoints} 点")

            save_base = p["save_path"]
            if os.path.isdir(save_base) or save_base.endswith(os.sep):
                fig_dir = save_base
            else:
                fig_dir = os.path.dirname(save_base) or "."
            ensure_dir(fig_dir)
            fig_path = os.path.join(fig_dir, f"single_scan_{time.strftime('%Y%m%d_%H%M%S')}.png")

            plt.figure(figsize=(8, 4))
            if wavelengths is not None and len(wavelengths) == npoints:
                plt.plot(wavelengths, powers)
                plt.xlabel("Wavelength (nm)")
            else:
                plt.plot(np.arange(npoints), powers)
                plt.xlabel("Point")
            plt.title("Single Scan")
            plt.ylabel("Power")
            plt.tight_layout()
            plt.savefig(fig_path)
            plt.close()
            self.log(f"[单次] 图像保存到 {fig_path}")

            csv_fn = os.path.join(fig_dir, f"single_scan_{time.strftime('%Y%m%d_%H%M%S')}.csv")
            with open(csv_fn, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Wavelength_nm", "Power"])
                for x, y in zip(wavelengths, powers):
                    # 单次扫描：波长保留 4 位小数，功率保持原样
                    w.writerow([f"{float(x):.4f}", f"{float(y):.6f}"])
            self.log(f"[单次] 光谱 CSV 保存到 {csv_fn}")

        except Exception as e:
            self.log(f"[错误] 单次扫描失败: {e}\n{traceback.format_exc()}")
            messagebox.showerror("错误", f"单次扫描失败: {e}")

    def run(self):
        # 保持原有的run方法
        if self.root.winfo_exists():
            self.root.mainloop()

if __name__ == "__main__":
    gui = CT_W_GUI()
    gui.run()

# pyinstaller --onefile --noconsole --icon="D:\pack\PreciLasers.ico" --hidden-import=pyvisa --clean "D:\Coding\Project\DataAutomation\InstrumentControlSystem\测试系统\qijian\01-电流温度调谐\CT_Wv12.py"
