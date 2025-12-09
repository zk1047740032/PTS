# 精测中心截图
from __future__ import annotations

import os
import time
import threading
import csv
import struct
import traceback
import math
from typing import Tuple, Optional, Any, Dict, List

import pyvisa
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog
import matplotlib
import matplotlib.ticker as mticker
from pyvisa.errors import VisaIOError
# PIL
try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except ImportError:
    try:
        import Image, ImageDraw, ImageFont, ImageTk
    except Exception:
        raise ImportError("请安装Pillow库: pip install pillow")

matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False

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
# LaserController (unchanged)
# -------------------------
class LaserController:
    def __init__(self, exe_path: str = r"C:\PTS\CT_L\Preci_Semi\Preci-Seed.exe",
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
# SpectrumAnalyzerController (for FSV3004 or compatible)
# -------------------------
class SpectrumAnalyzerController:
    """兼容 Rohde & Schwarz FSV3004 全固件版本的线宽控制类
    —— 按你提供的可运行脚本逻辑重写。
    """

    def __init__(self, resource: str, log_func=print):
        self.rm = pyvisa.ResourceManager()
        self.inst = None
        self.resource = resource
        self.log = log_func

    def connect(self):
        try:
            self.inst = self.rm.open_resource(self.resource)
            self.inst.timeout = 10000
            self.inst.write_termination = "\n"
            self.inst.read_termination = "\n"
            idn = self.inst.query("*IDN?").strip()
            self.log(f"[FSV] 已连接: {idn}")
            return idn
        except Exception as e:
            self.log(f"[FSV] 连接失败: {e}")
            raise

    def query_idn(self):
        try:
            return self.inst.query("*IDN?").strip()
        except Exception:
            return ""

    def query_format(self):
        return "ASCII"

    # --------------------- #
    # 线宽测量逻辑（简化版）
    # --------------------- #
    def measure_linewidth_kHz(self):
        """
        使用 Rohde & Schwarz FSV3004 测量线宽 (20 dB 带宽)
        - 优先使用硬件的 n dB down marker 功能 (CALC:MARK1:FUNC:NDBDown)
        - 若仪器不支持或超时，则自动回退到软件计算 (基于 TRACE 数据)
        返回值: 线宽 (kHz)
        """
        if self.inst is None:
            raise RuntimeError("频谱仪未连接。")

        try:
            self.log("[FSV] 开始测量线宽: 80MHz, span=1MHz, RBW=100Hz")
            #self.inst.clear()
            self.inst.timeout = 20000
            self.inst.write("*CLS")
            self.inst.write("INIT:CONT OFF")

            # 基本扫描设置
            self.inst.write("FREQ:CENT 80MHz")
            self.inst.write("FREQ:SPAN 1MHz")
            self.inst.write("BAND 100Hz")
            self.inst.write("SWE:POIN 2001")

            # 执行扫描并等待完成
            self.inst.write("INIT; *WAI")
            opc = self.inst.query("*OPC?")
            self.log(f"[FSV] 扫描完成确认: {opc.strip()}")

            # 开启 Marker 并执行 20 dB 带宽测量
            self.inst.write("CALC:MARK1 ON")
            self.inst.write("CALC:MARK1:MAX")
            self.inst.write("CALC:MARK1:FUNC:NDBDown 20")
            self.inst.write("CALC:MARK1:FUNC:NDBDown:STAT ON")

            # 等待计算完成
            time.sleep(1.0)
            self.inst.query("*OPC?")

            # 查询 3 dB 带宽结果
            try:
                bw_hz_str = self.inst.query("CALC:MARK1:FUNC:NDBDown:RES?").strip()
                bw_hz = float(bw_hz_str)
                self.log(f"[FSV] 成功读取 20 dB 带宽: {bw_hz:.3f} Hz")
            except Exception as e:
                raise RuntimeError(f"仪器返回无效带宽结果: {e}")

            # 可选修正（如需和旧版保持一致，可保留此项）
            corrected = bw_hz / (2 * np.sqrt(99))
            self.log(f"[FSV] 修正后线宽: {corrected / 1e3:.3f} kHz")

            return corrected / 1e3

        except (pyvisa.errors.VisaIOError, RuntimeError) as e:
            self.log(f"[FSV][警告] 硬件线宽读取失败: {e}")
            self.log("[FSV] 自动切换到软件线宽测量模式...")
            return self.measure_linewidth_from_trace()

        except Exception as e:
            self.log(f"[FSV][错误] 线宽测量失败: {e}")
            import traceback
            self.log(f"[FSV][错误] 详细错误信息: {traceback.format_exc()}")
            return float("nan")


    def measure_linewidth_from_trace(self):
        """软件计算线宽 (3 dB 带宽)，当仪器不支持 FUNC:RES? 时使用"""
        if self.inst is None:
            raise RuntimeError("频谱仪未连接。")
        try:
            self.log("[FSV] 启动软件线宽测量 (基于 Trace 数据)")

            # 清空命令缓冲区（部分LAN接口不支持clear）
            try:
                self.inst.clear()
            except pyvisa.errors.VisaIOError:
                self.log("[FSV] clear() 不支持，已自动跳过。")

            # 基本配置
            self.inst.write("*CLS")
            self.inst.write("INIT:CONT OFF")
            self.inst.write("SWE:POIN 2001")
            self.inst.write("DISP:WIND:TRAC:Y:RLEV 0dBm")
            self.inst.write("FREQ:CENT 80MHz")
            self.inst.write("FREQ:SPAN 1MHz")
            self.inst.write("BAND 100Hz")

            # ⭐立即触发单次扫描
            self.inst.write("INIT:IMM; *WAI")
            opc = self.inst.query("*OPC?")
            self.log(f"[FSV] 扫描完成确认: {opc.strip()}")

            # 读取 Trace 数据
            ydata = np.array(self.inst.query_ascii_values("TRAC:DATA? TRACE1"))
            start = float(self.inst.query("FREQ:STAR?"))
            stop = float(self.inst.query("FREQ:STOP?"))
            xdata = np.linspace(start, stop, len(ydata))

            # 寻找峰值与3dB宽度
            peak_idx = np.argmax(ydata)
            peak_power = ydata[peak_idx]
            half_power = peak_power - 3.0

            left_idxs = np.where(ydata[:peak_idx] <= half_power)[0]
            right_idxs = np.where(ydata[peak_idx:] <= half_power)[0]
            if len(left_idxs) == 0 or len(right_idxs) == 0:
                raise RuntimeError("未检测到有效的 20dB 交点，请检查信号曲线。")

            f_left = xdata[left_idxs[-1]]
            f_right = xdata[peak_idx + right_idxs[0]]
            bw_hz = abs(f_right - f_left)
            self.log(f"[FSV] 软件计算带宽: {bw_hz:.3f} Hz")

            return bw_hz / 1e3  # 转 kHz

        except Exception as e:
            self.log(f"[FSV][错误] 软件线宽计算失败: {e}")
            import traceback
            self.log(traceback.format_exc())
            return float("nan")
            
    # --------------------- #
    # Trace 读取（可选）
    # --------------------- #
    def fetch_trace(self):
        try:
            data = self.inst.query_ascii_values(":TRAC:DATA? TRACE1")
            n = len(data)
            start = float(self.inst.query("FREQ:STAR?"))
            stop = float(self.inst.query("FREQ:STOP?"))
            freqs = np.linspace(start, stop, n)
            return freqs, np.array(data)
        except Exception as e:
            raise RuntimeError(f"读取Trace失败: {e}")
        
    def save_last_trace_to_csv(self, local_path: str) -> str:
        """
        读取当前 Trace (TRACE1) 并保存为本地 CSV。
        返回本地文件路径 (str)。
        """
        if self.inst is None:
            raise RuntimeError("频谱仪未连接。")

        try:
            # 触发一次采样并读取 trace 点
            try:
                self.inst.write("INIT:IMM; *WAI")
                self.inst.query("*OPC?")
            except Exception:
                # 部分设备不支持 INIT:IMM，忽略
                pass

            # 读取 trace 数据（ASCII）
            ydata = self.inst.query_ascii_values("TRAC:DATA? TRACE1")
            if ydata is None or len(ydata) == 0:
                raise RuntimeError("读取TRAC:DATA? 返回空数据")

            # 读取频率起止，生成 x 轴
            start_hz = float(self.inst.query("FREQ:STAR?"))
            stop_hz = float(self.inst.query("FREQ:STOP?"))
            n = len(ydata)
            freqs = np.linspace(start_hz, stop_hz, n)

            # 写入 CSV
            ensure_dir(os.path.dirname(local_path) or ".")
            with open(local_path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Frequency_Hz", "Power_dBm"])
                for x, y in zip(freqs, ydata):
                    w.writerow([f"{float(x):.6f}", f"{float(y):.6f}"])
            self.log(f"[FSV] Trace 数据已保存到 {local_path}")
            return local_path

        except Exception as e:
            self.log(f"[FSV][警告] 保存 Trace 到 CSV 失败: {e}")
            raise

    def capture_screenshot_to_local(self, local_path: str, inst_file_name: Optional[str] = None, timeout: float = 5.0) -> str:
        """
        尝试通过仪器的 MMEM 命令生成屏幕截图并把它传回到本地。
        流程：
          1) 在仪器上执行: MMEM:STOR:IMAG '<inst_file_name>'
          2) 从仪器读取: MMEM:DATA? '<inst_file_name>' （二进制块）
          3) 解析二进制块并写入 local_path

        如果上述任一步失败，则抛出异常。调用方应当捕获异常并（可选）退回用本地绘图保存 trace 作为后备。
        """
        if self.inst is None:
            raise RuntimeError("频谱仪未连接。")

        try:
            # 确定在仪器上的文件名（如果未提供，用时间戳）
            if inst_file_name is None:
                inst_file_name = f"fine_capture_{int(time.time())}.png"

            # 尝试在仪器上生成截图文件
            try:
                # 该命令在不同固件间语法可能有差异；这儿使用常见格式 MMEM:STOR:IMAG 'name'
                cmd = f"MMEM:STOR:IMAG '{inst_file_name}'"
                self.inst.write(cmd)
                # 等待内部生成
                time.sleep(0.5)
                # query OPC to ensure completion if supported
                try:
                    self.inst.query("*OPC?")
                except Exception:
                    pass
            except Exception as e:
                raise RuntimeError(f"在仪器上生成截图失败: {e}")

            # 读取二进制文件内容到本地
            try:
                # 请求文件二进制数据
                self.inst.write(f"MMEM:DATA? '{inst_file_name}'")
                raw = self.inst.read_raw()  # returns bytes
            except Exception as e:
                raise RuntimeError(f"从仪器读取截图二进制数据失败: {e}")

            # 解析 IEEE block header (#<n><len><data>) 或直接原始字节
            data = raw
            try:
                if len(raw) >= 2 and raw[0:1] == b'#':
                    # 第一字节 '#' ; 第二字节为 header length digits count
                    header_len_digit = int(raw[1:2])
                    header_len = 2 + header_len_digit
                    total_len_digits = int(raw[2:2 + header_len_digit])
                    start_idx = 2 + header_len_digit
                    data = raw[start_idx:start_idx + total_len_digits]
                # else: data remains raw
            except Exception:
                # 如果解析出错，仍尝试写入原始 raw（某些设备返回裸图片）
                data = raw

            # 写文件
            ensure_dir(os.path.dirname(local_path) or ".")
            with open(local_path, "wb") as f:
                f.write(data)
            self.log(f"[FSV] 仪器截图已保存到 {local_path} (inst:{inst_file_name})")
            return local_path

        except Exception as e:
            self.log(f"[FSV][警告] capture_screenshot_to_local 失败: {e}")
            raise


# -------------------------
# TestRunner (modified to use Spectrum Analyzer)
# -------------------------
class TestRunner:
    def __init__(self, laser: Optional[LaserController], sa: SpectrumAnalyzerController, log_func=print):
        self.laser = laser
        self.sa = sa
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
        if start < stop:
            while t <= stop + 1e-9:
                out.append(round(t, 6))
                t += step_magnitude
        else:
            while t >= stop - 1e-9:
                out.append(round(t, 6))
                t -= step_magnitude
        return out

    def _build_temps_with_fine(self, start_temp: float, end_temp: float, step: float,
                            fine_center: Optional[float], fine_range: Optional[float]) -> List[float]:
        """
        生成包含精测区间（0.1步长）和粗测区间（step步长）的温度序列。
        自动处理：
        - 精测区间部分或全部越界；
        - 起止顺序升/降；
        - 边界去重。
        """
        from collections import OrderedDict

        def frange(a, b, s):
            out = []
            if s == 0:
                raise ValueError("step cannot be 0")
            if a > b:
                while a >= b - 1e-9:
                    out.append(round(a, 6))
                    a -= abs(s)
            else:
                while a <= b + 1e-9:
                    out.append(round(a, 6))
                    a += abs(s)
            return out

        start_temp = float(start_temp)
        end_temp = float(end_temp)
        step = float(step)

        # 若未定义精测参数，直接返回
        if fine_center is None or fine_range is None:
            return frange(start_temp, end_temp, step)

        fine_center = float(fine_center)
        fine_range = float(fine_range)

        # 计算理想精测区间
        fine_upper = round(fine_center + fine_range, 6)
        fine_lower = round(fine_center - fine_range, 6)

        # 计算整体扫描上下界
        t_min, t_max = (end_temp, start_temp) if start_temp > end_temp else (start_temp, end_temp)

        # 越界修正
        orig_upper, orig_lower = fine_upper, fine_lower
        if fine_upper > t_max:
            fine_upper = t_max
        if fine_lower < t_min:
            fine_lower = t_min
        if (fine_upper != orig_upper) or (fine_lower != orig_lower):
            self.log(f"[警告] 精测区间超出主扫描范围，已自动修正为 [{fine_lower}, {fine_upper}]")

        temps = []

        # 判断扫描方向
        descending = start_temp > end_temp

        # ---------- 生成温度序列 ----------
        if descending:
            # 上半段（高温到精测上界）
            if start_temp > fine_upper + 1e-6:
                coarse_high = frange(start_temp, fine_upper + step, step)
                if coarse_high:
                    if abs(coarse_high[-1] - fine_upper) < 1e-6:
                        coarse_high = coarse_high[:-1]
                    temps += coarse_high
            # 精测段（fine_upper → fine_lower, 步长0.1）
            temps += frange(fine_upper, fine_lower, 0.1)
            # 下半段（精测下界到终止）
            if fine_lower - step > end_temp + 1e-6:
                coarse_low = frange(fine_lower - step, end_temp, step)
                if coarse_low and abs(coarse_low[0] - fine_lower) < 1e-6:
                    coarse_low = coarse_low[1:]
                temps += coarse_low
        else:
            # 上升扫描
            if start_temp < fine_lower - 1e-6:
                coarse_low = frange(start_temp, fine_lower - step, step)
                if coarse_low and abs(coarse_low[-1] - fine_lower) < 1e-6:
                    coarse_low = coarse_low[:-1]
                temps += coarse_low
            temps += frange(fine_lower, fine_upper, 0.1)
            if fine_upper + step < end_temp - 1e-6:
                coarse_high = frange(fine_upper + step, end_temp, step)
                if coarse_high and abs(coarse_high[0] - fine_upper) < 1e-6:
                    coarse_high = coarse_high[1:]
                temps += coarse_high

        # 去重（防浮点误差）
        temps = list(OrderedDict.fromkeys([round(t, 6) for t in temps]))
        return temps




    def _append_summary(self, save_path: str, current_mA: float, temperature: Optional[float], linewidth_khz: float, test_group: int = 0, summary_filename: str = None):
        if os.path.isdir(save_path) or save_path.endswith(os.sep):
            out_dir = save_path
        else:
            out_dir = os.path.dirname(save_path) or "."
        ensure_dir(out_dir)
        if summary_filename:
            if not summary_filename.lower().endswith('.csv'):
                summary_filename += '.csv'
            summary_fn = os.path.join(out_dir, summary_filename)
        elif test_group == 1:
            summary_fn = os.path.join(out_dir, "Test1_summary.csv")
        elif test_group == 2:
            summary_fn = os.path.join(out_dir, "Test2_summary.csv")
        else:
            summary_fn = os.path.join(out_dir, "ct_tuning_summary.csv")

        header_needed = not os.path.exists(summary_fn)
        with open(summary_fn, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if header_needed:
                w.writerow(["Current_mA", "Temperature_C", "Linewidth_kHz"])
            temp_str = f"{temperature:.2f}" if temperature is not None else "N/A"
            w.writerow([f"{current_mA:.2f}", temp_str, f"{linewidth_khz:.6f}"])

    def _plot_xy_curve(self, x, y, xlabel, ylabel, title, out_dir, prefix, invert_x=False, save_csv=False, extra_cols=None):
        ensure_dir(out_dir)
        fig_path = os.path.join(out_dir, f"{prefix}.png")

        plt.figure(figsize=(20, 10))
        plt.plot(x, y, marker='o', linestyle='-', linewidth=2)
        if invert_x:
            plt.gca().invert_xaxis()
        plt.xlabel(xlabel, fontsize=20)
        plt.ylabel(ylabel, fontsize=20)
        plt.title(title, fontsize=22)

        ax = plt.gca()
        ax.ticklabel_format(style='plain', axis='y')
        ax.yaxis.get_major_formatter().set_scientific(False)
        ax.yaxis.get_major_formatter().set_useOffset(False)
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.3f'))
        ax.xaxis.get_major_formatter().set_scientific(False)
        ax.xaxis.get_major_formatter().set_useOffset(False)
        plt.xticks(fontsize=16)
        plt.yticks(fontsize=16)
        plt.grid(True, linestyle='--', alpha=0.7, which='major')
        ax.minorticks_on()
        ax.grid(True, axis='x', linestyle=':', alpha=0.5, which='minor')

        if "Temperature" in xlabel or "group1" in prefix:
            x_min, x_max = min(x), max(x)
            plt.xticks(np.arange(round(x_min), round(x_max) + 1, 1))
        elif "Current" in xlabel or "group2" in prefix:
            x_min, x_max = min(x), max(x)
            plt.xticks(np.arange(round(x_min), round(x_max) + 5, 5))

        plt.tight_layout()
        plt.savefig(fig_path, dpi=300)
        plt.close()
        self.log(f"[Runner] 图像保存到 {fig_path}")
        return fig_path

    def run_group1(self, start_temp: float, end_temp: float, step: float, save_path: str = "./data",
                   delay_s: float = 0.8, summary_filename: str = None, current_mA: float = None):
        self._stop = False
        try:
            if os.path.isdir(save_path) or save_path.endswith(os.sep):
                out_dir = save_path
            else:
                out_dir = os.path.dirname(save_path) or "."
            ensure_dir(out_dir)
            if summary_filename:
                if not summary_filename.lower().endswith('.csv'):
                    summary_filename += '.csv'
                file_path = os.path.join(out_dir, summary_filename)
            else:
                file_path = os.path.join(out_dir, "Test1_summary.csv")
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    self.log(f"[Runner] 已删除同名文件: {file_path}")
                except Exception as e:
                    self.log(f"[Runner] 删除文件失败: {e}")

            current_for_temp = 360.0
            if current_mA is not None:
                current_for_temp = current_mA
                if self.laser:
                    try:
                        self.laser.set_current_mA(current_for_temp)
                        self.log(f"[Runner] 已设置电流为 {current_for_temp} mA")
                        time.sleep(1.0)
                    except Exception as e:
                        self.log(f"[Runner] 设置电流失败: {e}")
                        val = self.laser.get_current_mA()
                        if val is not None:
                            current_for_temp = val
            elif self.laser:
                val = self.laser.get_current_mA()
                if val is not None:
                    current_for_temp = val

            # ---------- 新增精测逻辑 ----------
            fine_center = getattr(self, "fine_center_C", None)
            fine_range = getattr(self, "fine_range_C", None)
            if fine_center is not None and fine_range is not None:
                fine_upper = math.ceil(fine_center + fine_range)
                fine_lower = math.floor(fine_center - fine_range)

                # --- 检查是否超出范围 ---
                start_temp, end_temp = float(start_temp), float(end_temp)
                min_T, max_T = min(start_temp, end_temp), max(start_temp, end_temp)
                orig_upper, orig_lower = fine_upper, fine_lower

                if fine_upper > max_T:
                    fine_upper = max_T
                if fine_lower < min_T:
                    fine_lower = min_T

                # --- 打印修正日志 ---
                if (fine_upper != orig_upper) or (fine_lower != orig_lower):
                    self.log(f"[警告] 精测区间超出主扫描范围，已自动修正为 [{fine_lower}, {fine_upper}]")

                self.log(f"[Runner] 精测中心={fine_center}℃ 范围=±{fine_range}℃ → 最终精测区间 [{fine_lower}, {fine_upper}]")

                # 生成粗测与精测区段
                temps = self._build_temps_with_fine(start_temp, end_temp, step, getattr(self, "fine_center_C", None), getattr(self, "fine_range_C", None))

            self.log(f"[Runner] 实际温度点共 {len(temps)} 个: {temps}")

            self.log(f"[Runner] 组1: 电流 {current_for_temp} mA 温度扫描 {start_temp}->{end_temp} step {step} 共 {len(temps)} 步，稳定时间 {delay_s} 秒")
            stability_threshold = 0.1
            max_wait_time = delay_s * 5
            check_interval = 0.5

            # ---------- 生成 temps（含精测区间） ----------
            fine_temps = []
            fine_center = None
            fine_range = None
            if hasattr(self, "fine_center_C"):
                try:
                    fine_center = float(self.fine_center_C)
                except Exception:
                    fine_center = None
            if hasattr(self, "fine_range_C"):
                try:
                    fine_range = float(self.fine_range_C)
                except Exception:
                    fine_range = None

            if fine_center is not None and fine_range is not None:
                # 按规则计算精测上下界（向上取整/向下取整用于界定粗测分段）
                fine_upper_bound = fine_upper
                fine_lower_bound = fine_lower
                self.log(f"[Runner] 精测中心={fine_center}℃ 精测范围=±{fine_range}℃ → 精测区间界限 [{fine_lower_bound}, {fine_upper_bound}]")

                temps = []
                # 如果是从高到低（start_temp > end_temp）
                if start_temp > end_temp:
                    # 上半段（粗）
                    if start_temp > fine_upper_bound:
                        temps += self._float_range(start_temp, fine_upper_bound, step)
                    # 精测段：从 fine_upper_bound 以 0.1 递减到 fine_lower_bound
                    fine_temps = self._float_range(fine_upper_bound, fine_lower_bound, 0.1)
                    temps += fine_temps
                    # 下半段（粗）
                    if fine_lower_bound > end_temp:
                        temps += self._float_range(fine_lower_bound, end_temp, step)
                else:
                    # 正序扫描（低到高）
                    if start_temp < fine_lower_bound:
                        temps += self._float_range(start_temp, fine_lower_bound, step)
                    fine_temps = self._float_range(fine_lower_bound, fine_upper_bound, 0.1)
                    temps += fine_temps
                    if fine_upper_bound < end_temp:
                        temps += self._float_range(fine_upper_bound, end_temp, step)
            else:
                temps = self._float_range(start_temp, end_temp, step)

            self.log(f"[Runner] 组1: 温度点总数 {len(temps)} (精测点数 {len(fine_temps)})")
            # 标记精测中心点是否已保存
            fine_center_saved = False        

            # ---------- 循环测量 ----------
            for t in temps:
                if self._stop:
                    self.log("[Runner] 收到停止信号，结束组1")
                    break
                if self.laser:
                    try:
                        self.laser.set_temperature_C(t)
                        self.log(f"[Runner] 设置温度为 {t}°C，等待稳定...")
                        wait_time = 0
                        stable = False
                        time.sleep(delay_s * 0.5)
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
                                time.sleep(check_interval)
                                wait_time += check_interval
                        if not stable and not self._stop:
                            self.log(f"[Runner] 温度在 {max_wait_time}s 内未完全稳定，继续测量")
                    except Exception as e:
                        self.log(f"[Runner] 设置温度失败: {e}")
                        time.sleep(delay_s)
                else:
                    time.sleep(delay_s)

                try:
                    linewidth_khz = self.sa.measure_linewidth_kHz()
                except Exception as e:
                    self.log(f"[Runner] 组1 SA 读取失败 (temp {t}°C): {e}")
                    continue

                try:
                    self._append_summary(save_path, current_for_temp, t, linewidth_khz, test_group=1, summary_filename=summary_filename)
                    self.log(f"[Runner] 组1 {current_for_temp}mA, {t:.2f}°C -> 线宽 {linewidth_khz:.6f} kHz")
                except Exception as e:
                    self.log(f"[Runner] 组1 写入汇总失败: {e}")

                # ---- 新逻辑：在到达“精测中心温度点”时保存一次 Trace CSV + 仪器截图 ----
                try:
                    if (not fine_center_saved) and (fine_center is not None):
                        if abs(t - fine_center) <= 1e-6:
                            self.log(f"[Runner] 到达精测中心温度点 {fine_center}°C，开始保存该点 Trace 与截图...")

                            dat_filename = "fine_center.csv"
                            screenshot_name = "fine_center.png"

                            try:
                                # 保存 Trace 数据到仪器内部
                                instrument_path = f"C:\\PTS\\qijian\\CT_L\\{dat_filename}"
                                self.sa.inst.write("MMEM:MDIR 'C:\\PTS\\qijian\\CT_L'")
                                self.sa.inst.query("*OPC?")
                                self.sa.inst.write(f":MMEM:STOR:TRAC 1,'{instrument_path}'")
                                self.sa.inst.query("*OPC?")
                                self.log(f"[FSV] 精测中心数据已存储在仪器内部: {instrument_path}")

                                # 截图保存到仪器
                                self.sa.inst.write("HCOPy:DEST 'MMEM'")
                                self.sa.inst.write(f"MMEM:NAME 'C:\\PTS\\qijian\\CT_L\\{screenshot_name}'")
                                self.sa.inst.write("HCOPy:IMM")
                                self.sa.inst.query("*OPC?")
                                self.log("[FSV] 仪器已截图并保存。")

                                # 一次性复制整个目录到共享文件夹
                                instrument_ip = "192.168.29.11"
                                source_path = "C:\\PTS\\qijian\\CT_L"
                                dest_path = r"\\192.168.29.9\PTS\qijian\CT_L"
                                try:
                                    rm = pyvisa.ResourceManager()
                                    instr = rm.open_resource(f"TCPIP0::{instrument_ip}::inst0::INSTR")
                                    instr.write(f"MMEM:COPY '{source_path}\\*.*','{dest_path}'")
                                    instr.close()
                                    self.log(f"[FSV] 文件已从仪器复制到电脑共享文件夹：{dest_path}")
                                except Exception as e_copy:
                                    self.log(f"[FSV][警告] 文件复制失败: {e_copy}")

                                # # 直接尝试显示截图（无需等待同步）
                                # try:
                                #     shared_img_path = os.path.join(dest_path, screenshot_name)
                                #     if os.path.exists(shared_img_path):
                                #         self.log(f"[Runner] 从共享路径加载截图: {shared_img_path}")
                                #         self.log(f"[Runner] 精测中心截图显示完成。")
                                #     else:
                                #         self.log(f"[警告] 共享目录中未找到截图文件: {shared_img_path}")
                                # except Exception as e_show:
                                #     self.log(f"[Runner][警告] 显示截图失败: {e_show}")

                            except Exception as e:
                                self.log(f"[Runner][错误] 精测中心保存或截图失败: {e}")

                            fine_center_saved = True

                except Exception as e:
                    self.log(f"[Runner][错误] 精测中心保存/截图逻辑异常: {e}")
        except Exception as e:
            self.log(f"[Runner] 组1 出错: {e}")

        self.log("[Runner] 组1流程完成")

    def plot_group1_linewidth_vs_temperature(self, out_dir, summary_filename=None):
        try:
            filename = summary_filename if summary_filename else "Test1_summary.csv"
            if not filename.endswith('.csv'):
                filename += '.csv'
            file_path = os.path.join(out_dir, filename)
            if not os.path.exists(file_path):
                self.log(f"[Runner] {filename} 文件不存在: {file_path}")
                return

            temps, linewidths = [], []
            with open(file_path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader)
                for row in reader:
                    try:
                        temp = float(row[1])
                        lw = float(row[2])
                        temps.append(temp)
                        linewidths.append(lw)
                    except Exception as e:
                        self.log(f"[Runner] 跳过无效行 {row}: {e}")
                        continue

            if temps:
                uniq = {}
                for t, lw in zip(temps, linewidths):
                    uniq[t] = lw
                temps = sorted(uniq.keys(), reverse=True)
                linewidths = [uniq[t] for t in temps]

                return self._plot_xy_curve(
                    temps, linewidths,
                    xlabel="温度(°C)", ylabel="线宽(kHz)",
                    title=f"{self.laser.get_current_mA() if self.laser else 360:.2f} mA下温度-线宽关系",
                    out_dir=out_dir, prefix="温度线宽关系图",
                    invert_x=True, save_csv=False
                )
            else:
                self.log("[Runner] 组1 没有采集到有效线宽数据，请检查 CSV 内容")
                return None
        except Exception as e:
            self.log(f"[Runner] 组1 绘制失败: {e}")
            return None

    def run_group2(self, start_mA: float, step_mA: float, stop_mA: float, temp_C: float,
                   save_path: str = "./data", delay_s: float = 0.6, summary_filename: str = None):
        self._stop = False
        try:
            if os.path.isdir(save_path) or save_path.endswith(os.sep):
                out_dir = save_path
            else:
                out_dir = os.path.dirname(save_path) or "."
            ensure_dir(out_dir)
            if summary_filename:
                if not summary_filename.lower().endswith('.csv'):
                    summary_filename += '.csv'
                file_path = os.path.join(out_dir, summary_filename)
            else:
                file_path = os.path.join(out_dir, "Test2_summary.csv")
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    self.log(f"[Runner] 已删除同名文件: {file_path}")
                except Exception as e:
                    self.log(f"[Runner] 删除文件失败: {e}")

            if self.laser:
                self.laser.set_temperature_C(temp_C)
                self.log(f"[Runner] 组2: 设置温度为 {temp_C:.2f} °C")
                temp_stability_threshold = 0.1
                temp_max_wait_time = delay_s * 5
                temp_check_interval = 0.5
                self.log(f"[Runner] 等待温度稳定在 {temp_C:.2f}°C...")
                temp_wait_time = 0
                temp_stable = False
                time.sleep(delay_s * 0.5)
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
                        time.sleep(temp_check_interval)
                        temp_wait_time += temp_check_interval
                if not temp_stable and not self._stop:
                    self.log(f"[Runner] 温度在 {temp_max_wait_time}s 内未完全稳定，继续测量")
        except Exception as e:
            self.log(f"[Runner] 组2: 设置温度失败 {e}")

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
        peaks_lw = []

        stability_threshold = 1.0
        max_wait_time = delay_s * 3
        check_interval = 0.3

        for cur in currents:
            if self._stop:
                self.log("[Runner] 收到停止信号，提前结束组2")
                break
            try:
                if self.laser:
                    try:
                        self.laser.set_current_mA(cur)
                        self.log(f"[Runner] 设置电流为 {cur}mA，等待稳定...")
                        wait_time = 0
                        stable = False
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
                                time.sleep(check_interval)
                                wait_time += check_interval
                        if not stable and not self._stop:
                            self.log(f"[Runner] 电流在 {max_wait_time}s 内未完全稳定，继续测量")
                    except Exception as e:
                        self.log(f"[Runner] 设置电流 {cur} mA 失败: {e}")
                        time.sleep(delay_s)
                else:
                    self.log(f"[Runner] 未配置 LaserController，跳过设置电流 {cur} mA (仍会采集 SA)")
                    time.sleep(delay_s)

                time.sleep(delay_s * 0.5)

                try:
                    linewidth_khz = self.sa.measure_linewidth_kHz()
                except Exception as e:
                    self.log(f"[Runner] 组2 SA 读取失败 (current {cur} mA): {e}")
                    continue

                try:
                    self._append_summary(save_path, cur, temp_C, linewidth_khz, test_group=2, summary_filename=summary_filename)
                except Exception as e:
                    self.log(f"[Runner] 组2 写入汇总失败: {e}")

                peaks_curr.append(cur)
                peaks_lw.append(linewidth_khz)
                self.log(f"[Runner] 组2 {int(cur)}mA @ {temp_C:.2f}°C -> 线宽 {linewidth_khz:.6f} kHz")

            except Exception as e:
                self.log(f"[Runner] 组2 电流 {cur} mA 处理失败: {e}")
                continue

        if peaks_curr:
            self._plot_xy_curve(
                peaks_curr, peaks_lw,
                xlabel="电流(mA)", ylabel="线宽(kHz)",
                title=f"{temp_C:.2f}°C下电流-线宽关系",
                out_dir=save_path, prefix="电流线宽关系图",
                invert_x=False, save_csv=False,
                extra_cols={"Temperature_C": [f"{temp_C:.2f}"] * len(peaks_curr)}
            )
        else:
            self.log("[Runner] 组2 没有采集到线宽数据，跳过作图")

# -------------------------
# GUI (mostly unchanged, uses SA instead of OSA)
# -------------------------
class CT_L_GUI:
    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
        else:
            self.root= tk.Toplevel(parent)
        self.root.title('电流温度_线宽')
        self.root.resizable(True, True)
        self.set_center(1490, 1180)

        self.params = {
            "osa_ip": "192.168.29.11",
            "current_mA": 360.0,
            "t_start": 36.0,
            "t_stop": 15.0,
            "t_step": 1.0,
            "center_nm": 1550.0,
            "span_nm": 5.0,
            "laser_exe_path": r"C:\PTS\qijian\上位机软件\Preci_Semi\Preci-Seed.exe",
            "save_path": r"C:\PTS\qijian\CT_L",
            "group2_temp_C": 25.0,
            "group2_start_mA": 400.0,
            "group2_stop_mA": 0,
            "group2_step_mA": 5.0,
            "group1_delay_s": 5,
            "group2_delay_s": 2,
            "group1_summary_filename": "Test1_summary",
            "group2_summary_filename": "Test2_summary",
            "fine_center_C": 25.0,
            "fine_range_C": 1.0,
        }
        self.param_labels = {
            "laser_exe_path": "软件路径",
            "osa_ip": "IP地址",
            "current_mA": "电流 (mA)",
            "t_start": "初始温度 (℃)",
            "t_stop": "终止温度 (℃)",
            "t_step": "温度温度 (℃)",
            "save_path": "保存路径",
            "group2_temp_C": "组2 固定温度 (℃)",
            "group2_start_mA": "组2 初始电流 (mA)",
            "group2_stop_mA": "组2 终止电流 (mA)",
            "group2_step_mA": "组2 步进电流 (mA)",
            "group1_delay_s": "组1 温度稳定时间 (秒)",
            "group2_delay_s": "组2 电流稳定时间 (秒)",
            "group1_summary_filename": "组1文件名",
            "group2_summary_filename": "组2文件名"
        }

        self.create_widgets()
        self.laser: Optional[LaserController] = None
        self.sa: Optional[SpectrumAnalyzerController] = None
        self.runner: Optional[TestRunner] = None
        self.runner_thread: Optional[threading.Thread] = None
        self.group1_running = False
        self.group2_running = False

    def set_center(self, width: int, height: int):
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        posx = (screenwidth - width) // 2
        posy = (screenheight - height) // 2
        self.root.geometry(f'{width}x{height}+{posx}+{posy}')
    
    def create_widgets(self):
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        param_frame = tk.LabelFrame(main_container, text="参数设置", padx=8, pady=8)
        param_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.entries: Dict[str, tk.Entry] = {}
        connect_frame = tk.LabelFrame(param_frame, text="连接与地址", padx=8, pady=8)
        connect_frame.pack(fill=tk.X, padx=6, pady=4)

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

        group1_frame = tk.LabelFrame(param_frame, text="第一组测试", padx=6, pady=6)
        group1_frame.pack(fill="x", padx=6, pady=4)

        self._add_param_entry(group1_frame, "t_start", "初始温度:", self.params.get("t_start", 20.0), row=0)
        self._add_param_entry(group1_frame, "t_stop", "终止温度:", self.params.get("t_stop", 40.0), row=1)
        self._add_param_entry(group1_frame, "t_step", "步进温度:", self.params.get("t_step", 0.5), row=2)
        self._add_param_entry(group1_frame, "fine_center_C", "精测中心:", self.params.get("fine_center_C", 25.0), row=3)
        self._add_param_entry(group1_frame, "fine_range_C", "精测范围:", self.params.get("fine_range_C", 1.0), row=4)

        self._add_param_entry(group1_frame, "current_mA", "固定电流:", self.params.get("current_mA", 360.0), row=5)
        self._add_param_entry(group1_frame, "group1_delay_s", "稳定时间:", self.params.get("group1_delay_s", 5), row=6)
        self._add_param_entry(group1_frame, "group1_summary_filename", "保存文件名", self.params.get("group1_summary_filename", "Test1_summary.csv"), row=7)
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
    
        group2_frame = tk.LabelFrame(param_frame, text="第二组测试", padx=6, pady=6)
        group2_frame.pack(fill="x", padx=6, pady=4)
    
        self._add_param_entry(group2_frame, "group2_start_mA", "初始电流:", self.params.get("group2_start_mA", 400.0), row=0)
        self._add_param_entry(group2_frame, "group2_stop_mA", "终止电流:", self.params.get("group2_stop_mA", 0.5), row=1)
        self._add_param_entry(group2_frame, "group2_step_mA", "步进电流:", self.params.get("group2_step_mA", 5.0), row=2)
        self._add_param_entry(group2_frame, "group2_temp_C", "测试温度:", self.params.get("group2_temp_C", 25.0), row=3)
        self._add_param_entry(group2_frame, "group2_delay_s", "稳定时间:", self.params.get("group2_delay_s", 2), row=4)
        self._add_param_entry(group2_frame, "group2_summary_filename", "保存文件名:", self.params.get("group2_summary_filename", "Test2_summary.csv"), row=5)
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
            exe_path = p["laser_exe_path"]
            if not exe_path:
                messagebox.showerror("错误", "请先设置软件路径")
                return
            def _open_laser_thread():
                try:
                    self.laser = LaserController(exe_path=exe_path, window_title=r"Preci-Semi-Seed", log_func=self.log)
                    self.laser.connect()
                    self.log("[上位机] 已成功打开或连接到上位机软件")
                except Exception as e:
                    error_msg = f"[错误] 打开上位机软件失败: {e}"
                    self.log(error_msg)
                    self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                    self.laser = None
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
                    p[k] = self.params[k]
            except Exception:
                p[k] = float(self.params[k]) if k not in ("laser_exe_path", "osa_ip", "save_path", "group1_summary_filename", "group2_summary_filename") else self.params[k]
        return p

    def show_image_popup(self, img_path, title="测试完成 - 截图预览"):
        win = tk.Toplevel(self.root)
        win.title(title)
        try:
            img = Image.open(img_path)
        except Exception as e:
            self.log(f"[错误] 无法打开图片: {e}")
            messagebox.showerror("错误", f"无法打开图片: {e}")
            return
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        max_w, max_h = int(sw * 0.8), int(sh * 0.8)
        disp_img = img
        if img.width > max_w or img.height > max_h:
            scale = min(max_w / img.width, max_h / img.height)
            new_size = (int(img.width * scale), int(img.height * scale))
            disp_img = img.resize(new_size, Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(disp_img)
        win.img = img
        win.disp_img = disp_img
        win.img_tk = img_tk
        btn_frame = tk.Frame(win)
        btn_frame.pack(side=tk.TOP, pady=8)
        def save_img():
            save_path = filedialog.asksaveasfilename(
                defaultextension=".bmp",
                filetypes=[("BMP 文件", "*.bmp"), ("PNG 文件", "*.png"), ("所有文件", "*.*")],
                title="保存图片"
            )
            if save_path:
                win.img.save(save_path)
                messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")
        save_btn = tk.Button(btn_frame, text="保存图片", command=save_img)
        save_btn.pack()
        lbl = tk.Label(win, image=win.img_tk)
        lbl.pack(padx=6, pady=6)

    def diag_connect_and_query(self):
        ip_addr = self.entries["osa_ip"].get().strip()
        if not ip_addr:
            messagebox.showerror("错误", "请在诊断面板填写 IP 地址")
            return
        visa_addr = f"TCPIP0::{ip_addr}::INSTR"
        try:
            sa = SpectrumAnalyzerController(resource=visa_addr, log_func=self.log)
            sa.connect()
            idn = sa.query_idn()
            # try to set ASCII if supported
            try:
                sa.inst.write(":FORMat:DATA ASCII")
                time.sleep(0.2)
            except Exception:
                pass
            fmt = sa.query_format()
            self.log(f"[Diag] 连接成功, IDN={idn}, FORMAT={fmt}")
            self.sa = sa
        except Exception as e:
            self.log(f"[Diag] 连接/查询失败: {e}")
            messagebox.showerror("错误", f"诊断失败: {e}")

    def start_group1(self):
        p = self.get_params()
        self.btn_group1_start.config(state=tk.DISABLED)
        self.btn_group1_stop.config(state=tk.NORMAL)
        self.group1_running = True
        try:
            if not self.laser:
                self.laser = LaserController(exe_path=p["laser_exe_path"], window_title=r"Preci-Semi-Seed", log_func=self.log)
                try:
                    self.laser.connect()
                except Exception as e:
                    self.log(f"[错误] 激光控制软件连接失败: {e}")
                    if not messagebox.askyesno("警告", "激光控制软件连接失败，是否继续仅使用分析仪?"):
                        self.btn_group1_start.config(state=tk.NORMAL)
                        self.btn_group1_stop.config(state=tk.DISABLED)
                        self.group1_running = False
                        return
                    else:
                        self.laser = None

            if not self.sa:
                visa_address = f"TCPIP0::{p['osa_ip']}::INSTR"
                self.sa = SpectrumAnalyzerController(resource=visa_address, log_func=self.log)
                self.sa.connect()

            if not self.runner:
                self.runner = TestRunner(self.laser, self.sa, log_func=self.log)
            else:
                self.runner._stop = False

            def target():
                try:
                    self.runner.fine_center_C = p.get("fine_center_C", None)
                    self.runner.fine_range_C = p.get("fine_range_C", None)
                    self.runner.run_group1(
                        start_temp=p["t_start"],
                        end_temp=p["t_stop"],
                        step=p["t_step"],
                        save_path=p["save_path"],
                        delay_s=p["group1_delay_s"],
                        summary_filename=p["group1_summary_filename"],
                        current_mA=p["current_mA"]
                    )
                    img_path = self.runner.plot_group1_linewidth_vs_temperature(
                        p["save_path"], 
                        summary_filename=p["group1_summary_filename"]
                    )
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
            # try:
            #     self.log("[初始化] 正在清空共享文件夹和仪器内部文件夹...")
            #     # ================= 清空电脑共享目录 =================
            #     local_dir = r"\\192.168.29.9\PTS\zhongzi\CT_L\FSV3004"
            #     if os.path.exists(local_dir):
            #         for f in os.listdir(local_dir):
            #             fp = os.path.join(local_dir, f)
            #             try:
            #                 if os.path.isfile(fp) or os.path.islink(fp):
            #                     os.remove(fp)
            #                 elif os.path.isdir(fp):
            #                     import shutil
            #                     shutil.rmtree(fp)
            #             except Exception as e:
            #                 self.log(f"[警告] 删除 {fp} 失败: {e}")
            #     # ================= 清空仪器内部目录 =================
            #     try:
            #         ip = "192.168.29.11"
            #         rm = pyvisa.ResourceManager()
            #         inst = rm.open_resource(f"TCPIP0::{ip}::5025::SOCKET")
            #         inst.write("MMEM:MDIR 'C:\\PTS\\CT_L'")  # 确保路径存在
            #         inst.write("MMEM:DEL 'C:\\PTS\\CT_L\\*.*'")
            #         # inst.query("*OPC?")  # 可选等待完成
            #         inst.close()
            #         rm.close()
            #     except Exception as e:
            #         self.log(f"[警告] 仪器文件夹清理失败: {e}")

            #     self.log("[初始化] 仪器与共享文件夹清理完成。")

            # except Exception as e:
            #     self.log(f"[错误] 文件夹清理过程中出现异常: {e}")

        except Exception as e:
            self.log(f"[错误] 启动第一组测试失败: {e}")
            messagebox.showerror("错误", f"启动第一组测试失败: {e}")
            self.btn_group1_start.config(state=tk.NORMAL)
            self.btn_group1_stop.config(state=tk.DISABLED)
            self.group1_running = False

    def stop_group1(self):
        if self.runner and self.group1_running:
            try:
                self.runner.stop()
                self.log("[主] 第一组测试停止信号已发送")
            except Exception as e:
                self.log(f"[错误] 停止第一组测试失败: {e}")
        else:
            self.log("[主] 没有正在运行的第一组测试")

    def start_group2(self):
        p = self.get_params()
        self.btn_group2_start.config(state=tk.DISABLED)
        self.btn_group2_stop.config(state=tk.NORMAL)
        self.group2_running = True
        try:
            if not self.laser:
                self.laser = LaserController(exe_path=p["laser_exe_path"], window_title=r"Preci-Semi-Seed", log_func=self.log)
                try:
                    self.laser.connect()
                except Exception as e:
                    self.log(f"[错误] 激光控制软件连接失败: {e}")
                    if not messagebox.askyesno("警告", "激光控制软件连接失败，是否继续仅使用分析仪?"):
                        self.btn_group2_start.config(state=tk.NORMAL)
                        self.btn_group2_stop.config(state=tk.DISABLED)
                        self.group2_running = False
                        return
                    else:
                        self.laser = None

            if not self.sa:
                visa_address = f"TCPIP0::{p['osa_ip']}::INSTR"
                self.sa = SpectrumAnalyzerController(resource=visa_address, log_func=self.log)
                self.sa.connect()

            if not self.runner:
                self.runner = TestRunner(self.laser, self.sa, log_func=self.log)
            else:
                self.runner._stop = False

            def target():
                try:
                    img_path = None
                    self.runner.run_group2(
                        start_mA=p["group2_start_mA"],
                        step_mA=p["group2_step_mA"],
                        stop_mA=p["group2_stop_mA"],
                        temp_C=p["group2_temp_C"],
                        save_path=p["save_path"],
                        delay_s=p["group2_delay_s"],
                        summary_filename=p["group2_summary_filename"]
                    )
                    import glob
                    pattern = os.path.join(p["save_path"], "电流线宽关系图_*.png")
                    group2_files = glob.glob(pattern)
                    if group2_files:
                        group2_files.sort(key=os.path.getmtime, reverse=True)
                        img_path = group2_files[0]
                        self.log(f"[Runner] 找到最新的第二组测试图像: {img_path}")
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
            if not self.sa:
                visa_address = f"TCPIP0::{p['osa_ip']}::INSTR"
                self.sa = SpectrumAnalyzerController(resource=visa_address, log_func=self.log)
                self.sa.connect()

            try:
                self.sa.sweep_and_fetch()
            except Exception:
                pass

            freqs, powers = self.sa.fetch_trace()
            npoints = len(powers)
            self.log(f"[单次] 读取到 {npoints} 点")

            save_base = p["save_path"]
            if os.path.isdir(save_base) or save_base.endswith(os.sep):
                fig_dir = save_base
            else:
                fig_dir = os.path.dirname(save_base) or "."
            ensure_dir(fig_dir)
            fig_path = os.path.join(fig_dir, "single_scan.png")

            plt.figure(figsize=(8, 4))
            if freqs is not None and len(freqs) == npoints:
                plt.plot(freqs, powers)
                plt.xlabel("Frequency (Hz)")
            else:
                plt.plot(np.arange(npoints), powers)
                plt.xlabel("Point")
            plt.title("Single Scan")
            plt.ylabel("Power (dBm)")
            plt.tight_layout()
            plt.savefig(fig_path)
            plt.close()
            self.log(f"[单次] 图像保存到 {fig_path}")

            csv_fn = os.path.join(fig_dir, "single_scan.csv")
            with open(csv_fn, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Frequency_Hz", "Power_dBm"])
                for x, y in zip(freqs, powers):
                    w.writerow([f"{float(x):.3f}", f"{float(y):.6f}"])
            self.log(f"[单次] 光谱 CSV 保存到 {csv_fn}")

        except Exception as e:
            self.log(f"[错误] 单次扫描失败: {e}\n{traceback.format_exc()}")
            messagebox.showerror("错误", f"单次扫描失败: {e}")

    def run(self):
        if self.root.winfo_exists():
            self.root.mainloop()

if __name__ == "__main__":
    gui = CT_L_GUI()
    gui.run()
