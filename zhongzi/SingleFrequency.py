import os
import re
import csv
import time
import math
import threading
import pyvisa
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog

import matplotlib
matplotlib.use('Agg')  # 后端绘图，不阻塞 GUI
import matplotlib.pyplot as plt

from PIL import Image, ImageTk

# ===============  上位机控制（pywinauto）  ===============
try:
    from pywinauto.application import Application
    from pywinauto import timings
    PYW_AVAILABLE = True
except Exception:
    PYW_AVAILABLE = False


from pywinauto.application import Application
import time


class LaserController:
    def __init__(self, exe_path, window_title=".*Preci-Seed.*", log_func=print):
        self.exe_path = exe_path
        self.window_title = window_title
        self.app = None
        self.win = None
        self.log = log_func

    def start_or_connect(self, timeout=15.0):
        if not PYW_AVAILABLE:
            raise RuntimeError("未安装 pywinauto，请先 pip install pywinauto")
        try:
            # 优先连接已运行实例
            self.app = Application(backend="uia").connect(title_re=self.window_title, timeout=5)
            self.win = self.app.window(title_re=self.window_title)
            self.win.set_focus()
            self.log("[上位机] 已连接到运行中的窗口")
        except Exception:
            self.log("[上位机] 未找到运行实例，尝试启动…")
            self.app = Application(backend="uia").start(cmd_line=f'"{self.exe_path}"')
            self.app.connect(title_re=self.window_title, timeout=timeout)
            self.win = self.app.window(title_re=self.window_title)
            self.win.wait("ready", timeout=timeout)
            self.win.set_focus()
            self.log("[上位机] 已启动并连接")
        timings.wait_until_passes(5, 0.5, lambda: self.win.exists() and self.win.is_visible())
        return True

    # ================= 波长控制 =================
    def get_wavelength_nm(self) -> float | None:
        """读取当前中心波长 (auto_id=label_Wavelength)"""
        if self.win is None:
            # 未连接/未获取到窗口句柄
            self.log("[上位机] 读取波长失败：未连接窗口 (self.win is None)")
            return None
        try:
            label = self.win.child_window(auto_id="label_Wavelength", control_type="Text")
            txt = label.window_text()
            return float(txt)
        except Exception as e:
            self.log(f"[错误] 读取波长失败: {e}")
            return None

    def set_wavelength_nm(self, val_nm: float):
        """设置中心波长 (auto_id=txtWavelength + Apply按钮)"""
        try:
            edit = self.win.child_window(auto_id="textBox_Wavelength", control_type="Edit")
            edit.set_edit_text(f"{val_nm:.6f}")
            btn = self.win.child_window(title="Set", control_type="Button")
            btn.click()
            self.log(f"[上位机] 已设置波长: {val_nm:.6f} nm")
            time.sleep(1.0)
        except Exception as e:
            self.log(f"[错误] 设置波长失败: {e}")

    # ================= 电流控制 =================
    def get_current_mA(self) -> float | None:
        """读取当前工作电流 (auto_id=Label_Current)"""
        try:
            edit = self.win.child_window(auto_id="Label_current", control_type="Text")
            txt = edit.window_text()
            return float(txt)
        except Exception as e:
            self.log(f"[错误] 读取电流失败: {e}")
            return None

    def set_current_mA(self, val_mA: float):
        """设置电流 (auto_id=txtCurrent + Apply按钮)"""
        try:
            edit = self.win.child_window(auto_id="textBox_Current", control_type="Edit")
            edit.set_edit_text(f"{val_mA:.2f}")
            btn = self.win.child_window(title="Set", control_type="Button")
            btn.click()
            self.log(f"[上位机] 已设置电流: {val_mA:.2f} mA")
            time.sleep(1.0)
        except Exception as e:
            self.log(f"[错误] 设置电流失败: {e}")

    # ================= 温度控制 =================
    def get_temperature_c(self) -> float | None:
        """读取当前工作温度 (auto_id=Label_Temperature)"""
        try:
            edit = self.win.child_window(auto_id="Label_Temperature", control_type="Text")
            txt = edit.window_text()
            return float(txt)
        except Exception as e:
            self.log(f"[错误] 读取温度失败: {e}")
            return None

    def set_temperature_c(self, val_c: float):
        """设置温度 (auto_id=TextBox_Temperature + Apply按钮)"""
        try:
            edit = self.win.child_window(auto_id="TextBox_Temperature", control_type="Edit")
            edit.set_edit_text(f"{val_c:.2f}")
            edit.type_keys("{ENTER}")
            # btn = self.win.child_window(title="Set", control_type="Button")
            # btn.click()
            self.log(f"[上位机] 已设置温度: {val_c:.2f} °C")
            time.sleep(1.0)
        except Exception as e:
            self.log(f"[错误] 设置温度失败: {e}")

# ===============  频谱仪控制 & 峰值检测  ===============
class SingleFrequency:
    def __init__(self, ip, timeout_s=60.0, log=print, cmd_map=None):
        self.ip = ip
        self.timeout_s = timeout_s
        self.log = log
        self.rm = None
        self.sa = None
        self.last_rbw_hz = None
        self.last_vbw_hz = None
        self.CMD = {
            'idn': '*IDN?\n',
            'abort': ':ABORt\n',
            'opc': '*OPC?\n',
            'f_center': ':SENSe:FREQuency:CENTer {hz}\n',
            'f_span':   ':SENSe:FREQuency:SPAN {hz}\n',
            'f_start':  ':SENSe:FREQuency:STARt {hz}\n',
            'f_stop':   ':SENSe:FREQuency:STOP {hz}\n',
            'q_start':  ':SENSe:FREQuency:STARt?\n',
            'q_stop':   ':SENSe:FREQuency:STOP?\n',
            'rbw': ':SENSe:BANDwidth:RESolution {hz}\n',
            'vbw': ':SENSe:BANDwidth:VIDeo {hz}\n',
            'rbw?': ':SENSe:BANDwidth:RESolution?\n',
            'vbw?': ':SENSe:BANDwidth:VIDeo?\n',
            'sweep_points?': ':SWEep:POINts?\n',
            'trace_mode_write': ':TRACe:MODE WRITe\n',
            'trace_mode_max':   ':TRACe:MODE MAXHold\n',
            'trace_clear':      ':TRACe:CLEAr\n',
            'trace_data':       ':TRACe:DATA? TRACE1\n',
            'init_once': ':INITiate:IMMediate\n',
            'avg_on':  ':AVERage:STATe ON\n',
            'avg_off': ':AVERage:STATe OFF\n',
            'avg_count': ':AVERage:COUNt {n}\n',
        }
        if cmd_map:
            self.CMD.update(cmd_map)

    def open(self):
        self.rm = pyvisa.ResourceManager()
        #self.sa = self.rm.open_resource(f"TCPIP::{self.ip}::INSTR")
        self.sa = self.rm.open_resource(f"TCPIP::{self.ip}::5025::SOCKET")
        self.sa.timeout = int(self.timeout_s * 1000)
        self.sa.write_termination = '\n'
        self.sa.read_termination = '\n'
        
        idn = self.query(self.CMD['idn']).strip()
        self.log(f"[频谱仪] 已连接：{idn}")

        self.write(":CALC:MARK1:MODE NORM")        # 普通标记模式（必须）
        self.write(":CALC:MARK1 ON")                # 打开标记1显示
        self.write(":CALC:MARK1:FUNC NOIS")         # 开启噪声标记（手册第111页精确命令）
        self.log("[频谱仪] 噪声标记已开启 → Nrs dBm/Hz")
        self.write(":CALC:MARK1:MAX")              # 立即跳到最高峰（最实用）
        self.write(":SWE:TYPE:AUTO:RUL DRAN")        # 打开动态范围优先
        self.log("[频谱仪] 动态范围优先已开启")
        self.write(":UNIT:POW DBM")          # 纵轴刻度单位设置为DBM
        self.log("[频谱仪] 纵轴刻度单位已设置为DBM")

        return idn

    def close(self):
        try:
            if self.sa:
                self.sa.close()
        finally:
            if self.rm:
                self.rm.close()

    def write(self, scpi):
        self.sa.write(scpi)

    def query(self, scpi):
        return self.sa.query(scpi)

    def opc(self, label='操作'):
        self.query(self.CMD['opc'])
        self.log(f"[频谱仪] {label} 完成")

    def set_freq_span(self, center=None, span=None, start=None, stop=None):
        if center is not None:
            self.write(self.CMD['f_center'].format(hz=float(center)))
            #time.sleep(0.5)  # 新增
        if span is not None:
            self.write(self.CMD['f_span'].format(hz=float(span)))
            #time.sleep(0.5)
        if start is not None:
            self.write(self.CMD['f_start'].format(hz=float(start)))
            #time.sleep(0.5)
        if stop is not None:
            self.write(self.CMD['f_stop'].format(hz=float(stop)))
            #time.sleep(0.5)

    def set_bw(self, rbw_hz, vbw_hz=None):
        self.write(self.CMD['rbw'].format(hz=float(rbw_hz)))
        time.sleep(0.5)
        self.last_rbw_hz = float(rbw_hz)
        if vbw_hz is not None:
            self.write(self.CMD['vbw'].format(hz=float(vbw_hz)))
            time.sleep(0.5)
            self.last_vbw_hz = float(vbw_hz)
        # 查询实际 RBW（容错：去单位）
        try:
            q = self.CMD.get('rbw?')
            if q:
                resp = self.query(q)
                num = re.findall(r"[-+]?\d*\.?\d+", str(resp))
                if num:
                    self.last_rbw_hz = float(num[0])
        except Exception:
            pass

    def set_avg(self, on=True, count=4):
        self.write(self.CMD['avg_on' if on else 'avg_off'])
        try:
            self.write(self.CMD['avg_count'].format(n=int(count)))
        except Exception:
            pass

    def set_trace_mode(self, max_hold=False):
        self.write(self.CMD['trace_mode_max' if max_hold else 'trace_mode_write'])
        time.sleep(0.5)

    def set_sweep_type(self, sweep_type: str):
        """
        设置扫描优先级
        sweep_type: 'SPD'（速度优先）或 'DYN'（动态范围优先）
        """
        try:
            self.write(f":SWE:TYPE {sweep_type}")
            time.sleep(0.2)
            self.log(f"[频谱仪] 设置扫描优先级为: {sweep_type}")
        except Exception as e:
            self.log(f"[错误] 设置扫描优先级失败: {e}")

    def set_sweep_time(self, sweep_time_s: float):
        """
        设置扫描时间（秒）
        """
        try:
            self.write(f":SWE:TIME {sweep_time_s}")
            time.sleep(0.2)
            #self.log(f"[频谱仪] 设置扫描时间为: {sweep_time_s}s")
        except Exception as e:
            self.log(f"[错误] 设置扫描时间失败: {e}")

    def sweep_once(self, label='扫频'):
        self.write(self.CMD['trace_clear'])
        self.write(self.CMD['init_once'])
        self.log("[频谱仪] 已触发一次扫频")
        time.sleep(2.5)
        self.opc(label)
        #time.sleep(1.4)

    # def set_detector(self, mode: str = "RMS"):
    #     """设置检波器模式 (POS / NEG / SAMP / RMS)"""
    #     self.write(f":DET {mode}")
    def set_detector(self, mode: str = "RMS", trace: int = 1):
        """
        设置检波器模式
        常见模式: POSitive, NEGative, SAMPle, RMS
        """
        try:
            self.write(f":DETector:FUNCtion{trace} {mode}")
            time.sleep(0.5)
            self.log(f"[频谱仪] 已设置检波器模式: {mode}")
        except Exception as e:
            self.log(f"[错误] 设置检波器失败: {e}")

    def get_trace_xy(self):
        try:
            # 添加调试信息
            raw_data = self.sa.query_ascii_values(self.CMD['trace_data'])
            
            y_dbm = np.array(raw_data) # 直接转为numpy数组

            f_start = float(self.query(self.CMD['q_start']))
            f_stop = float(self.query(self.CMD['q_stop']))
            n = int(float(self.query(self.CMD['sweep_points?'])))
            #self.log(f"[调试] 频率范围: {f_start/1e9:.3f} GHz ~ {f_stop/1e9:.3f} GHz, 点数: {n}")
            
            x = np.linspace(f_start, f_stop, num=n)
            return x, y_dbm
        except Exception as e:
            #self.log(f"[错误] 读取谱线失败：{e}")
            # 如果出现异常，返回一个包含调试信息的数组
            try:
                f_start = float(self.query(self.CMD['q_start']))
                f_stop = float(self.query(self.CMD['q_stop']))
                n = int(float(self.query(self.CMD['sweep_points?'])))
                x = np.linspace(f_start, f_stop, num=n)
                # 返回一个全-100 dBm的数组作为占位符
                return x, np.full(n, -100.0)
            except:
                # 如果频率范围也获取失败，返回空数组
                return np.array([]), np.array([])

    def sweep_continuous_on(self, label="连续粗扫"):
        """开启连续扫描（:INITiate:CONTinuous ON）。尽量先清空 trace，以避免历史峰污染。"""
        try:
            # 清除之前的 trace（容错取 CMD）
            self.write(self.CMD.get('trace_clear', ':TRACe:CLEAr\n'))
            time.sleep(0.5)
        except Exception:
            pass
        try:
            self.write(':INITiate:CONTinuous ON\n')
            time.sleep(0.5)
            # 小延时让仪器进入连续模式（可选）
            self.query_opc(timeout=5000)
            self.log(f"[频谱仪] 开启连续扫: {label}")
        except Exception as e:
            self.log(f"[频谱仪] 开启连续扫失败: {e}")
            raise

    def sweep_continuous_off(self, label="停止连续粗扫"):
        """关闭连续扫描（:INITiate:CONTinuous OFF）。"""
        try:
            self.write(':INITiate:CONTinuous OFF\n')
            # 小延时保证状态切换完成
            time.sleep(0.5)
            self.log(f"[频谱仪] 关闭连续扫: {label}")
        except Exception as e:
            self.log(f"[频谱仪] 关闭连续扫失败: {e}")

    def query_opc(self, timeout: float | None = None) -> bool:
        """
        等待操作完成 (*OPC?)。
        timeout: 秒 (可选)，如果不传就用 self.timeout_s
        """
        try:
            if self.sa is None:
                raise RuntimeError("未连接频谱仪 (self.sa is None)")

            # 设置超时（pyvisa 的单位是 ms）
            if timeout is not None:
                self.sa.timeout = int(timeout * 1000)
            else:
                self.sa.timeout = int(self.timeout_s * 1000)

            resp = self.sa.query(self.CMD.get("opc", "*OPC?\n"))
            return resp.strip() == "1"

        except Exception as e:
            self.log(f"[频谱仪] query_opc 失败: {e}")
            return False

class PeakDetector:
    def __init__(self, thresh_db=1.0, prom_db=1.0, guard=10, log_func=print):
        self.thresh_db = float(thresh_db)
        self.prom_db = float(prom_db)
        self.guard = int(guard)
        self.log = log_func

    def find(self, x, y_dbm):
        if len(y_dbm) < 2 * self.guard + 1:
            return []
        
        # 改进的噪声估计：使用频谱边缘的噪声，更准确
        # 取频谱前10%和后10%的数据点计算噪声平均值
        edge_points = int(len(y_dbm) * 0.1)
        if edge_points < 10:  # 确保至少有10个点用于噪声估计
            edge_points = 10
        
        # 从频谱两端取点计算噪声
        edge_data = np.concatenate([y_dbm[:edge_points], y_dbm[-edge_points:]])
        noise = float(np.mean(edge_data))
        
        peaks = []
        g = self.guard
        
        # 调试信息：显示噪声水平和检测参数
        #self.log(f"[峰值检测] 噪声水平: {noise:.2f} dBm, 阈值: {self.thresh_db} dB, 显著性: {self.prom_db} dB")
        
        # 对于非常窄的峰，使用更小的保护带
        # 如果保护带大于1，尝试使用更小的保护带进行局部最大值判断
        narrow_guard = max(1, int(g / 2))  # 缩小保护带以检测更窄的峰
        
        for i in range(g, len(y_dbm) - g):
            y = float(y_dbm[i])
            
            # 检查当前点是否是局部最大值（使用缩小的保护带）
            is_local_max = True
            for j in range(1, narrow_guard + 1):
                if y <= float(y_dbm[i-j]) or y <= float(y_dbm[i+j]):
                    is_local_max = False
                    break
            
            # 计算左右邻域平均值（用于显著性判断）
            # 使用稍大的邻域来更准确地评估局部背景
            left_nb = y_dbm[max(0, i-g-2):i]
            right_nb = y_dbm[i+1:min(len(y_dbm), i+g+3)]
            left_mean = float(np.mean(left_nb)) if len(left_nb) else noise
            right_mean = float(np.mean(right_nb)) if len(right_nb) else noise
            
            # 计算局部背景噪声
            local_noise = min(left_mean, right_mean, noise)
            
            # 调试信息：显示每个可能的峰值点
            # if y - noise >= 0.5:  # 只记录接近噪声阈值的点
            #     self.log(f"[峰值调试] 频率: {x[i]/1e9:.3f} GHz, 功率: {y:.2f} dBm, "
            #             f"局部最大值: {is_local_max}, 局部噪声: {local_noise:.2f} dBm, "
            #             f"噪声差: {y - noise:.2f} dB, 显著性差: {y - max(left_mean, right_mean):.2f} dB")
            
            # 峰值检测条件
            if (
                is_local_max and  # 必须是局部最大值
                (y - local_noise >= self.thresh_db) and  # 高于局部噪声阈值
                (y - max(left_mean, right_mean) >= self.prom_db * 0.8)  # 稍微降低显著性要求
            ):
                peaks.append((float(x[i]), y, local_noise))
                self.log(f"[峰值检测] 检测到峰值: {x[i]/1e9:.3f} GHz, 功率: {y:.2f} dBm")
        
        return peaks

    def save_csv_png(self, x, y, peaks, out_dir, name, rbw_hz=1e3):
        os.makedirs(out_dir, exist_ok=True)
        csv_path = os.path.join(out_dir, f'{name}.csv')
        with open(csv_path, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['Frequency(Hz)', 'Power(dBm)'])
            for xi, yi in zip(x, y):
                w.writerow([xi, yi])
        peak_csv = os.path.join(out_dir, f'{name}_peaks.csv')
        with open(peak_csv, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['PeakFreq(Hz)', 'PeakPower(dBm)', 'NoiseFloor(dBm)'])
            for (fx, py, nb) in peaks:
                w.writerow([fx, py, nb])

        # 绘图
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 微软雅黑，支持中文
        plt.rcParams['axes.unicode_minus'] = False    # 正确显示负号
        png_path = os.path.join(out_dir, f'{name}.png')
        x_mhz = np.array(x) / 1e6

        fig, ax = plt.subplots(figsize=(12, 6))
        ax.set_facecolor('black')         # 坐标区背景设为黑色
        ax.plot(x_mhz, y, linewidth=1.2, color='yellow')  # 曲线设为黄色
        ax.set_xlabel('Frequency (MHz)', fontsize=18)
        ax.set_ylabel('Power (dBm)', fontsize=18)
        ax.margins(x=0)
        #ax.set_title(name)
        ax.grid(linestyle=':', linewidth=0.8, alpha=0.6, color='white')
        # 坐标轴刻度设置
        ax.tick_params(axis='x', colors='black', size=7, labelsize=15)
        ax.tick_params(axis='y', colors='black', size=7, labelsize=15)

        if peaks:
            # 先找功率最大的峰
            main_peak = max(peaks, key=lambda p: p[1])  # (freq, power, noise)
            for (fx, py, nb) in peaks:
                fx_mhz = fx / 1e6
                # 所有峰都画竖线
                ax.axvline(fx_mhz, linestyle='--', linewidth=0.8, color='gray', alpha=0.8)

            # 只给最大峰做标注
            fx, py, nb = main_peak
            fx_mhz = fx / 1e6
        
            lines = [
                f"单频: {fx_mhz:.2f} MHz",
                f"Y: {py:.2f} dBm",
            ]
            txt = "\n".join(lines)

            ax.annotate(
                txt,
                xy=(fx_mhz, py),
                xytext=(8, -6),
                textcoords='offset points',
                ha='left',
                va='top',
                fontsize=12,
                color='black',
                bbox=dict(boxstyle='round,pad=0.3', fc='white', alpha=0.6),
                arrowprops=dict(arrowstyle='->', color='white', lw=0.6)
            )

        plt.tight_layout()
        fig.savefig(png_path, dpi=600)
        plt.close(fig)
        return csv_path, png_path, peak_csv


# ===============  GUI & 流程编排  ===============
class SingleFrequencyGUI:
    def __init__(self, parent=None):
        self.parent = parent
        
        # --- 核心修改：如果是集成模式，直接使用父控件作为 root ---
        if parent is None:
            self.root = tk.Tk()
            self.root.title("单频 - 独立模式")
            self.root.geometry("1480x930") 
            self.root.resizable(True, True)
        else:
            self.root = parent # <--- 修改点：直接使用父 Frame

        self.params_1um = {
            # 仪器 & 输出
            'IP地址': '192.168.7.15',
            '输出目录': r'C:\PTS\zhongzi\SingleFrequency\1.0μm',
    
            # 上位机
            '上位机路径': r'C:\PTS\zhongzi\SingleFrequency\shangweiji\Preci-Seed.exe',
            '窗口标题(正则)': r'Preci_Fiber_DFB_2Diode.*',

            # 温度参数
            '温度上限(°C)': 56.0,
            '温度下限(°C)': 20.0,
            '温度步长(°C)': 0.1,
            '温度变化频率(s)': 3.0,
            
            # 电流参数
            '电流上限(mA)': 600.0,
            '电流下限(mA)': 100.0,
            '电流步长(mA)': 5.0,
            '电流变化频率(s)': 5.0,

            # 扫描参数
            '细扫邻域点数': 10,
            '细扫峰值阈值(dB)': 5.0,
            '细扫邻域显著性(dB)': 5.0,
        }

        self.params_1_5um = {
            # 仪器 & 输出
            'IP地址': '192.168.7.15',
            '输出目录': r'C:\PTS\zhongzi\SingleFrequency\1.5μm',

            # 上位机
            '上位机路径': r'C:\PTS\zhongzi\SingleFrequency\shangweiji\Preci-Seed.exe',
            '窗口标题(正则)': r'Preci_Fiber_DFB_2Diode.*',

            # 温度参数
            '温度上限(°C)': 56.0,
            '温度下限(°C)': 20.0,
            '温度步长(°C)': 0.1,
            '温度变化频率(s)': 3.0,
            
            # 电流参数
            '电流上限(mA)': 1350.0,
            '电流下限(mA)': 450.0,
            '电流步长(mA)': 450.0,
            '电流变化频率(s)': 5.0,

            # 扫描参数
            '细扫邻域点数': 10,
            '细扫峰值阈值(dB)': 5.0,
            '细扫邻域显著性(dB)': 5.0,
        }

        self.test_type_var = tk.StringVar(value="1μm")
        # type_frame = tk.Frame(self.root)
        # type_frame.pack(fill=tk.X, padx=10, pady=4)
        # tk.Label(type_frame, text="测试类型:").pack(side=tk.LEFT)
        # type_combo = tk.OptionMenu(type_frame, self.test_type_var, "1μm", "1.5μm", command=self._on_test_type_change)
        # type_combo.pack(side=tk.LEFT)

        self._build_ui()
        self.stop_flag = threading.Event()
        self.pause_flag = threading.Event()
        self.worker = None

    # —— UI ——
    def _build_ui(self):
        # 创建主框架，分为左右两部分
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧框架 - 参数设置
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))
        
        # 右侧框架 - 运行日志
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 测试类型选择区（居中+外框）- 放在左侧
        type_labelframe = tk.LabelFrame(left_frame, text='测试项选择', padx=10, pady=10)
        type_labelframe.pack(fill=tk.X, pady=8)
        type_frame = tk.Frame(type_labelframe)
        type_frame.pack(expand=True)
        tk.Label(type_frame, text="测试类型:").pack(side=tk.LEFT, padx=6)
        type_combo = tk.OptionMenu(type_frame, self.test_type_var, "1μm", "1.5μm", command=self._on_test_type_change)
        type_combo.pack(side=tk.LEFT, padx=6)
        # 居中
        type_frame.pack(anchor='center')

        # 连接与地址设置区 - 放在左侧上方
        conn_frame = tk.LabelFrame(left_frame, text='连接与地址', padx=8, pady=8)
        conn_frame.pack(fill=tk.X, pady=6)
        
        # 参数设置区 - 放在左侧下方
        param_frame = tk.LabelFrame(left_frame, text='参数设置', padx=8, pady=8)
        param_frame.pack(fill=tk.X, pady=6)
        
        self.entries = {}
        self.param_labels = {}  # 新增：用于存储参数标签
        self.conn_frame = conn_frame  # 保存框架引用
        self.param_frame = param_frame  # 保存框架引用
        
        # 首次构建UI（使用1μm的参数）
        self._build_param_ui()
        
        # 按钮区域放在参数设置框下方，居中显示
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=8)
        
        # 创建一个内部框架来容纳按钮，实现居中
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')
        
        #tk.Button(inner_btn_frame, text='保存参数', bg="#4CAF50", command=self._save_params).pack(side='left', padx=6)
        tk.Button(inner_btn_frame, text='开始测试', bg="#4CAF50",fg= "#FFFFFF", command=self.start).pack(side='left', padx=6)
        # Pause / Resume button (will toggle between 暂停 and 继续)
        self.pause_btn = tk.Button(inner_btn_frame, text='暂停', bg="#FFA000", fg="#FFFFFF", command=self._toggle_pause)
        self.pause_btn.pack(side='left', padx=6)
        tk.Button(inner_btn_frame, text='停止测试', bg="#f44336", fg= "#FFFFFF", command=self.stop).pack(side='left', padx=6)

        # 运行日志区 - 放在右侧
        logf = tk.LabelFrame(right_frame, text='运行日志', padx=6, pady=6)
        logf.pack(fill=tk.BOTH, expand=True)
        self.log_box = tk.Text(logf)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _safe_log_append(self, text):
        self.log_box.insert(tk.END, text)
        self.log_box.see(tk.END)

    def log(self, msg):
        t = time.strftime('[%H:%M:%S]')
        self.root.after(0, lambda: self._safe_log_append(f"{t} {msg}\n"))

    def _save_params(self):
        for k, e in self.entries.items():
            v = e.get()
            try:
                val = float(v)
                self.params_1um[k] = int(val) if val.is_integer() else val
            except Exception:
                self.params_1um[k] = v
        self.log('[参数] 已更新')

    def _toggle_pause(self):
        """切换暂停/继续状态。"""
        if not hasattr(self, 'pause_flag'):
            self.pause_flag = threading.Event()
        if not self.pause_flag.is_set():
            # pause
            self.pause_flag.set()
            try:
                self.pause_btn.config(text='继续', bg='#4CAF50')
            except Exception:
                pass
            self.log('[用户] 已暂停，点击继续以恢复')
        else:
            # resume
            self.pause_flag.clear()
            try:
                self.pause_btn.config(text='暂停', bg='#FFA000')
            except Exception:
                pass
            self.log('[用户] 已继续，恢复运行')

    def _pause_point(self):
        """在长循环中调用来实现暂停：当 pause_flag 被设置时阻塞，直到清除或 stop_flag 被设置。
        不要在极短的循环里频繁调用以避免性能影响。
        """
        # 如果没有 pause_flag 则不阻塞
        if not hasattr(self, 'pause_flag'):
            return
        while self.pause_flag.is_set():
            # 每隔一段时间检查一次
            time.sleep(0.2)
            if self.stop_flag.is_set():
                # 如果已请求停止，抛出以让上层及时响应
                raise KeyboardInterrupt

    def start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo('提示', '测试已在进行中')
            return
        self._save_params()
        self.stop_flag.clear()
        self.worker = threading.Thread(target=self._run, daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_flag.set()
        self.log('[用户] 请求停止…')

    # —— 图片弹窗 ——
    def show_image_popup(self, image_path, title="结果预览"):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.resizable(False, False)
        pil_img = Image.open(image_path)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        max_w, max_h = int(sw * 0.8), int(sh * 0.8)
        disp_img = pil_img
        if pil_img.width > max_w or pil_img.height > max_h:
            scale = min(max_w / pil_img.width, max_h / pil_img.height)
            new_size = (int(pil_img.width * scale), int(pil_img.height * scale))
            disp_img = pil_img.resize(new_size, Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(disp_img)
        win.orig_img = pil_img
        win.img_tk = img_tk
        btn_frame = tk.Frame(win)
        btn_frame.pack(side=tk.TOP, fill='x', pady=8)

        def _save_img():
            save_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                     filetypes=[("PNG 文件", "*.png"), ("BMP 文件", "*.bmp"), ("所有文件", "*.*")],
                                                     title="保存图片")
            if save_path:
                try:
                    win.orig_img.save(save_path)
                    messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")
                except Exception as ex:
                    messagebox.showerror("保存失败", str(ex))
        tk.Button(btn_frame, text="保存图片", command=_save_img).pack()
        tk.Label(win, image=win.img_tk).pack(padx=6, pady=6)
        win.update_idletasks()
        w = win.winfo_width(); h = win.winfo_height()
        x = (sw - w) // 2; y = (sh - h) // 2
        win.geometry(f"+{x}+{y}")

    # —— 频谱扫描工具函数 ——
    def _coarse_scan_and_check(self, sa: 'SingleFrequency', p: dict,
                               tag: str, out_dir: str) -> tuple[bool, str | None]:
        """执行一次全带宽粗扫；若发现异常峰 -> 保存&弹窗 -> 返回(True, png)。否则返回(False, None)。"""
        if self.stop_flag.is_set():
            raise KeyboardInterrupt
        self.log('[粗扫] 全频带，最大保持…')
        sa.set_sweep_type('DYN') # 动态范围优先
        sa.set_sweep_time(5) # 5秒扫完
        sa.set_trace_mode(max_hold=True)

        # === 等待稳定阶段 ===
        sa.sweep_continuous_on("等待稳定")   # 开启连续扫，只让仪器扫，不取数据

        same = 0
        last_val = None
        tol = 0.001
        consec_ok = 3
        max_wait = 300.0
        interval = 0.01

        t0 = time.time()
        t0 = time.time()
        while time.time() - t0 < max_wait:
            # pause support
            if getattr(self, 'pause_flag', None) and self.pause_flag.is_set():
                self._pause_point()
            if self.stop_flag.is_set():
                break

            wl = self.lc.get_wavelength_nm()
            if wl is None:
                time.sleep(interval)
                continue

            if last_val is not None and abs(wl - last_val) < tol:
                same += 1
                if same >= consec_ok:
                    self.log(f"[等待稳定] 波长已稳定 (连续{consec_ok}次Δ<{tol}nm)")
                    break
            else:
                same = 0

            last_val = wl
            if same > 0:  # 只在有连续稳定计数时打印日志
                self.log(f"[等待稳定] 当前波长 {wl:.4f} nm, 已连续稳定 {same}/{consec_ok}")
            time.sleep(interval)

        # === 稳定后 ===
        try:
            sa.sweep_continuous_off("等待稳定结束")
            time.sleep(0.5)  # 等待状态切换稳定
            sa.query_opc(timeout=5000)   # 等待确认仪器已准备好
        except Exception as e:
            self.log(f"[粗扫] 关闭连续扫失败: {e}")
            return False, None
        sa.set_bw(rbw_hz=100.0 * 1e3)
        sa.set_freq_span(start=0.0 * 1e6,
                        stop=18000.0 * 1e6)

        sa.sweep_once(f'粗扫@{tag}')   # 再来一次完整的单扫，数据才是完整的
        time.sleep(0.5)  # 给仪器一点时间写入 trace buffer
        x, y = sa.get_trace_xy()
        # 粗扫峰值检测参数
        coarse_peak = PeakDetector(thresh_db=float(p['粗扫峰值阈值(dB)']), prom_db=float(p['粗扫邻域显著性(dB)']), guard=int(p['粗扫邻域点数']), log_func=self.log)
        peaks = coarse_peak.find(x, y)
        rbw_used = sa.last_rbw_hz if getattr(sa, 'last_rbw_hz', None) else 100.0 * 1e3

        if peaks:
            csvp, pngp, peakcsv = coarse_peak.save_csv_png(x, y, peaks, out_dir,
                                                    f'coarse_{tag}', rbw_hz=rbw_used)
            self.log(f"[粗扫] 发现异常峰 {len(peaks)} 个 -> 保存并终止")
            self.root.after(0, lambda: self.show_image_popup(pngp, title=f"粗扫异常：{tag}"))
            return True, pngp

        return False, None



    def _wait_wavelength_stable(self, p: dict, context: str = "") -> bool:
        """等待波长稳定
        返回: True 表示达到稳定，False 表示超时未稳定
        """
        same = 0
        last_val = None
        tol = 0.001
        consec_ok = 3
        max_wait = 300.0
        interval = 0.2

        t0 = time.time()
        while time.time() - t0 < max_wait:
            if getattr(self, 'pause_flag', None) and self.pause_flag.is_set():
                self._pause_point()
            if self.stop_flag.is_set():
                raise KeyboardInterrupt

            wl = self.lc.get_wavelength_nm()
            if wl is None:
                time.sleep(interval)
                continue

            if last_val is not None and abs(wl - last_val) < tol:
                same += 1
                if same >= consec_ok:
                    self.log(f"[等待稳定{context}] 波长已稳定 (连续{consec_ok}次Δ<{tol}nm)")
                    return True
            else:
                same = 0

            last_val = wl
            if same > 0:
                self.log(f"[等待稳定{context}] 当前波长 {wl:.4f} nm, 已连续稳定 {same}/{consec_ok}")
            time.sleep(interval)

        self.log(f"[等待稳定{context}] 超时未稳定")
        return False

    def _fine_scan_and_check(self, sa: 'SingleFrequency', peak: 'PeakDetector', p: dict,
                              tag: str, out_dir: str) -> tuple[bool, str | None]:
        """细扫 step 扫描；发现异常继续扫描；返回 (found, png_path)。"""
        if self.stop_flag.is_set():
            raise KeyboardInterrupt
        sa.set_sweep_type('SPD') # 速度优先
        sa.set_sweep_time(1) # 1秒扫完
        sa.set_trace_mode(max_hold=False)
        sa.set_detector("POS")   # 细扫时用峰值检波（更适合检测窄峰）
        span = float(p['细扫跨度(MHz)']) * 1e6
        step = float(p['细扫步进(MHz)']) * 1e6
        f_start = float(p['起始频率(MHz)']) * 1e6
        f_stop = float(p['终止频率(MHz)']) * 1e6
        center = f_start + span / 2.0
        
        # 先设置频宽为500MHz，再设置RBW为30kHz，避免耦合问题
        sa.set_freq_span(center=center, span=span)
        sa.set_bw(rbw_hz=float(p['细扫RBW(kHz)']) * 1e3)
        
        idx = 0
        found = False
        last_pngp = None  # 保存最后一个异常的图片路径
        while center - span / 2.0 < f_stop:
            # pause support
            if getattr(self, 'pause_flag', None) and self.pause_flag.is_set():
                self._pause_point()
            if self.stop_flag.is_set():
                raise KeyboardInterrupt
            # sa.set_freq_span(center=center, span=span)
            # sa.set_sweep_time(1) # 1秒扫完
            for repeat in range(1):
                sa.set_freq_span(center=center, span=span)
                sa.set_sweep_time(1) # 1秒扫完
                sa.sweep_once(f'细扫@{center/1e9:.3f}GHz')
                #time.sleep(2.8)
                x, y = sa.get_trace_xy()
                
                # 添加调试信息：显示当前扫描的频率范围和数据统计
                self.log(f"[细扫调试] 中心频率: {center/1e9:.3f} GHz, 跨度: {span/1e6:.0f} MHz")
                self.log(f"[细扫调试] 数据点数量: {len(x)}, 功率范围: {min(y):.2f} ~ {max(y):.2f} dBm")
                
                peaks = peak.find(x, y)
                if peaks:
                    found = True
                    tag2 = f"fine_{tag}_{int(center/1e6)}MHz"
                    rbw_used = sa.last_rbw_hz if getattr(sa, 'last_rbw_hz', None) else float(p['细扫RBW(kHz)']) * 1e3
                    csvp, pngp, peakcsv = peak.save_csv_png(x, y, peaks, out_dir, tag2, rbw_hz=rbw_used)
                    last_pngp = pngp  # 更新最后一个异常的图片路径
                    self.log(f"[细扫] 命中异常峰，保存：{tag2}.csv/.png/_peaks.csv -> 继续扫描")
                    self.root.after(0, lambda: self.show_image_popup(pngp, title=f"细扫异常：{center/1e9:.3f} GHz"))
                    break
            center += step
            idx += 1
            #time.sleep(0.5)
            if idx % 5 == 0:
                self.log(f"[细扫] 进度：center≈{center/1e9:.3f} GHz")
        if found:
            self.log('[细扫] 完成扫描，发现异常峰')
            return True, last_pngp
        else:
            self.log('[细扫] 完成扫描，未发现异常峰')
            return False, None

    def _continuous_coarse_monitor_until_stable(self, sa: 'SingleFrequency', p: dict,
                                                out_dir: str, tag: str,
                                                lc: 'LaserController', stable_params: dict) -> bool:
        """
        在等待稳定期间使用连续扫（CONT ON）监测。
        返回 True 表示在稳定前发现异常峰（已经保存并弹窗）；False 表示达到稳定（或超时未发现异常）。
        """
        mode = stable_params.get('mode', 'wavelength')
        poll = float(stable_params.get('poll', 0.1))
        max_wait = float(stable_params.get('max_wait', 300.0))
        consec_ok = int(stable_params.get('consec_ok', 1))
        tol = float(stable_params.get('tol', 0.01))
        delay_s = float(stable_params.get('delay_s', 1.5))

        # 先把粗扫参数写到仪器上（RBW / span）
        sa.set_bw(rbw_hz=100.0 * 1e3)
        sa.set_freq_span(start=0.0 * 1e6, stop=18250.0 * 1e6)
        # 用最大保持模式更容易在连续扫期间看到短时突发峰（如果你想每次只看当前扫，改为 max_hold=False）
        sa.set_trace_mode(max_hold=True)

        # 清理旧数据（再尝试）
        try:
            sa.write(sa.CMD.get('trace_clear', ':TRACe:CLEAr\n'))
        except Exception:
            pass

        t0 = time.time()
        last_check = 0.0
        same = 0
        last_val = None
        idx = 0
        found = False
        detected_peaks = set()  # 记录已检测到的峰值频率（用于去重）
        peak_tolerance = 1.0e6  # 峰值频率容差（±1MHz视为同一峰值）

        # 尝试开启连续扫描；若不支持则fallback回原先的单扫循环（避免完全失败）
        try:
            sa.sweep_continuous_on(f"{tag}_monitor")
        except Exception as e:
            self.log(f"[监测] 无法开启连续扫（回退为单次扫）：{e}")
            # fallback：重复单次扫 + 波长稳定检查（原有行为）
            while time.time() - t0 < (delay_s if mode == 'delay' else max_wait):
                # pause support
                if getattr(self, 'pause_flag', None) and self.pause_flag.is_set():
                    self._pause_point()
                if self.stop_flag.is_set():
                    raise KeyboardInterrupt
                if time.time() - last_check > poll:
                    if self._coarse_scan_and_check(sa, p, tag=f"{tag}_fallback_{idx}", out_dir=out_dir)[0]:
                        return True
                    last_check = time.time()
                    idx += 1
                if mode != 'delay':
                    v = self.lc.get_wavelength_nm()
                    if v is None:
                        time.sleep(0.5)
                        continue
                    if last_val is not None and abs(v - last_val) < tol:
                        same += 1
                        if same >= consec_ok:
                            return False
                    else:
                        same = 0
                    last_val = v
            return False

        try:
            if mode == 'delay':
                # 仅在 delay_s 时间内监测异常峰
                while time.time() - t0 < delay_s:
                    # pause support
                    if getattr(self, 'pause_flag', None) and self.pause_flag.is_set():
                        self._pause_point()
                    if self.stop_flag.is_set():
                        raise KeyboardInterrupt
                    if time.time() - last_check >= poll:
                        try:
                            x, y = sa.get_trace_xy()
                        except Exception as e:
                            self.log(f"[监测] 读取谱线失败（继续重试）：{e}")
                            time.sleep(min(0.1, poll))
                            continue
                        # 粗扫峰值检测参数（监测阶段使用粗扫参数）
                        coarse_peak = PeakDetector(thresh_db=float(p['粗扫峰值阈值(dB)']), prom_db=float(p['粗扫邻域显著性(dB)']), guard=int(p['粗扫邻域点数']), log_func=self.log)
                        peaks = coarse_peak.find(x, y)
                        if peaks:
                            # 检查是否有新的未检测过的峰值
                            new_peaks = []
                            for peak in peaks:
                                freq = peak[0]
                                # 计算峰值频率的近似值（以容差为单位）
                                approx_freq = round(freq / peak_tolerance) * peak_tolerance
                                if approx_freq not in detected_peaks:
                                    new_peaks.append(peak)
                                    detected_peaks.add(approx_freq)
                            
                            if new_peaks:
                                rbw_used = sa.last_rbw_hz if getattr(sa, 'last_rbw_hz', None) else 100.0 * 1e3
                                csvp, pngp, peakcsv = coarse_peak.save_csv_png(x, y, new_peaks, out_dir, f'cont_{tag}_{idx}', rbw_hz=rbw_used)
                                self.log(f"[监测] 稳定前发现异常峰 -> 保存 ({pngp})")
                                self.root.after(0, lambda p=pngp: self.show_image_popup(p, title=f"稳定前异常：{tag}"))
                                found = True
                            else:
                                self.log(f"[监测] 检测到已记录的峰值，跳过保存")
                        
                        last_check = time.time()
                        idx += 1
                    time.sleep(0.01)
                return found
            else:
                # 通过波长收敛次数判断稳定，同时持续检测 trace
                while time.time() - t0 < max_wait:
                    # pause support
                    if getattr(self, 'pause_flag', None) and self.pause_flag.is_set():
                        self._pause_point()
                    if self.stop_flag.is_set():
                        raise KeyboardInterrupt
                    if time.time() - last_check >= poll:
                        try:
                            x, y = sa.get_trace_xy()
                        except Exception as e:
                            self.log(f"[监测] 读取谱线失败（继续）：{e}")
                            time.sleep(min(0.1, poll))
                            continue
                        # 粗扫峰值检测参数（监测阶段使用粗扫参数）
                        coarse_peak = PeakDetector(thresh_db=float(p['粗扫峰值阈值(dB)']), prom_db=float(p['粗扫邻域显著性(dB)']), guard=int(p['粗扫邻域点数']), log_func=self.log)
                        peaks = coarse_peak.find(x, y)
                        if peaks:
                            # 检查是否有新的未检测过的峰值
                            new_peaks = []
                            for peak in peaks:
                                freq = peak[0]
                                # 计算峰值频率的近似值（以容差为单位）
                                approx_freq = round(freq / peak_tolerance) * peak_tolerance
                                if approx_freq not in detected_peaks:
                                    new_peaks.append(peak)
                                    detected_peaks.add(approx_freq)
                            
                            if new_peaks:
                                rbw_used = sa.last_rbw_hz if getattr(sa, 'last_rbw_hz', None) else 100.0 * 1e3
                                csvp, pngp, peakcsv = coarse_peak.save_csv_png(x, y, new_peaks, out_dir, f'cont_{tag}_{idx}', rbw_hz=rbw_used)
                                self.log(f"[监测] 稳定前发现异常峰 -> 保存并终止 ({pngp})")
                                self.root.after(0, lambda p=pngp: self.show_image_popup(p, title=f"稳定前异常：{tag}"))
                                found = True
                            else:
                                self.log(f"[监测] 检测到已记录的峰值，跳过保存")
                
                        last_check = time.time()
                        idx += 1

                    # 波长稳定性判据（和你原来的逻辑一致）
                    v = self.lc.get_wavelength_nm()
                    if v is None:
                        time.sleep(0.5)
                        continue
                    if last_val is not None and abs(v - last_val) < tol:
                        same += 1
                        if same >= consec_ok:
                            # 达到稳定：退出（found False）
                            return False
                    else:
                        same = 0
                    last_val = v
                    time.sleep(0.01)
                return found
        finally:
            # 不管怎样，都尝试关闭连续扫，保证状态干净
            try:
                sa.sweep_continuous_off(f"{tag}_monitor")
            except Exception as e:
                self.log(f"[监测] 关闭连续扫失败（忽略）：{e}")

    # —— 主流程 ——
    def _run(self):
        # 根据当前选择的测试类型选择参数集
        p = self.params_1um if self.test_type_var.get() == "1μm" else self.params_1_5um
        out_dir = os.path.abspath(str(p['输出目录']))
        os.makedirs(out_dir, exist_ok=True)

        sa = SingleFrequency(ip=str(p['IP地址']), timeout_s=60.0, log=self.log)

        # 上位机初始化
        lc = LaserController(exe_path=str(p['上位机路径']), window_title=str(p['窗口标题(正则)']), log_func=self.log)

        try:
            # 连接设备
            lc.start_or_connect()
            self.lc = lc
            sa.open()
            sa.set_avg(on=True, count=2)

            # 读取初始状态（仅用于记录）
            wl0 = self.lc.get_wavelength_nm()
            cur0 = self.lc.get_current_mA()
            temp0 = self.lc.get_temperature_c()
            self.log(f"[上位机] 初始状态：波长 {wl0} nm，电流 {cur0} mA，温度 {temp0} °C")

            # 获取温度相关参数
            temp_max = float(p.get("温度上限(°C)", 30.0))
            temp_min = float(p.get("温度下限(°C)", 20.0))
            temp_step = float(p.get("温度步长(°C)", 1.0))
            temp_freq = float(p.get("温度变化频率(s)", 5.0))
            
            # 获取电流相关参数
            cur_max = float(p.get("电流上限(mA)", 600.0))
            cur_min = float(p.get("电流下限(mA)", 100.0))
            cur_step = float(p.get("电流步长(mA)", 50.0))
            cur_freq = float(p.get("电流变化频率(s)", 5.0))
            
            # 初始化温度和电流
            temp = temp_min
            cur = cur_min
            
            # 设置初始温度和电流
            self.log(f"[上位机] 设置初始温度: {temp:.2f} °C")
            lc.set_temperature_c(temp)
            
            self.log(f"[上位机] 设置初始电流: {cur:.2f} mA")
            lc.set_current_mA(cur)
            
            # 初始化细扫参数
            sa.write(":TRACe1:MODE WRITe")                  # 强制 Clear Write 模式（只显示当前扫）
            sa.write(":TRACe:CLEar TRACE1")                 # 清空历史残留
            sa.set_trace_mode(max_hold=False)               # 关闭 MaxHold
            sa.set_avg(on=True, count=2)                   # 强制开2次平均
            sa.set_detector("RMS", trace=1)                 # RMS 检波 + 平均 = 超级细线
            sa.write(":BANDwidth:VIDeo:RATIO 1")            # VBW 强制 = RBW（去毛刺）
            sa.set_sweep_type('SPD')                        # 速度优先
            sa.set_sweep_time(1)                            # 1秒扫完
            
            span = 500.0 * 1e6
            step = 500.0 * 1e6
            f_start = 0.0 * 1e6
            f_stop = 18000.0 * 1e6
            center = f_start + span / 2.0
            
            # 先设置频宽为500MHz，再设置RBW为30kHz
            sa.set_freq_span(center=center, span=span)
            sa.set_bw(rbw_hz=30.0 * 1e3)
            
            # 创建共享变量和线程锁
            shared_data = {
                'temp': temp,
                'cur': cur,
                'temp_increasing': True,
                'cur_increasing': True
            }
            data_lock = threading.Lock()
            
            # 定义温度电流控制线程函数
            def temp_cur_control_thread():
                last_temp_update = time.time()
                last_cur_update = time.time()
                
                while not self.stop_flag.is_set():
                    try:
                        current_time = time.time()
                        
                        # 检查暂停状态
                        if self.pause_flag.is_set():
                            time.sleep(0.2)
                            continue
                        
                        # 检查是否需要更新温度
                        if current_time - last_temp_update >= temp_freq:
                            with data_lock:
                                if shared_data['temp_increasing']:
                                    shared_data['temp'] += temp_step
                                    if shared_data['temp'] >= temp_max:
                                        shared_data['temp'] = temp_max
                                        shared_data['temp_increasing'] = False
                                else:
                                    shared_data['temp'] -= temp_step
                                    if shared_data['temp'] <= temp_min:
                                        shared_data['temp'] = temp_min
                                        shared_data['temp_increasing'] = True
                                
                                # 获取当前温度值用于日志和设置
                                current_temp = shared_data['temp']
                            
                            # 设置新温度
                            self.log(f"[上位机] 设置温度: {current_temp:.2f} °C")
                            lc.set_temperature_c(current_temp)
                            last_temp_update = current_time
                        
                        # 检查是否需要更新电流
                        if current_time - last_cur_update >= cur_freq:
                            with data_lock:
                                if shared_data['cur_increasing']:
                                    shared_data['cur'] += cur_step
                                    if shared_data['cur'] >= cur_max:
                                        shared_data['cur'] = cur_max
                                        shared_data['cur_increasing'] = False
                                else:
                                    shared_data['cur'] -= cur_step
                                    if shared_data['cur'] <= cur_min:
                                        shared_data['cur'] = cur_min
                                        shared_data['cur_increasing'] = True
                                
                                # 获取当前电流值用于日志和设置
                                current_cur = shared_data['cur']
                            
                            # 设置新电流
                            self.log(f"[上位机] 设置电流: {current_cur:.2f} mA")
                            lc.set_current_mA(current_cur)
                            last_cur_update = current_time
                        
                        # 短暂休眠，避免CPU占用过高
                        time.sleep(0.1)
                    except Exception as e:
                        self.log(f"[错误] 温度电流控制线程出错：{e}")
                        # 继续运行，不中断线程
                        time.sleep(1.0)
            
            # 创建并启动温度电流控制线程
            temp_cur_thread = threading.Thread(target=temp_cur_control_thread, daemon=True)
            temp_cur_thread.start()
            
            # 主循环：持续进行细扫，从共享变量获取温度和电流值
            while not self.stop_flag.is_set():
                # 检查暂停状态
                if self.pause_flag.is_set():
                    self._pause_point()
                
                # 执行细扫
                sa.set_freq_span(center=center, span=span)
                
                # 每次新跨度开始前彻底清屏 + 重新开平均
                sa.write(":TRACe:CLEar TRACE1")
                sa.write(":AVERage:COUNt 2")
                sa.write(":AVERage:STATe ON")
                
                for repeat in range(2):
                    sa.set_sweep_time(1)
                    sa.sweep_once(f'细扫@{center/1e9:.3f}GHz')
                    x, y = sa.get_trace_xy()
                    
                    # 细扫峰值检测
                    fine_peak = PeakDetector(thresh_db=float(p['细扫峰值阈值(dB)']), prom_db=float(p['细扫邻域显著性(dB)']), guard=int(p['细扫邻域点数']), log_func=self.log)
                    peaks = fine_peak.find(x, y)
                    if peaks:
                        # 获取实际温度和电流值
                        actual_temp = self.lc.get_temperature_c()
                        actual_cur = self.lc.get_current_mA()
                        
                        # 处理实际值可能为None的情况，使用默认值
                        temp_str = f"{actual_temp:.3f}" if actual_temp is not None else "unknown"
                        cur_str = f"{actual_cur:.1f}" if actual_cur is not None else "unknown"
                        
                        # 保存数据，包含温度和电流信息
                        tag = f"T{temp_str}C_I{cur_str}mA"
                        tag2 = f"fine_{tag}_{int(center/1e6)}MHz"
                        rbw_used = sa.last_rbw_hz if getattr(sa, 'last_rbw_hz', None) else 30.0 * 1e3
                        csvp, pngp, peakcsv = fine_peak.save_csv_png(x, y, peaks, out_dir, tag2, rbw_hz=rbw_used)
                        self.log(f"[细扫] 命中异常峰，保存：{tag2}.csv/.png/_peaks.csv")
                        self.root.after(0, lambda p=pngp, t=actual_temp, i=actual_cur, c=center: self.show_image_popup(p, title=f"细扫异常：{c/1e9:.3f} GHz, T={t:.3f}°C, I={i:.1f}mA"))
                        break
                
                # 更新细扫中心频率
                center += step
                if center - span / 2.0 >= f_stop:
                    # 细扫完成一轮，重新开始
                    center = f_start + span / 2.0
            
            self.log("— 全部流程结束 —")

        except StopIteration as stop_reason:
            self.log(f'[终止] {stop_reason}')
            self.root.after(0, lambda: messagebox.showinfo('终止', str(stop_reason)))
        except KeyboardInterrupt:
            self.log('[停止] 用户终止。')
            self.root.after(0, lambda: messagebox.showinfo('已停止', '已根据请求停止扫描。'))
        except Exception as e:
            self.log(f'[错误] 测试失败：{e}')
            self.root.after(0, lambda err=str(e): messagebox.showerror('错误', err))
        finally:
            try:
                # 恢复初始状态
                if wl0 is not None:
                    self.log("[恢复] 正在将设备恢复到初始状态...")
                    self.lc.set_wavelength_nm(wl0)
                    if self._wait_wavelength_stable(p, " - 恢复初始波长"):
                        self.log(f"[恢复] 波长已恢复到初始值: {wl0:.6f} nm")
                    else:
                        self.log("[警告] 恢复初始波长超时")
                
                if cur0 is not None:
                    self.lc.set_current_mA(cur0)
                    time.sleep(2.0)  # 等待电流稳定
                    self.log(f"[恢复] 电流已恢复到初始值: {cur0:.2f} mA")
                
                if temp0 is not None:
                    self.lc.set_temperature_c(temp0)
                    time.sleep(2.0)  # 等待温度稳定
                    self.log(f"[恢复] 温度已恢复到初始值: {temp0:.2f} °C")
                
                self.log("[恢复] 设备已恢复到初始状态")
            except Exception as e:
                self.log(f"[警告] 恢复初始状态时出错: {e}")
            finally:
                try:
                    sa.close()
                except Exception:
                    pass

    def run(self):
        self.root.mainloop()

    def _on_test_type_change(self, event=None):
        if self.test_type_var.get() == "1μm":
            self.params = self.params_1um.copy()
        else:
            self.params = self.params_1_5um.copy()
        self._refresh_entries()

    def _refresh_entries(self):
        # 清除旧的UI元素
        for widget in self.conn_frame.winfo_children():
            widget.destroy()
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        
        self.entries.clear()
        self.param_labels.clear()
        
        # 重新构建UI
        self._build_param_ui()
    
    def _build_param_ui(self):
        """构建参数UI界面"""
        current_params = self.params_1um if self.test_type_var.get() == "1μm" else self.params_1_5um
        
        # 连接与地址参数
        conn_params = ['IP地址', '输出目录', '上位机路径', '窗口标题(正则)']
        for i, k in enumerate(conn_params):
            if k in current_params:
                label = tk.Label(self.conn_frame, text=k)
                label.grid(row=i, column=0, sticky='e')
                e = tk.Entry(self.conn_frame, width=28)
                e.insert(0, str(current_params[k]))
                e.grid(row=i, column=1, padx=4, pady=2)
                self.entries[k] = e
                self.param_labels[k] = label
        
        # 其他参数
        other_params = [k for k in current_params.keys() if k not in conn_params]
        for i, k in enumerate(other_params):
            label = tk.Label(self.param_frame, text=k)
            label.grid(row=i, column=0, sticky='e')
            e = tk.Entry(self.param_frame, width=28)
            e.insert(0, str(current_params[k]))
            e.grid(row=i, column=1, padx=4, pady=2)
            self.entries[k] = e
            self.param_labels[k] = label


if __name__ == '__main__':
    gui = SingleFrequencyGUI()
    gui.run()
# pyinstaller -F -w "d:\Coding\Project\PreciTestSystem\PTS\zhongzi\SingleFrequency.py"