import os
import time
import math
import threading
import struct
import csv
from datetime import datetime
from io import BytesIO, StringIO
import ctypes

import pyvisa
import numpy as np

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

from PIL import Image, ImageTk

import matplotlib
matplotlib.use('Agg')  # 后端绘图，不阻塞 GUI
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # <-- 补全这个

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

# -----------------------------
# Defaults - change to match your env
# -----------------------------
DEFAULT_IP = "192.168.7.10"
DEFAULT_OUTPUT_DIR = r"C:\PTS\zhongzi\Rin\Ceyear4051"
DEFAULT_POINTS = 2001
DEFAULT_SEGMENTS = [
    (10, 100, 5, 20, "File.DAT"),
    (100, 1000, 5, 20, "File_001.DAT"),
    (1000, 10000, 30, 20, "File_002.DAT"),
    (10000, 100000, 30, 20, "File_003.DAT"),
    (100000, 1000000, 30, 20, "File_004.DAT"),
    (1000000, 10000000, 30, 5, "File_005.DAT"),
]

# -----------------------------
# Helpers
# -----------------------------

def now_str():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def ensure_dir(d):
    os.makedirs(d, exist_ok=True)

# -----------------------------
# Instrument layer: Rin_4051
# -----------------------------
class Rin_4051:
    def __init__(self, ip=DEFAULT_IP, timeout_s=60.0, log_callback=None):
        self.ip = ip
        self.timeout_s = timeout_s
        self.rm = None
        self.inst = None
        self.log_callback = log_callback or (lambda s: print(s))

    def log(self, s):
        try:
            self.log_callback(s)
        except Exception:
            print(s)

    def connect(self):
        self.log(f"尝试连接: {self.ip}")

        try:
            # 每次重连都重新建 ResourceManager，避免 close() 清空的问题
            self.rm = pyvisa.ResourceManager()

            # ---- 先试 VXI-11 (带 inst0) ----
            try:
                res_str = f"TCPIP0::{self.ip}::inst0::INSTR"
                self.inst = self.rm.open_resource(res_str)
                self.inst.timeout = int(self.timeout_s * 1000)
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'
                idn = self.inst.query("*IDN?").strip()
                self.log(f"连接成功 (VXI-11 inst0): {idn}")
                return True
            except Exception as e1:
                self.log(f"VXI-11 (inst0) 失败: {e1}")
                self.close()

            # ---- 再试 VXI-11 (不带 inst0) ----
            try:
                self.rm = pyvisa.ResourceManager()
                res_str = f"TCPIP0::{self.ip}::INSTR"
                self.inst = self.rm.open_resource(res_str)
                self.inst.timeout = int(self.timeout_s * 1000)
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'
                idn = self.inst.query("*IDN?").strip()
                self.log(f"连接成功 (VXI-11): {idn}")
                return True
            except Exception as e2:
                self.log(f"VXI-11 (无 inst0) 失败: {e2}")
                self.close()

            # ---- 最后试 SOCKET ----
            try:
                self.rm = pyvisa.ResourceManager()
                res_str = f"TCPIP0::{self.ip}::5025::SOCKET"
                self.inst = self.rm.open_resource(res_str)
                self.inst.timeout = int(self.timeout_s * 1000)
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'
                idn = self.inst.query("*IDN?").strip()
                self.log(f"连接成功 (SOCKET 5025): {idn}")
                return True
            except Exception as e3:
                self.log(f"SOCKET 失败: {e3}")
                self.close()
                return False

        except Exception as e:
            self.log(f"连接失败: {e}")
            self.close()
            return False



    def close(self):
        try:
            if self.inst:
                try:
                    self.inst.close()
                except Exception:
                    pass
                self.inst = None
            if self.rm:
                try:
                    self.rm.close()
                except Exception:
                    pass
                self.rm = None
            self.log("断开连接")
        except Exception as e:
            print("关闭时异常:", e)

    def write(self, cmd):
        try:
            self.log(f"写入 => {cmd}")
            return self.inst.write(cmd)
        except Exception as e:
            self.log(f"写入失败: {e}")
            raise

    def query(self, cmd, delay=None):
        try:
            self.log(f"查询 => {cmd}")
            if delay is not None:
                time.sleep(delay)  # 如果提供了delay参数，则先等待指定时间
            return self.inst.query(cmd)
        except Exception as e:
            self.log(f"查询失败: {e}")
            raise

    def configure(self, start_hz, stop_hz, rbw_hz=1000, vbw_hz=None, points=DEFAULT_POINTS, avg_count=1):
        if self.inst is None:
            raise RuntimeError("未连接到仪器")
        try:
            # 基础配置
            self.write(f":FREQ:STARt {start_hz}")
            self.write(f":FREQ:STOP {stop_hz}")
            self.write(f":SWE:POINts {int(points)}")
            self.write(f":BAND {rbw_hz}")
            if vbw_hz is not None:
                self.write(f":BAND:VID {vbw_hz}")
            # 设置扫描类型规则为扫描速度优先
            self.write(":SWE:TYPE:AUTO:RUL SPEed")
            try:
                self.write(":UNIT:POW V")
            except Exception:
                pass
            # 平均配置
            if int(avg_count) <= 1:
                self.write(":AVER:STATe OFF")
            else:
                self.write(":AVER:STATe ON")
                self.write(f":AVER:COUNt {int(avg_count)}")
            
            # 添加轨迹配置 - 确保轨迹1处于活动状态并显示
            self.write(":DISP:WIND:TRAC:MODE WRITE")  # 设置轨迹模式为写入
            self.write(":TRAC1:MODE CLEAR WRITE")  # 清除并写入轨迹1
            self.write(":TRAC1:TYPE AVER")  # 设置轨迹1为平均类型
            
            # 设置连续/单次扫描模式
            self.write(":INIT:CONTinuous OFF")  # 设置为单次扫描模式
            
            # 显示刷新 - 确保屏幕上显示轨迹
            self.write(":DISP:UPD ON")  # 开启显示更新
            
            self.log("配置完成")
            return True
        except Exception as e:
            self.log(f"配置失败: {e}")
            return False

    # 在single_sweep_fetch方法中改进数据获取逻辑
    def single_sweep_fetch(self, prefer_binary=True):
        if self.inst is None:
            raise RuntimeError("未连接到仪器")

        # 启动单次扫描（也可以改成连续模式）
        self.write(":INIT:CONT OFF")
        self.write(":INIT")

        try:
            # 动态估算时间
            rbw = float(self.query(":BAND?"))
            points = int(self.query(":SWE:POINts?"))
            avg_count = int(self.query(":AVER:COUNt?")) if self.query(":AVER:STATe?").strip() == "1" else 1

            # 粗略估计扫描时间（经验公式，可根据实测调整）
            base_time = 0.05  # 基础时间 50ms
            rbw_factor = max(1.0, 1000.0 / rbw)
            points_factor = points / 1000.0
            est_time = base_time * rbw_factor * points_factor * avg_count

            # 给一点安全余量
            wait_time = min(est_time * 1.2, 10.0)  # 最多等 10 秒
            self.log(f"[智能等待] 估算 {wait_time:.2f}s (RBW={rbw}, 点数={points}, 平均={avg_count})")
        except Exception:
            wait_time = 1.0  # 参数获取失败时，默认等 1s

        # 主动问 OPC，直到扫完
        start = time.time()
        while time.time() - start < wait_time:
            try:
                if self.query("*OPC?").strip() == "1":
                    break
            except Exception:
                pass
            time.sleep(0.05)

        # 获取频率范围
        try:
            fstart = float(self.query(":FREQ:STAR?"))
            fstop = float(self.query(":FREQ:STOP?"))
        except Exception:
            fstart, fstop = 0.0, 1.0

        # 读取 trace 数据
        if prefer_binary:
            try:
                self.write(":FORM:DATA REAL,32")
                vals = self.inst.query_binary_values(":TRAC:DATA? TRACE1", datatype="f", is_big_endian=False)
                vals = np.array(vals, dtype=float)
                freqs = np.linspace(fstart, fstop, len(vals))
                return freqs, vals, True
            except Exception as e:
                self.log(f"二进制读取失败: {e}, 改用 ASCII")

        # ASCII 备选
        raw = self.inst.query(":TRAC:DATA? TRACE1")
        parts = [p for p in raw.replace("\n", ",").split(",") if p.strip()]
        vals = np.array([float(p) for p in parts], dtype=float)
        freqs = np.linspace(fstart, fstop, len(vals))
        return freqs, vals, False


    def _parse_scpi_block(self, raw_bytes):
        if not raw_bytes or not raw_bytes.startswith(b'#'):
            raise ValueError("不是 SCPI 二进制块")
        try:
            head_len_digit = int(chr(raw_bytes[1]))
        except Exception as e:
            raise ValueError("无法解析头长度") from e
        len_start = 2
        len_end = len_start + head_len_digit
        data_len = int(raw_bytes[len_start:len_end].decode('ascii'))
        data_start = len_end
        data_end = data_start + data_len
        data_block = raw_bytes[data_start:data_end]
        if data_len % 4 != 0:
            # still allow but warn
            self.log("警告: 数据长度不是4的倍数")
        count = data_len // 4
        vals = struct.unpack(f"<{count}f", data_block[:count*4])
        return list(vals)

    def fetch_and_save_trace(self, output_dir, base_name=None, prefer_binary=True, save_csv=True, save_dat=True):
        base_name = base_name or now_str()
        ensure_dir(output_dir)
        freqs, values, was_binary = self.single_sweep_fetch(prefer_binary=prefer_binary)
        csv_path = None
        dat_path = None
        if save_csv:
            csv_path = os.path.join(output_dir, base_name + ".csv")
            try:
                with open(csv_path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Frequency(Hz)", "Value"])
                    for fr, va in zip(freqs, values):
                        writer.writerow([f"{fr:.9f}", f"{va:.9e}"])
                self.log(f"保存 CSV: {csv_path}")
            except Exception as e:
                self.log(f"CSV 保存失败: {e}")
                csv_path = None
        if save_dat:
            dat_path = os.path.join(output_dir, base_name + ".dat")
            try:
                data_block = struct.pack(f"<{len(values)}f", *[float(v) for v in values])
                # SCPI-like block header formation (guard maximum len-of-len = 9)
                data_len_ascii = str(len(data_block)).encode('ascii')   # e.g. b'1024'
                if len(data_len_ascii) > 9:
                    raise ValueError("数据块太大，无法用标准 SCPI 单字符头表示")
                len_of_len = str(len(data_len_ascii)).encode('ascii')  # single-digit
                header = b"#" + len_of_len + data_len_ascii
                with open(dat_path, 'wb') as f:
                    f.write(header)
                    f.write(data_block)
                self.log(f"保存 DAT: {dat_path}")
            except Exception as e:
                self.log(f"DAT 保存失败: {e}")
                dat_path = None
        return csv_path, dat_path, freqs, values

# -----------------------------
# RinWorkflow - measurement + processing
# -----------------------------
class RinWorkflow:
    def __init__(self, analyzer: Rin_4051, output_dir=DEFAULT_OUTPUT_DIR, log_callback=None):
        self.analyzer = analyzer
        self.output_dir = output_dir
        self.log = log_callback or (lambda s: print(s))
        self.stop_flag = False
        self.dc_value = 1.20
        self.amplification = 14
        self.segments = DEFAULT_SEGMENTS.copy()
        self.points_expected = DEFAULT_POINTS
        # processed
        self.freqs_all = []
        self.values_all = []
        self.rin_ddx = []
        self.rin_ddy = []
        self.rin_power = []

    def request_stop(self):
        self.log("[用户] 请求停止")
        self.stop_flag = True

    def run_measurement(self, prefer_binary=True, save_csv=True, save_dat=True, progress_callback=None):
        self.stop_flag = False
        timestamp = now_str()
        session_dir = os.path.join(self.output_dir, f"RIN_{timestamp}")
        ensure_dir(session_dir)
        self.freqs_all = []
        self.values_all = []

        if not self.analyzer.inst:
            ok = self.analyzer.connect()
            if not ok:
                raise RuntimeError("无法连接到仪器")

        # 在run_measurement方法中修改循环部分
        seg_count = len(self.segments)
        for idx, seg in enumerate(self.segments):
            if self.stop_flag:
                self.log("检测到停止标志，终止测量")
                break
            start, stop, rbw, avg, fname = seg
            self.log(f"段 {idx+1}/{seg_count}: {start}Hz -> {stop}Hz, RBW={rbw}, AVG={avg}")
            
            # 根据RBW和频率范围动态调整超时时间
            base_timeout = 60.0
            # 计算频率范围跨度因子
            freq_span = stop - start
            freq_factor = min(2.0, math.sqrt(freq_span / 1000000.0))  # 频率跨度越大，因子越大，最多2倍
            dynamic_timeout = base_timeout + (rbw / 10) * freq_factor
            self.analyzer.timeout_s = dynamic_timeout
        
            if progress_callback:
                progress_callback(idx/seg_count, f"配置第{idx+1}段...")
            ok = self.analyzer.configure(start_hz=start, stop_hz=stop, rbw_hz=rbw, vbw_hz=None, points=self.points_expected, avg_count=avg)
            if not ok:
                self.log(f"段 {idx+1} 配置失败，跳过")
                continue
            
            # 为1k-10k频段(第三个频段)添加额外的等待和稳定性处理
            if idx == 2:  # 第三个频段(1k-10k)
                self.log(f"[特殊处理] 为1k-10k频段增加额外等待时间，确保RBW从5Hz平稳过渡到30Hz")
                time.sleep(5)  # 额外等待5秒
                
                # 增加额外的稳定性检查
                try:
                    # 发送额外的查询命令来确认仪器状态
                    self.analyzer.query("*IDN?")
                    self.log("[特殊处理] 仪器状态确认成功")
                except Exception as e:
                    self.log(f"[特殊处理] 状态确认异常: {e}")
                    # 重新配置这个频段
                    self.log("[特殊处理] 重新配置1k-10k频段...")
                    ok = self.analyzer.configure(start_hz=start, stop_hz=stop, rbw_hz=rbw, vbw_hz=None, points=self.points_expected, avg_count=avg)
                    if not ok:
                        self.log("[特殊处理] 重新配置失败，跳过该段")
                        continue
        
            # 为最后一段(频率范围最大的段)添加额外等待时间
            if idx == seg_count - 1:  # 最后一段
                self.log(f"[特殊处理] 为最后一段(1MHz-10MHz)增加额外等待时间，确保扫描完成")
                time.sleep(1)  # 额外等待1秒
                
                # 增加额外的稳定性检查
                try:
                    # 发送额外的查询命令来确认仪器状态
                    self.analyzer.query("*IDN?")
                    self.log("[特殊处理] 仪器状态确认成功")
                except Exception as e:
                    self.log(f"[特殊处理] 状态确认异常: {e}")
            
            if progress_callback:
                progress_callback((idx+0.2)/seg_count, f"测量第{idx+1}段...")
            try:
                base_name = f"{fname.split('.')[0]}_{timestamp}"
                csvp, datap, freqs, vals = self.analyzer.fetch_and_save_trace(session_dir, base_name=base_name, prefer_binary=prefer_binary, save_csv=save_csv, save_dat=save_dat)
            except Exception as e:
                self.log(f"段 {idx+1} 测量失败: {e}")
                continue
            # extend lists robustly
            self.freqs_all.extend(np.array(freqs, dtype=float).tolist())
            self.values_all.extend(np.array(vals, dtype=float).tolist())
            if progress_callback:
                progress_callback((idx+1)/seg_count, f"完成第{idx+1}/{seg_count}段")
        if self.stop_flag:
            self.log("测量中止，跳过处理")
            return False
        self.log("全部段完成，开始处理")
        self._process_data()
        self.log("处理完成")
        return True

    def _process_data(self):
        self.rin_ddx = []
        self.rin_ddy = []
        n = len(self.values_all)
        if n == 0:
            self.log("无数据，处理结束")
            return
        freqs = np.array(self.freqs_all, dtype=float)
        values = np.array(self.values_all, dtype=float)
        # build scale mapping similar to user's earlier logic
        scale = np.ones(n, dtype=float)
        seg_lengths = []
        pos = 0
        for i, seg in enumerate(self.segments):
            expected = self.points_expected
            actual = min(expected, max(0, n - pos))
            seg_lengths.append(actual)
            pos += actual
            if pos >= n:
                break
        rem = n - sum(seg_lengths)
        if rem > 0:
            seg_lengths.append(rem)
        idx0 = 0
        for i, L in enumerate(seg_lengths):
            idx1 = idx0 + L
            if i < 2:
                scale[idx0:idx1] = math.sqrt(5)
            else:
                scale[idx0:idx1] = math.sqrt(30)
            idx0 = idx1
        ddx = freqs.tolist()
        ddy = []
        for i, v in enumerate(values):
            try:
                if v <= 0 or not np.isfinite(v):
                    ddy.append(float('-inf'))
                else:
                    denom = (self.dc_value * self.amplification * scale[i])
                    if denom == 0:
                        ddy.append(float('-inf'))
                    else:
                        ddy.append(20.0 * math.log10(v / denom))
            except Exception:
                ddy.append(float('-inf'))
        self.rin_ddx = ddx
        self.rin_ddy = ddy
        self.rin_power = self.compute_rin_power(self.rin_ddx, self.rin_ddy)

    def compute_rin_power(self, x, y):
        power = []
        segment_length = 6
        if len(x) < 2:
            return power
        for k in range(1, len(x) // segment_length + 1):
            sub_x = x[:k*segment_length]
            sub_y = [np.power(10, val / 10.0) if np.isfinite(val) else 0.0 for val in y[:k*segment_length]]
            integral = 0.0
            for i in range(1, len(sub_x)):
                dx = sub_x[i] - sub_x[i-1]
                integral += dx * (sub_y[i] + sub_y[i-1]) / 2.0
            power.append(math.sqrt(integral))
        return power

# -----------------------------
# GUI class following the user's reference style
# -----------------------------
class Rin_4051_GUI:
    def __init__(self, parent=None):
        self.parent = parent
        
        # --- 核心修改：如果是集成模式，直接使用父控件作为 root ---
        if parent is None:
            self.root = tk.Tk()
            self.root.title("Rin_4051 - 独立模式")
            self.root.geometry("1350x370")
            self.root.resizable(True, True)
            # ... 其他窗口设置 ...
        else:
            self.root = parent # <--- 修改点：直接使用父 Frame

        # internal params (keys are used internally)
        self.params = {
            "IP_ADDRESS": DEFAULT_IP,  # 改为直接存储 IP 地址
            "OUTPUT_DIR": DEFAULT_OUTPUT_DIR,
            "DC_INPUT": 2.40,   # user input (will be divided by 2 in code)
            "AMPLIFICATION": 14,
            "POINTS": DEFAULT_POINTS,
        }

        self.param_labels = {
            "IP_ADDRESS": "IP地址",
            "OUTPUT_DIR": "输出目录",
            "DC_INPUT": "DC值",
            "AMPLIFICATION": "放大倍数",
            "POINTS": "采样点数",
        }

        self.create_widgets()

        # placeholders
        self.analyzer = None
        self.workflow = None
        self.worker_thread = None

    def create_widgets(self):
        # 主容器：左侧为参数设置，右侧为运行日志
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 左列容器：参数设置在上，按钮居中在下
        left_col = tk.Frame(main_frame)
        left_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10), pady=5)

        param_frame = tk.LabelFrame(left_col, text="参数设置", padx=8, pady=8)
        param_frame.pack(fill=tk.Y)

        self.entries = {}
        row = 0
        for key, val in self.params.items():
            label_text = self.param_labels.get(key, key)
            tk.Label(param_frame, text=label_text, anchor="e").grid(row=row, column=0, sticky="e", padx=4, pady=4)
            ent = tk.Entry(param_frame, width=30)
            ent.insert(0, str(val))
            ent.grid(row=row, column=1, columnspan=2, padx=4, pady=4)
            self.entries[key] = ent
            row += 1

        # 按钮区域：放在参数设置框的下面并居中（保持按钮横向排列不变）
        buttons_frame = tk.Frame(left_col)
        buttons_frame.pack(pady=8)

        tk.Button(buttons_frame, text="开始测试", command=self.start_test, bg="#4CAF50", fg="#FFFFFF", width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(buttons_frame, text="停止测试", command=self.stop_test, bg="#f44336", fg="#FFFFFF", width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(buttons_frame, text="底噪测试", command=self.measure_background, bg="#f4a236", fg="#FFFFFF", width=12).pack(side=tk.LEFT, padx=6)

        # 运行日志 (右侧)
        log_frame = tk.LabelFrame(main_frame, text="运行日志", padx=6, pady=6)
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=5)
        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def log(self, msg):
        t = time.strftime("[%H:%M:%S]")
        try:
            self.log_box.insert(tk.END, f"{t} {msg}\n")
            self.log_box.see(tk.END)
            try:
                self.root.update_idletasks()
            except Exception:
                pass
            print(f"{t} {msg}")
        except Exception:
            print(f"{t} {msg}")

    def update_params(self):
        for k, widget in self.entries.items():
            v = widget.get().strip()
            # try to convert numeric fields
            if k in ("DC_INPUT", "AMPLIFICATION", "POINTS"):
                try:
                    if '.' in v:
                        self.params[k] = float(v)
                    else:
                        self.params[k] = int(v)
                except Exception:
                    self.params[k] = v
            else:
                self.params[k] = v
        self.log("[设置] 参数已保存")

    def start_test(self):
        # ensure params saved
        self.update_params()
        ip_address = self.params.get("IP_ADDRESS", DEFAULT_IP)
        outdir = self.params.get("OUTPUT_DIR") or DEFAULT_OUTPUT_DIR
        ensure_dir(outdir)
        try:
            dc_input = float(self.params.get("DC_INPUT") or 2.40)
        except Exception:
            dc_input = 2.40
        try:
            amp = float(self.params.get("AMPLIFICATION") or 14)
        except Exception:
            amp = 14
        try:
            points = int(self.params.get("POINTS") or DEFAULT_POINTS)
        except Exception:
            points = DEFAULT_POINTS

        segments = DEFAULT_SEGMENTS

        # create analyzer and workflow
        self.log(f"准备连接到频谱仪: {ip_address}")
        self.analyzer = Rin_4051(ip=ip_address, log_callback=self.log)
        self.workflow = RinWorkflow(self.analyzer, output_dir=outdir, log_callback=self.log)
        self.workflow.dc_value = float(dc_input) / 2.0  # per previous program behavior
        self.workflow.amplification = amp
        self.workflow.points_expected = points
        self.workflow.segments = segments

        # guard: do not start multiple
        if self.worker_thread and getattr(self.worker_thread, "is_alive", lambda: False)():
            self.log("[警告] 已有测量在运行")
            return

        # start worker thread
        self.log("[主控] 启动测量线程...")
        self.worker_thread = threading.Thread(target=self._worker_measure, daemon=True)
        self.worker_thread.start()

    def _worker_measure(self):
        try:
            ok = self.analyzer.connect()
            if not ok:
                self.log("[错误] 连接失败，测量终止")
                return
            def progress_cb(frac, msg):
                # update log and optionally a visual progress (we only log here)
                self.log(f"[进度] {int(frac*100)}% - {msg}")
            success = self.workflow.run_measurement(prefer_binary=True, save_csv=True, save_dat=True, progress_callback=progress_cb)
            if success:
                self.log("[主控] 测量完成，准备生成结果图...")
                png_path = self.visualize_data()
                if png_path:
                    try:
                        self.root.after(0, lambda p=png_path: self.show_image_popup(p))
                    except Exception:
                        self.show_image_popup(png_path)
            else:
                self.log("[主控] 测量未完成（可能被中止）")
        except Exception as e:
            self.log(f"[错误] 测量时异常: {e}")
        finally:
            try:
                self.analyzer.close()
            except Exception:
                pass
            self.log("[主控] 线程结束")

    def stop_test(self):
        if self.workflow:
            self.workflow.request_stop()
            self.log("[主控] 已请求停止 - 正在等待线程响应")

    def measure_background(self):
        """测量底噪并弹窗显示曲线"""
        self.update_params()
        outdir = self.params.get("OUTPUT_DIR") or DEFAULT_OUTPUT_DIR
        ensure_dir(outdir)
        ip_to_use = self.params.get("IP_ADDRESS") or DEFAULT_IP

        # 增加超时时间从60秒到120秒
        analyzer = Rin_4051(ip=ip_to_use, timeout_s=120.0, log_callback=self.log)
        ok = analyzer.connect()
        if not ok:
            messagebox.showerror("错误", "连接频谱仪失败", parent=self.root)
            return

        try:
            # 设置扫描参数
            analyzer.write(":SWE:POINts 2001")
            analyzer.write(":UNIT:POW V")
            analyzer.write(":INIT:CONT OFF")
            analyzer.write(":FREQ:STAR 10")
            analyzer.write(":FREQ:STOP 100000000")
            analyzer.write(":BAND 30")
            analyzer.write(":AVER:STAT ON")
            
            # 只执行一次初始化和等待
            analyzer.write(":INIT")
            
            # 增强的超时处理 - 增加等待时间并添加重试逻辑
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries:
                try:
                    self.log(f"[底噪] 等待操作完成 (尝试 {retry_count+1}/{max_retries})...")
                    analyzer.query("*OPC?")
                    break  # 成功获取响应，退出循环
                except Exception as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        self.log(f"[底噪] *OPC? 查询失败，将等待较长时间: {e}")
                        time.sleep(10)  # 最后的尝试：等待更长时间
                    else:
                        self.log(f"[底噪] *OPC? 查询超时，等待2秒后重试: {e}")
                        time.sleep(2)  # 等待2秒后重试

            # 修改：直接读取数据，不再调用 single_sweep_fetch（避免重复初始化）
            self.log("尝试二进制读取 TRACE (REAL,32)...")
            analyzer.write(":FORM:DATA REAL,32")
            vals = analyzer.inst.query_binary_values(":TRAC:DATA? TRACE1", datatype='f', is_big_endian=False)
            vals = np.array(vals, dtype=float)
            
            # 获取频率信息
            fstart = float(analyzer.query(":FREQ:STAR?"))
            fstop = float(analyzer.query(":FREQ:STOP?"))
            pts = len(vals)
            freqs = np.linspace(fstart, fstop, pts)
            
            # 修改：先进行数据预处理，再绘图
            # 去除可能存在的极端异常值
            valid_indices = np.isfinite(vals) & (vals > np.min(vals) * 10)  # 简单过滤
            freqs = freqs[valid_indices]
            vals = vals[valid_indices]
            
            if not freqs.size or not vals.size:
                raise RuntimeError("未能获取有效的曲线数据")

            # 弹窗显示曲线
            win = tk.Toplevel(self.root)
            win.title("底噪曲线")
            win.geometry("800x600")

            fig, ax = plt.subplots(figsize=(7, 4))
            ax.plot(freqs, vals, linewidth=1)
            ax.set_xscale("log")
            ax.set_title("Noise Floor", fontsize=14, fontweight='bold')
            ax.set_xlabel("Frequency (Hz)", fontsize=12)
            ax.set_ylabel("Power (V)", fontsize=12)
            
            # 添加网格线样式区分
            ax.grid(which='major', linestyle='-', linewidth='0.7', color='gray')
            ax.grid(which='minor', linestyle=':', linewidth='0.5', color='gray')

            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # 保存按钮
            def save_curve():
                fn = filedialog.asksaveasfilename(
                    initialdir=outdir,
                    defaultextension=".png",
                    filetypes=[("PNG 文件", "*.png")]
                )
                if fn:
                    fig.savefig(fn, dpi=300, bbox_inches="tight")
                    messagebox.showinfo("提示", f"曲线已保存到: {fn}", parent=win)

            tk.Button(win, text="保存曲线", command=save_curve).pack(pady=8)

            # 居中弹窗
            win.update_idletasks()
            x = (win.winfo_screenwidth() - win.winfo_width()) // 2
            y = (win.winfo_screenheight() - win.winfo_height()) // 2
            win.geometry(f"+{x}+{y}")

        except Exception as e:
            self.log(f"[底噪] 操作失败: {e}")
            messagebox.showerror("错误", f"底噪测量失败: {e}", parent=self.root)
        finally:
            analyzer.close()

    def visualize_data(self):
        """可视化 RIN 测试结果"""
        if not self.workflow or not self.workflow.rin_ddx or not self.workflow.rin_ddy:
            print("没有可视化的数据")
            return

        ddx = self.workflow.rin_ddx
        ddy = self.workflow.rin_ddy
        rin_power = self.workflow.rin_power

        root = tk.Toplevel(self.root)
        root.title("测RIN数据可视化")

        fig, (ax1, ax2) = plt.subplots(
            2, 1, figsize=(10, 8),
            gridspec_kw={'height_ratios': [3, 2]}
        )

        # -------- 保存按钮 --------
        def save_figure():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"),
                        ("JPEG files", "*.jpg"),
                        ("All files", "*.*")]
            )
            if file_path:
                fig.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("成功", f"图像已保存到 {file_path}")

        top_frame = tk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=10)
        tk.Button(
            top_frame, text="保存", command=save_figure,
            font=('SimHei', 20)
        ).pack(side=tk.TOP)

        # -------- 图1: RIN 曲线 --------
        ax1.plot(ddx, ddy, color="#085cab", linewidth=2)
        ax1.set_xscale('log')
        ax1.margins(x=0)
        ax1.set_ylabel('RIN (dBc/Hz)', fontsize=18, fontweight='bold')
        ax1.grid(True, which='both', lw=2, linestyle='--', alpha=1)
        ax1.set_xlim(10, 10**7)

        # 边框加粗
        for spine in ax1.spines.values():
            spine.set_linewidth(2.5)
        # 坐标刻度字体
        for label in ax1.get_xticklabels() + ax1.get_yticklabels():
            label.set_fontname('Times New Roman')
            label.set_fontsize(20)
            label.set_fontweight('bold')

        finite_ddy = [v for v in ddy if np.isfinite(v)]
        if finite_ddy:
            ax1.set_ylim(
                np.floor(min(finite_ddy)/10)*10,
                np.ceil(max(finite_ddy)/10)*10
            )
        ax1.yaxis.set_major_locator(MaxNLocator(nbins=6, integer=True))

        # -------- 图2: RMS 积分曲线 --------
        adjusted_power = [p * 100 for p in rin_power]
        ax2.plot(ddx[::6], adjusted_power, color="#085cab", linewidth=2)
        ax2.set_xscale('log')
        ax2.margins(x=0)
        ax2.set_xlim(10, 10**7)

        y2_min, y2_max = np.min(adjusted_power), np.max(adjusted_power)
        if np.isclose(y2_min, y2_max):
            y2_min, y2_max = y2_min - 1, y2_max + 1
        y2_mid = (y2_min + y2_max) / 2
        ax2.set_yticks([y2_min, y2_mid, y2_max])
        ax2.set_yticklabels(
            [f"{y2_min:.3f}%", f"{y2_mid:.3f}%", f"{y2_max:.3f}%"]
        )

        ax2.set_xlabel('Frequency(Hz)', fontsize=18, fontweight='bold')
        ax2.set_ylabel('Integrated RMS', fontsize=18, fontweight='bold')
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

        # -------- 提取指定点的 RIN 值 --------
        target_xs = [1000, 10000, 100000, 1000000]
        result_text = ""
        for tx in target_xs:
            idx = np.argmin(np.abs(np.array(ddx) - tx))
            x_val = ddx[idx]
            y_val = ddy[idx]
            result_text += f"x={x_val:.0f} 时, y={y_val:.3f} dBc/Hz\n"

        messagebox.showinfo("指定点的RIN值", result_text, parent=root)



    def show_image_popup(self, img_path, save_button_top_center=False):
        try:
            win = tk.Toplevel(self.root)
            win.title("图像预览")
            img = Image.open(img_path)
            maxw, maxh = 900, 700
            w, h = img.size
            if w > maxw or h > maxh:
                ratio = min(maxw / w, maxh / h)
                img = img.resize((int(w*ratio), int(h*ratio)))
            img_tk = ImageTk.PhotoImage(img)

            lbl = tk.Label(win, image=img_tk)
            lbl.image = img_tk
            lbl.pack(padx=10, pady=10)

            def save_img():
                tgt = filedialog.asksaveasfilename(defaultextension=".png",
                                                filetypes=[("PNG files","*.png"),
                                                            ("JPEG files","*.jpg")],
                                                parent=win)
                if tgt:
                    try:
                        img.save(tgt)
                        messagebox.showinfo("保存", f"已保存到: {tgt}", parent=win)
                    except Exception as e:
                        messagebox.showerror("保存失败", f"保存图片时出错: {e}", parent=win)

            btn_frame = tk.Frame(win)
            btn_frame.pack(pady=6)

            if save_button_top_center:
                # 保存按钮居中，去掉关闭按钮
                tk.Button(btn_frame, text="保存图片", command=save_img).pack(side=tk.TOP, pady=4)
            else:
                # 默认布局：保存 + 关闭
                tk.Button(btn_frame, text="保存图片", command=save_img).pack(side=tk.LEFT, padx=6)
                tk.Button(btn_frame, text="关闭", command=win.destroy).pack(side=tk.LEFT, padx=6)

        except Exception as e:
            self.log(f"[弹窗] 打开图片失败: {e}")
            messagebox.showerror("错误", f"打开图片失败: {e}", parent=self.root)


    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    gui = Rin_4051_GUI()
    gui.run()

# pyinstaller --onefile --noconsole --icon="D:\pack2\PreciLasers.ico" --hidden-import=pyvisa --clean "D:\pack2\RinTester.py"