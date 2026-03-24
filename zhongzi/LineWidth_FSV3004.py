import pyvisa
import time
import tkinter as tk
from tkinter import messagebox, filedialog
import os
from PIL import Image, ImageTk, ImageDraw, ImageFont
import threading
import shutil
import ctypes
import csv
import math

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

# ============ 信号发生器控制类 ============
# 本类用于通过 LAN 接口远程控制信号发生器（如 Keysight/Agilent 33500B 系列），
# 实现以下功能：
# 1. 建立/关闭 TCPIP 连接；
# 2. 配置输出波形类型、频率、幅值与直流偏移；
# 3. 打开/关闭输出；
# 4. 提供线程安全的日志回调，方便 GUI 实时显示状态。
# 对外主要方法：
# - connect(ip)          : 建立连接；
# - configure(...)       : 一次性写入所有波形参数；
# - set_output(on=True)  : 控制输出开关；
# - close()              : 释放 VISA 资源。
class SignalGenerator:
    def __init__(self, log_callback=None) -> None:
        """
        初始化信号发生器控制类

        参数:
            log_callback (callable, optional): 日志回调函数，用于输出日志信息
        """
        self.rm = None # 存储Visa资源管理器
        self.inst = None # 存储仪器实例
        self.log = log_callback or (lambda msg: None)

    def connect(self, ip_address, max_retries=3, retry_interval=2):
        """
        连接信号发生器（添加重试机制）

        参数:
            ip_address (str): 信号发生器的 IP 地址
            max_retries (int, optional): 最大重试次数，默认3次
            retry_interval (int, optional): 重试间隔时间（秒），默认2秒
        返回:
            bool: 连接成功返回 True，失败返回 False
        """
        for attempt in range(max_retries+1):
            try:
                self.rm = pyvisa.ResourceManager()
                self.inst = self.rm.open_resource(f'TCPIP0::{ip_address}::INSTR')
                self.inst.timeout = 10000 # 设置通信超时时间，防止仪器无响应时程序无限等待
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'
                # 验证连接是否可用
                self.inst.query('*IDN?')
                self.log(f"[信号源] 已连接到信号发生器（第{attempt+1}次尝试）")
                return True
            except Exception as e:
                if attempt < max_retries:
                    self.log(f"[信号源] 第{attempt+1}次连接失败：{e}，{retry_interval}秒后重试...")
                    time.sleep(retry_interval)
                else:
                    self.log(f"[信号源] 连接失败，共尝试{max_retries}次，请尝试手动断开重连")
                    return False
        return False

    def configure(self, waveform="SIN", freq=0.1, volt=0, offset=1):
        """
        配置信号发生器输出参数

        参数:
            waveform (str, optional): 波形类型，如"SIN"表示正弦波。默认"SIN"。
            freq (float, optional): 输出频率，单位Hz。默认0.1 Hz。
            volt (float, optional): 输出幅值，单位Vpp。默认0 Vpp。
            offset (float, optional): 直流偏移，单位Vdc。默认1 Vdc。

        返回:
            bool: 配置成功返回True，失败返回False。
        """
        if not self.inst:
            self.log("[信号源] 未连接到信号发生器")
            return False
        try:
            # 设置波形类型（SIN=正弦波）
            self.inst.write(f":SOUR1:FUNC {waveform}")
            self.log(f"[信号源] 设置波形: {waveform}")
            
            # 设置频率（单位：Hz）
            self.inst.write(f":SOUR1:FREQ {freq}Hz")
            self.log(f"[信号源] 设置频率: {freq} Hz")
            
            # 设置幅值（单位：Vpp）
            self.inst.write(f":SOUR1:VOLT {volt}")
            self.log(f"[信号源] 设置幅值: {volt} Vpp")
            
            # 设置偏移（单位：Vdc）
            self.inst.write(f":SOUR1:VOLT:OFFS {offset}")
            self.log(f"[信号源] 设置偏移: {offset} Vdc")
            
            return True
        except Exception as e:
            self.log(f"[信号源] 配置失败: {e}")
            return False

    def set_output(self, on=True):
        '''
        设置信号发生器输出开关状态

        参数:
            on (bool, optional): 是否打开输出。True 表示打开，False 表示关闭。默认 True。

        返回:
            bool: 设置成功返回 True，失败返回 False。
        '''
        if not self.inst:
            self.log("[信号源] 未连接到信号发生器")
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

    def close(self):
        """
        关闭信号发生器连接与资源管理器，释放VISA资源。
        """
        if self.inst:
            try:
                self.inst.close()
                self.log("[信号源] 已关闭信号发生器连接")
                self.inst = None
            except Exception:
                pass
        if self.rm:
            try:
                self.rm.close()
                self.inst = None
            except Exception:
                pass

# ============ 频谱仪控制类 ============
# 该类用于通过 LAN 接口远程控制 R&S FSV3004 频谱仪，完成线宽测试全流程：
# 1. 连接与初始化；
# 2. 根据给定参数配置频谱仪（参考电平、中心频率、Span、RBW、NdB down 等）；
# 3. 触发单次扫描并放置 Marker1；
# 4. 将截图与 Trace 数据保存到仪器本地指定目录；
# 5. 通过 VISA 的 MMEM:COPY 指令把文件转存到电脑共享文件夹；
# 6. 自动生成同名的 .dat 文件（复制 .csv 改扩展名）；
# 7. 支持线程安全停止（stop_flag 事件）；
# 8. 提供关闭连接与资源释放接口。
# 对外主要方法：
# - connect(ip)               : 建立 TCPIP 连接；
# - configure(...)            : 一次性写入所有测试参数；
# - measure()                 : 执行单次扫描并计算 NdB down；
# - save_data(...)            : 完成截图、Trace 保存及文件拷贝；
# - close() / stop()          : 资源释放与中止控制。
class LinewidthTester:
    def __init__(self, log_callback=None) -> None:
        """
        初始化线宽测试仪控制类

        参数:
            log_callback (callable, optional): 日志回调函数，用于输出日志信息
        """
        self.rm = pyvisa.ResourceManager()
        self.inst = None
        self.log = log_callback or (lambda msg: None)
        self.stop_flag = threading.Event()

    def connect(self, ip_address):
        """
        建立与频谱仪的 TCPIP 连接。

        参数:
            ip_address (str): 频谱仪的 IP 地址。
        """
        self.inst = self.rm.open_resource(f'TCPIP0::{ip_address}::inst0::INSTR')
        self.inst.timeout = 10000
        self.log("已连接到频谱仪")

    def configure(self, ref_level, M1_position, center_freq, span, rbw, n_db_down):
        """
        配置频谱仪测量参数。

        参数:
            ref_level (float): 参考电平，单位为 mV。
            M1_position (float): Marker1 所在频率位置，单位为 MHz。
            center_freq (float): 中心频率，单位为 MHz。
            span (float): 扫频宽度，单位为 kHz。
            rbw (float): 分辨率带宽，单位为 Hz。
            n_db_down (float): 用于线宽测量的 N dB 下降值。
        """
        self.inst.write("INIT:CONT OFF")  # 关闭连续扫描
        # 添加单位：中心频率使用MHZ，带宽使用MHZ，RBW使用HZ
        self.inst.write(f"DISP:TRAC:Y:RLEV {ref_level}mV")  # 设置参考电平
        self.inst.write(f"CALC:MARKer1:X {M1_position}MHz") # 设置M1位置
        # 设置中心频率，单位 MHz
        self.inst.write(f"FREQ:CENT {center_freq}MHZ")
        # 设置扫频宽度（Span），单位 kHz
        self.inst.write(f"FREQ:SPAN {span}KHZ")
        # 设置分辨率带宽（RBW），单位 Hz
        self.inst.write(f"BAND {rbw}HZ")
        # 将功率单位设置为伏特（V），便于后续以电压方式读取测量结果
        self.inst.write(f"CALC:UNIT:POW V")
        self.inst.write("SWE:POIN 2001")  # 设置扫描点数
        self.inst.write(":AVER:COUN 20") # 设置单位为V
        #self.log("设置Count数为20")
        self.n_db_down = n_db_down
        self.log("完成参数设置")

    def measure(self):
        """
        执行单次扫描，启用 Marker1 并设置 NdB down 功能，用于后续线宽计算。

        流程：
        1. 检查停止标志，若已置位则直接返回 False；
        2. 触发单次扫描并等待完成；
        3. 打开 Marker1 及 NdB down 功能，写入预设的 N 值；
        4. 记录日志并返回 True 表示测量成功。

        返回:
            bool: 成功完成测量返回 True；若中途被停止或异常返回 False。
        """
        if self.stop_flag.is_set():
            return False
        self.inst.write("INIT;*WAI")  # 开始测量并等待完成
        if self.stop_flag.is_set():
            return False
        self.inst.write("CALC:MARK1 ON")  # 启用 Marker1
        self.inst.write("CALC:MARK:FUNC:NDBD:STAT ON")  # 打开NdBdown
        self.inst.write(f"CALC:MARK1:FUNC:NDBD {self.n_db_down}")  # 设置 N 的值
        #self.inst.write("CALC:MARK:MAX:AUTO ON")  # 移动 Marker 到最大峰值
        #self.inst.write("CALC:MARK1:FUNC:EXEC")  # 执行功能计算
        self.log("测量完成")
        return True

    def get_ndbdown_result(self):
        """
        读取 NdBdown 的测量结果值。

        返回:
            float: NdBdown 的测量结果值；若失败返回 None。
        """
        try:
            result = self.inst.query("CALC:MARK:FUNC:NDBD:RES?").strip()
            time.sleep(1.5)
            ndbdown_value = float(result)
            self.log(f"[频谱仪] NdBdown值: {ndbdown_value/1000:.2f}kHz")
            return ndbdown_value
        except Exception as e:
            self.log(f"[频谱仪] 读取NdBdown值失败: {e}")
            return None

    def save_ndbdown_to_csv(self, ndbdown_value, csv_file):
        """
        将 NdBdown 值保存到 CSV 文件。

        参数:
            ndbdown_value (float): 要保存的 NdBdown 值。
            csv_file (str): CSV 文件的路径。
        """
        try:
            file_exists = os.path.exists(csv_file)
            # 计算ndbdown除以2倍根号99后的值
            sqrt_99 = math.sqrt(99)
            calculated_value = ndbdown_value / (2 * sqrt_99)
            with open(csv_file, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                if not file_exists:
                    writer.writerow(['NdBdown'])
                writer.writerow([f"{calculated_value/1000:.2f}"])
        except Exception as e:
            self.log(f"[频谱仪] 保存NdBdown失败: {e}")

    def save_data(self, instr_image_path, instr_trace_csv, pc_shared_folder, ndbdown_value=None):
        """
        将频谱仪截图与Trace数据保存到仪器本地，并拷贝到电脑共享文件夹，同时生成同名.dat文件。
        如果提供了ndbdown_value，则在截图上添加计算公式注释。

        参数:
            instr_image_path (str): 期望保存截图的仪器端完整路径（仅用于提取文件名）。
            instr_trace_csv (str): 期望保存Trace数据的仪器端完整路径（仅用于提取文件名）。
            pc_shared_folder (str): 电脑端共享文件夹路径，用于接收拷贝文件。
            ndbdown_value (float, optional): NdBdown 的测量值，用于在图片上添加计算公式注释。

        返回:
            str: 电脑端截图文件的完整路径；若失败则返回None（异常会被抛出）。

        异常:
            Exception: 任意步骤失败时抛出，供上层捕获并记录日志。
        """
        if self.stop_flag.is_set():
            return False
        try:
            # 确保仪器本地路径使用C:\PTS\LineWidth目录
            # 提取文件名
            image_filename = os.path.basename(instr_image_path)
            csv_filename = os.path.basename(instr_trace_csv)
            
            # 构建仪器本地完整路径
            instrument_image_path = f"C:\\PTS\\zhongzi\\LineWidth\\{image_filename}"
            instrument_csv_path = f"C:\\PTS\\zhongzi\\LineWidth\\{csv_filename}"
            
            # 1. 保存截图到仪器本地路径
            self.inst.write("HCOPy:DEST 'MMEM'")
            self.inst.write(f"MMEM:NAME '{instrument_image_path}'")
            self.inst.write("HCOPy:IMM")
            self.inst.query("*OPC?")
            self.log(f"截图已保存到仪器内部: {instrument_image_path}")

            # 2. 保存Trace数据到仪器本地路径
            self.inst.write(f":MMEM:STOR:TRAC 1, '{instrument_csv_path}'")
            self.inst.query("*OPC?")
            self.log(f"Trace数据已保存到仪器内部: {instrument_csv_path}")

            # 3. 将文件从仪器复制到电脑共享文件夹，使用与仪器本地路径相同的文件名
            dat_filename = os.path.splitext(csv_filename)[0] + '.dat'
            
            # 构建电脑共享文件夹中的完整路径
            pc_image_path = os.path.join(pc_shared_folder, image_filename)
            pc_trace_csv = os.path.join(pc_shared_folder, csv_filename)
            pc_trace_dat = os.path.join(pc_shared_folder, dat_filename)
            
            # 复制文件
            self.inst.write(f"MMEM:COPY '{instrument_image_path}', '{pc_image_path}'")
            self.inst.query("*OPC?")
            self.log(f"截图已复制到电脑共享文件夹: {image_filename}")

            self.inst.write(f"MMEM:COPY '{instrument_csv_path}', '{pc_trace_csv}'")
            self.inst.query("*OPC?")
            self.log(f"Trace数据已复制到电脑共享文件夹: {csv_filename}")

            # 4. 生成dat文件，复制csv改扩展名
            if os.path.exists(pc_trace_csv):
                shutil.copyfile(pc_trace_csv, pc_trace_dat)
                self.log(f"已生成同目录的dat 文件: {dat_filename}")
            
            # 5. 如果提供了 ndbdown_value，在截图上添加计算公式注释
            if ndbdown_value is not None and os.path.exists(pc_image_path):
                try:
                    self._add_ndbdown_text_to_image(pc_image_path, ndbdown_value)
                except Exception as e:
                    self.log(f"[警告] 添加NdBdown文字到图片失败: {e}")
            
            return pc_image_path

        except Exception as e:
            self.log(f"保存数据失败: {e}")
            raise

    def _add_ndbdown_text_to_image(self, image_path, ndbdown_value):
        """
        在图片上添加 NdBdown 计算公式的文字注释。
        公式：NdBdown / (2×√99) ≈ 计算结果

        参数:
            image_path (str): 图片的文件路径。
            ndbdown_value (float): NdBdown 的测量值。
        """
        try:
            # 打开图片
            img = Image.open(image_path)
            draw = ImageDraw.Draw(img)
            
            # 计算结果：ndbdown_value / (2 * sqrt(99))
            sqrt_99 = math.sqrt(99)
            result = ndbdown_value / (2 * sqrt_99) / 1000
            
            # 创建要显示的文本
            text = f"{ndbdown_value/1000:.2f}kHz÷(2√99)≈{result:.2f}kHz"
            
            # 尝试加载字体，优先使用 Arial，否则使用默认字体
            try:
                font = ImageFont.truetype("arial.ttf", 32)
            except:
                try:
                    font = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 32)
                except:
                    font = ImageFont.load_default()
            
            # 在图片左上角添加文字（位置可调整）
            # 使用白色文字，黑色背景增加可读性
            text_position = (800, 400)
            
            # 添加黑色背景框
            # bbox = draw.textbbox(text_position, text, font=font)
            # draw.rectangle([bbox[0]-5, bbox[1]-5, bbox[2]+5, bbox[3]+5], fill="black")
            
            # 绘制白色文字
            draw.text(text_position, text, font=font, fill="black")
            
            # 保存修改后的图片
            img.save(image_path)
            self.log(f"[频谱仪] 已在图片上添加NdBdown计算公式: {text}")
            
        except Exception as e:
            self.log(f"[错误] 处理图片失败: {e}")
            raise

    def close(self):
        """
        关闭频谱仪连接与资源管理器，释放Visa资源。
        """
        if self.inst:
            try:
                self.inst.close()
                self.log(f"已关闭频谱仪连接")
                self.inst = None
            except Exception:
                pass
        if self.rm:
            try:
                self.rm.close()
                self.rm = None
            except Exception:
                pass

    def stop(self):
        """
        停止当前测量线程。
        
        通过设置 stop_flag 事件通知测量线程立即终止后续操作，
        并记录停止日志。
        """
        self.stop_flag.set()
        self.log("测量已停止")

# ============ 线宽测试 GUI 控制类 ============
# 该类负责构建并管理线宽测试的图形化界面，支持以下功能：
# 1. 参数输入与持久化（频谱仪/信号源 IP、测试参数等）；
# 2. 一键启动/停止多 Span 自动化测试流程；
# 3. 实时日志输出；
# 4. 测试结果（截图 + Trace 数据）集中预览、手动保存；
# 5. 支持独立窗口（tk.Tk）或嵌入外部 Frame（tk.Frame）两种运行模式；
# 6. 线程安全，后台测试不阻塞 GUI。
# 对外暴露：
# - 实例化即可自动构建界面；
# - 若传入 parent=Frame，则作为子控件嵌入；否则生成独立窗口。
class LineWidth_FSV3004_GUI:
    def __init__(self, parent=None):
        """
        初始化线宽测试 GUI 控制类
        参数:
            parent (tk.Widget, optional): 父控件。若为 None，则创建独立窗口；否则作为子控件嵌入 parent。
        """
        self.parent = parent
        # --- 核心修改：如果是集成模式，直接使用父控件作为 root ---
        if parent is None:
            self.root = tk.Tk()
            self.root.title("线宽 - 独立模式")
            self.root.geometry("1150x550") 
            self.root.resizable(True, True)
        else:
            self.root = parent # <--- 修改点：直接使用父 Frame
        
        # 初始化参数
        self.params = {
            # 仪器参数（只保存数值，不保存单位）
            'Ref Level(mV)': '1000',
            'M1位置(MHz)': '80',
            '中心频率(MHz)': '80',
            'RBW(Hz)': '100',
            'N dB down': '20',
            
            # 连接与地址
            '频谱仪IP': '192.168.7.10',
            '信号发生器IP': '192.168.7.11',
            '仪器本地图片路径': r"C:\PTS\zhongzi\LineWidth\image.png",
            '仪器本地数据路径': r"C:\PTS\zhongzi\LineWidth\data.csv",
            '输出目录': r"\\192.168.7.7\PTS\zhongzi\LineWidth"
        }
        
        self.worker = None
        self.tester = None
        self.stop_flag = threading.Event()
        
        # 构建UI
        self._build_ui()
    
    def set_center(self, window, width, height):
        """
        将指定窗口在屏幕上居中显示。

        参数:
            window (tk.Toplevel | tk.Tk): 需要居中的窗口对象。
            width (int): 窗口宽度（像素）。
            height (int): 窗口高度（像素）。
        """
        sw = window.winfo_screenwidth()
        sh = window.winfo_screenheight()
        x = (sw - width) // 2
        y = (sh - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    def _build_ui(self):
        """
        构建并布局整个 GUI 界面。

        功能：
        1. 创建左右分栏的主框架；
        2. 左侧依次放置“连接与地址”“参数设置”两个标签框架，以及居中显示的“开始/停止”按钮区；
        3. 右侧放置“运行日志”标签框架，内含可滚动文本框；
        4. 所有输入框统一宽度与标签对齐方式，保证界面整洁；
        5. 将输入控件实例存入 self.entries，供后续读取与保存参数使用。
        """
        # 创建主框架，分为左右两部分
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧框架 - 参数设置
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))
        
        # 右侧框架 - 运行日志
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 连接与地址设置区 - 放在左侧上方
        conn_frame = tk.LabelFrame(left_frame, text='连接与地址', padx=8, pady=8)
        conn_frame.pack(fill=tk.X, pady=6)
        
        # 参数设置区 - 放在左侧下方
        param_frame = tk.LabelFrame(left_frame, text='参数设置', padx=8, pady=8)
        param_frame.pack(fill=tk.X, pady=6)
        
        self.entries = {}
        
        # 连接与地址参数
        conn_params = ['频谱仪IP', '信号发生器IP', '输出目录']
        for i, k in enumerate(conn_params):
            # 显示时使用更友好的标签名
            display_name = k
            # 统一标签宽度和对齐方式
            label = tk.Label(conn_frame, text=display_name, width=12, anchor='e')
            label.grid(row=i, column=0, sticky='e', padx=5, pady=2)
            # 统一输入框宽度
            e = tk.Entry(conn_frame, width=20)
            e.insert(0, str(self.params[k]))
            e.grid(row=i, column=1, padx=4, pady=2)
            self.entries[k] = e
        
        # 参数设置
        test_params = ['Ref Level(mV)', 'M1位置(MHz)', '中心频率(MHz)', 'RBW(Hz)', 'N dB down']
        for i, k in enumerate(test_params):
            # 统一标签宽度和对齐方式
            label = tk.Label(param_frame, text=k, width=12, anchor='e')
            label.grid(row=i, column=0, sticky='e', padx=5, pady=2)
            # 统一输入框宽度
            e = tk.Entry(param_frame, width=20)
            e.insert(0, str(self.params[k]))
            e.grid(row=i, column=1, padx=4, pady=2)
            self.entries[k] = e
        
        # 按钮区域放在参数设置框下方，居中显示
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=8)
        
        # 创建一个内部框架来容纳按钮，实现居中
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')
        
        self.start_btn = tk.Button(inner_btn_frame, text='开始测试', bg="#4CAF50", fg="#FFFFFF", command=self.start_measurement, cursor="hand2")
        self.start_btn.pack(side='left', padx=6)
        
        self.stop_btn = tk.Button(inner_btn_frame, text='停止测试', bg="#f44336", fg="#FFFFFF", command=self.stop_measurement, state=tk.DISABLED, cursor="hand2")
        self.stop_btn.pack(side='left', padx=6)
        
        # 运行日志区 - 放在右侧
        logf = tk.LabelFrame(right_frame, text='运行日志', padx=6, pady=6)
        logf.pack(fill=tk.BOTH, expand=True)
        self.log_box = tk.Text(logf, font=('Arial', 10))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        
    def log(self, msg):
        """
        将日志信息输出到 GUI 的日志框中，并自动滚动到最新内容。

        参数:
            msg (str): 要显示的日志文本。
        """
        t = time.strftime('[%H:%M:%S]')
        self.root.after(0, lambda: self._safe_log_append(f"{t} {msg}\n"))
    
    def _safe_log_append(self, text):
        """
        线程安全地向日志文本框追加内容，并自动滚动到最新一行。

        参数:
            text (str): 需要追加的日志文本。
        """
        self.log_box.insert(tk.END, text)
        self.log_box.see(tk.END)
    
    def _save_params(self):
        """
        将当前界面中所有输入框的值保存到 self.params 字典中，
        并在日志中提示参数已更新。
        """
        for k, e in self.entries.items():
            v = e.get()
            self.params[k] = v
        self.log('[参数] 已更新')
    
    def start_measurement(self):
        """
        启动线宽测试流程。

        功能：
        1. 检查是否已有测试线程在运行，防止重复启动；
        2. 保存当前界面参数到 self.params；
        3. 禁用“开始”按钮、启用“停止”按钮；
        4. 清空电脑共享文件夹与仪器内部文件夹；
        5. 按预设的 Span 列表依次完成多次线宽测量；
        6. 控制信号发生器输出 0.1 Hz、0 Vpp、1 Vdc 的正弦波，并在 Span=500 kHz 下追加一次额外测试；
        7. 收集所有截图路径，测试结束后弹出结果选择窗口供用户预览与手动保存；
        8. 无论成功或异常，最终恢复按钮状态并关闭仪器连接。
        """
        if self.worker and self.worker.is_alive():
            messagebox.showinfo('提示', '测试已在进行中')
            return
        
        # 保存参数
        self._save_params()
        
        # 更新按钮状态
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        self.stop_flag.clear()
        
        def task():
            """
            后台线程任务：执行完整的线宽自动化测试流程。

            步骤概览：
            1. 清空电脑共享文件夹与仪器内部文件夹；
            2. 按预设 Span 列表（100/200/500/1000/2000 kHz）依次配置频谱仪并测量；
            3. 控制信号发生器输出 0.1 Hz、0 Vpp、1 Vdc 正弦波，追加 Span=500 kHz 的额外测试；
            4. 收集所有截图路径，测试结束后弹出结果选择窗口供用户预览与手动保存；
            5. 无论成功或异常，最终恢复按钮状态并关闭仪器连接。

            异常处理：
            - 任意步骤失败均记录日志并弹出错误提示；
            - 用户点击“停止”后立即终止后续操作。
            """
            try:
                self.log("[开始] 线宽测试开始")
                
                # 清空仪器和电脑共享文件夹
                self.log("[初始化] 正在清空共享文件夹和仪器内部文件夹...")
                
                # 1. 清空电脑共享文件夹
                local_dir = self.params['输出目录']
                if os.path.exists(local_dir):
                    try:
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
                        self.log(f"[初始化] 已清空电脑共享文件夹: {local_dir}")
                    except Exception as e:
                        self.log(f"[错误] 清空电脑共享文件夹失败: {e}")
                
                # 2. 清空仪器内部文件夹
                try:
                    # 连接仪器以清空文件夹
                    temp_rm = pyvisa.ResourceManager()
                    temp_inst = temp_rm.open_resource(f'TCPIP0::{self.params["频谱仪IP"]}::inst0::INSTR')
                    temp_inst.timeout = 10000
                    
                    # 创建目录（如果不存在）
                    temp_inst.write("MMEM:MDIR 'C:\\PTS\\zhongzi\\LineWidth'")
                    # 清空目录
                    temp_inst.write("MMEM:DEL 'C:\\PTS\\zhongzi\\LineWidth\\*.*'")
                    temp_inst.close()
                    temp_rm.close()
                    self.log("[初始化] 已清空仪器内部文件夹: C:\\PTS\\zhongzi\\LineWidth")
                except Exception as e:
                    self.log(f"[警告] 清空仪器文件夹失败: {e}")
                
                self.log("[初始化] 文件夹清理完成。")
                
                # 创建测试实例
                self.tester = LinewidthTester(log_callback=self.log)
                
                # 连接仪器
                self.tester.connect(self.params['频谱仪IP'])
                
                # 定义要测试的四个Span值
                span_values = ['100', '200', '500', '1000', '2000']
                
                # 准备 ndbdown.csv 文件路径
                ndbdown_csv_path = os.path.join(self.params['输出目录'], 'ndbdown.csv')
                
                # 保存所有测试结果图片路径和对应的Span值
                all_results = []
                
                for span in span_values:
                    if self.tester.stop_flag.is_set():
                        self.log(f"[停止] 已停止测试，当前完成到Span: {span}")
                        break
                        
                    self.log(f"\n[Span测试] 开始测试Span: {span}")
                    
                    # 配置参数，使用当前Span值
                    self.tester.configure(
                        ref_level = self.params['Ref Level(mV)'],
                        M1_position= self.params['M1位置(MHz)'],
                        center_freq=self.params['中心频率(MHz)'],
                        span=span,
                        rbw=self.params['RBW(Hz)'],
                        n_db_down=self.params['N dB down']
                    )
                    
                    # 执行测量
                    if not self.tester.measure():
                        self.log(f"[Span测试] 测量失败，跳过Span: {span}")
                        continue
                    
                    # 读取 NdBdown 测量结果值
                    ndbdown_value = self.tester.get_ndbdown_result()
                    
                    # 如果成功读取 NdBdown 值，保存到 CSV
                    if ndbdown_value is not None:
                        self.tester.save_ndbdown_to_csv(ndbdown_value, ndbdown_csv_path)
                    
                    # 为不同Span值生成唯一文件名
                    base_name = os.path.splitext(os.path.basename(self.params['仪器本地图片路径']))[0]
                    span_suffix = span.replace('KHZ', 'K').replace('MHZ', 'M')
                    
                    # 构建文件路径
                    instr_image_path = os.path.join(
                        os.path.dirname(self.params['仪器本地图片路径']),
                        f"{base_name}_{span_suffix}.png"
                    )
                    
                    instr_trace_csv = os.path.join(
                        os.path.dirname(self.params['仪器本地数据路径']),
                        f"{base_name}_{span_suffix}.csv"
                    )
                    
                    # 保存数据（传递 ndbdown_value 参数）
                    image_path = self.tester.save_data(
                        instr_image_path=instr_image_path,
                        instr_trace_csv=instr_trace_csv,
                        pc_shared_folder=self.params['输出目录'],
                        ndbdown_value=ndbdown_value
                    )
                    
                    if image_path and os.path.exists(image_path):
                        # 保存结果信息
                        all_results.append({
                            'image_path': image_path,
                            'span_value': span,
                            'file_name': os.path.basename(image_path)
                        })
                        self.log(f"[Span测试] Span: {span} 测试完成，结果已保存")
                    else:
                        self.log(f"[Span测试] Span: {span} 未找到截图文件")
                
                self.log(f"\n[完成] 线宽测试结束，共完成 {len(all_results)} 个Span测试")
                
                # ============ 信号发生器控制与额外测试 ============
                self.log("\n[信号源] 开始配置信号发生器")
                
                try:
                    # 创建并连接信号发生器
                    signal_gen = SignalGenerator(log_callback=self.log)
                    connected = signal_gen.connect(self.params['信号发生器IP'])
                    
                    # 只有在成功连接信号发生器时才执行额外测试
                    if connected:
                        # 配置信号发生器：正弦波、频率0.1Hz、幅值0vpp、偏移1vdc
                        signal_gen.configure(waveform="SIN", freq=0.1, volt=0, offset=1)
                        
                        # 打开信号发生器输出
                        signal_gen.set_output(on=True)
                        
                        # 等待信号稳定
                        time.sleep(1)
                        
                        # ============ 额外线宽测试（Span=200kHz） ============
                        self.log("\n[额外测试] 开始Span=200kHz的线宽测试")
                        
                        # 配置频谱仪Span=500kHz
                        span = '500'
                        self.log(f"[额外测试] 开始测试Span: {span}")
                        
                        # 配置参数，使用500kHz Span
                        self.tester.configure(
                            ref_level = self.params['Ref Level(mV)'],
                            M1_position= self.params['M1位置(MHz)'],
                            center_freq=self.params['中心频率(MHz)'],
                            span=span,
                            rbw=self.params['RBW(Hz)'],
                            n_db_down=self.params['N dB down']
                        )
                        
                        # 执行测量
                        if self.tester.measure():
                            # 读取 NdBdown 测量结果值
                            ndbdown_value = self.tester.get_ndbdown_result()
                            
                            # 如果成功读取 NdBdown 值，保存到 CSV
                            if ndbdown_value is not None:
                                self.tester.save_ndbdown_to_csv(ndbdown_value, ndbdown_csv_path)
                            
                            # 为额外测试生成唯一文件名
                            base_name = os.path.splitext(os.path.basename(self.params['仪器本地图片路径']))[0]
                            span_suffix = '500+1v'  # 200kHz = 200K
                            
                            # 构建文件路径
                            instr_image_path = os.path.join(
                                os.path.dirname(self.params['仪器本地图片路径']),
                                f"{base_name}_{span_suffix}_with_signal.png"
                            )
                            
                            instr_trace_csv = os.path.join(
                                os.path.dirname(self.params['仪器本地数据路径']),
                                f"{base_name}_{span_suffix}_with_signal.csv"
                            )
                            
                            # 保存数据（传递 ndbdown_value 参数）
                            image_path = self.tester.save_data(
                                instr_image_path=instr_image_path,
                                instr_trace_csv=instr_trace_csv,
                                pc_shared_folder=self.params['输出目录'],
                                ndbdown_value=ndbdown_value
                            )
                            
                            if image_path and os.path.exists(image_path):
                                # 保存结果信息
                                all_results.append({
                                    'image_path': image_path,
                                    'span_value': span,
                                    'file_name': os.path.basename(image_path)
                                })
                                self.log(f"[额外测试] Span: {span} 测试完成，结果已保存")
                            else:
                                self.log(f"[额外测试] Span: {span} 未找到截图文件")
                        
                        # 关闭信号发生器输出
                        signal_gen.set_output(on=False)
                    else:
                        self.log("[信号源] 未成功连接信号发生器，跳过额外测试")
                        
                except Exception as e:
                    self.log(f"[错误] 信号发生器控制或额外测试失败：{e}")
                finally:
                    # 关闭信号发生器连接
                    if 'signal_gen' in locals():
                        signal_gen.close()
                
                # ============ 显示所有测试结果 ============
                self.log(f"\n[最终结果] 共完成 {len(all_results)} 个Span测试")
                
                # 测试全部完成后，显示结果选择界面
                if all_results:
                    self.root.after(0, lambda results=all_results: self.show_results_selection(results))
                else:
                    self.root.after(0, lambda: messagebox.showinfo("完成", "未完成任何Span测试！"))
                
            except Exception as e:
                self.log(f"[错误] 测试失败：{e}")
                self.root.after(0, lambda err=str(e): messagebox.showerror('错误', err))
            finally:
                # 关闭连接
                if self.tester:
                    self.tester.close()
                # 恢复按钮状态
                self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
        
        self.worker = threading.Thread(target=task, daemon=True)
        self.worker.start()
    
    def stop_measurement(self):
        """
        立即停止当前正在进行的测量任务。

        通过调用测试实例的 stop() 方法以及设置本类的 stop_flag 事件，
        通知后台线程立即终止后续测量步骤，并在日志中记录停止原因。
        """
        if self.tester:
            self.tester.stop()
        self.stop_flag.set()
        self.log("[停止] 用户请求停止测量")
    
    def show_results_selection(self, all_results):
        """
        弹出测试结果选择窗口，供用户预览并手动保存所有测试截图。

        参数:
            all_results (list[dict]): 包含所有测试结果的列表，每个元素为字典，需包含：
                - 'image_path' (str): 截图文件的完整路径
                - 'span_value' (str): 对应的 Span 值
                - 'file_name' (str): 截图文件名
        """
        win = tk.Toplevel(self.root)
        win.title("测试结果选择")
        win.transient(self.root)
        win.resizable(True, True)
        
        # 设置弹窗大小和居中
        self.set_center(win, 2100, 1300)
        
        # 创建主框架
        main_frame = tk.Frame(win)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧：结果列表
        left_frame = tk.Frame(main_frame, width=200)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # 右侧：图片预览
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 结果列表标题
        list_title = tk.Label(left_frame, text="测试结果列表", font=('Arial', 14, 'bold'))
        list_title.pack(pady=10)
        
        # 结果列表框
        listbox = tk.Listbox(left_frame, font=('Arial', 12), selectmode=tk.SINGLE)
        listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 填充结果列表
        for i, result in enumerate(all_results):
            listbox.insert(tk.END, f"Span: {result['span_value']} - {result['file_name']}")
        
        # 默认选中第一个结果
        if all_results:
            listbox.select_set(0)
        
        # 右侧：图片显示区域
        img_frame = tk.LabelFrame(right_frame, text="图片预览", padx=10, pady=10)
        img_frame.pack(fill=tk.BOTH, expand=True)
        
        # 图片标签
        img_label = tk.Label(img_frame)
        img_label.pack(fill=tk.BOTH, expand=True)
        
        # 底部按钮区域
        btn_frame = tk.Frame(right_frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        # 显示当前选中的图片
        def show_selected_image(event=None):
            """
            在右侧图片预览区显示当前选中的测试结果截图。

            参数:
                event (tk.Event, optional): Listbox 选择事件对象，可省略。
            """
            selected_index = listbox.curselection()
            if not selected_index:
                return
            
            index = selected_index[0]
            result = all_results[index]
            
            # 加载并显示图片
            pil_img = Image.open(result['image_path'])
            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            max_w, max_h = int(sw * 0.6), int(sh * 0.6)
            
            # 调整图片大小
            scale = min(max_w / pil_img.width, max_h / pil_img.height)
            new_size = (int(pil_img.width * scale), int(pil_img.height * scale))
            disp_img = pil_img.resize(new_size, Image.LANCZOS)
            
            img_tk = ImageTk.PhotoImage(disp_img)
            img_label.config(image=img_tk)
            img_label.image = img_tk
            
            # 保存当前图片信息
            img_label.current_img = pil_img
            img_label.current_result = result
        
        # 保存当前选中的图片
        def save_selected_image():
            """
            保存当前在图片预览区显示的图片到用户指定路径。
            功能：
            1. 检查是否已加载图片；
            2. 弹出文件保存对话框，默认格式为 PNG；
            3. 将图片保存到用户选择的路径，并给出成功或失败提示。

            """
            if not hasattr(img_label, 'current_img'):
                messagebox.showwarning("提示", "请先选择要保存的图片")
                return
            
            save_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                     filetypes=[("PNG 文件", "*.png"), ("所有文件", "*.*")],
                                                     title="保存图片")
            if save_path:
                try:
                    img_label.current_img.save(save_path)
                    messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")
                except Exception as ex:
                    messagebox.showerror("保存失败", str(ex))
        
        # 列表框选择事件
        listbox.bind('<<ListboxSelect>>', show_selected_image)
        
        # 保存按钮
        save_btn = tk.Button(btn_frame, text="保存选中图片", font=('Arial', 12), bg="#4CAF50", fg="white", command=save_selected_image, cursor="hand2")
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # 关闭按钮
        close_btn = tk.Button(btn_frame, text="关闭", font=('Arial', 12), bg="#f44336", fg="white", command=win.destroy, cursor="hand2")
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        # 初始显示第一张图片
        show_selected_image()
    
    def show_image_popup(self, image_path, span_value=None):
        """
        弹出独立窗口，用于预览并手动保存单张测试截图。

        参数:
            image_path (str): 待显示图片的完整文件路径。
            span_value (str, optional): 当前截图对应的 Span 值，用于窗口标题提示；可省略。
        """
        win = tk.Toplevel(self.root)
        
        # 设置窗口标题，显示当前Span值
        if span_value:
            win.title(f"测量结果预览 - Span: {span_value}")
        else:
            win.title("测量结果预览")
        
        win.transient(self.root)
        win.resizable(True, True)
        
        # 设置弹窗大小和居中
        self.set_center(win, 1800, 1600)
        
        # 加载并显示图片
        pil_img = Image.open(image_path)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        max_w, max_h = int(sw * 0.7), int(sh * 0.7)
        
        # 调整图片大小
        scale = min(max_w / pil_img.width, max_h / pil_img.height)
        new_size = (int(pil_img.width * scale), int(pil_img.height * scale))
        disp_img = pil_img.resize(new_size, Image.LANCZOS)
        
        img_tk = ImageTk.PhotoImage(disp_img)
        win.orig_img = pil_img
        win.img_tk = img_tk
        
        # 创建按钮框架
        btn_frame = tk.Frame(win)
        btn_frame.pack(side=tk.TOP, fill='x', pady=8)
        
        # 保存图片按钮
        def _save_img():
            """
            弹出文件保存对话框，将当前窗口显示的图片保存到用户指定路径。
            
            功能：
            1. 弹出保存对话框，默认格式为 PNG；
            2. 若用户确认保存，则将图片写入指定路径；
            3. 保存成功或失败均给出对应提示。
            """
            save_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                     filetypes=[("PNG 文件", "*.png"), ("所有文件", "*.*")],
                                                     title="保存图片")
            if save_path:
                try:
                    win.orig_img.save(save_path)
                    messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")
                except Exception as ex:
                    messagebox.showerror("保存失败", str(ex))
        
        # 关闭窗口按钮
        def _close_window():
            """
            关闭当前弹出的图片预览窗口。
            """
            win.destroy()
        
        # 添加保存按钮
        tk.Button(btn_frame, text="保存图片", font=('Arial', 12), command=_save_img, cursor="hand2").pack(side=tk.LEFT, padx=10)
        
        # 添加关闭按钮
        tk.Button(btn_frame, text="关闭", font=('Arial', 12), command=_close_window, cursor="hand2").pack(side=tk.RIGHT, padx=10)
        
        # 显示图片的标签
        img_label = tk.Label(win, image=win.img_tk)
        img_label.pack(padx=6, pady=6, fill=tk.BOTH, expand=True)
        
        # 绑定窗口关闭事件
        win.protocol("WM_DELETE_WINDOW", _close_window)
    
    def run(self):
        """
        启动 GUI 主事件循环，使窗口进入可交互状态并等待用户操作。
        """
        self.root.mainloop()

# ============ 程序入口 ============
if __name__ == '__main__':
    gui = LineWidth_FSV3004_GUI()
    gui.run()