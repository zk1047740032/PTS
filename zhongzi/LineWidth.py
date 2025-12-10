import pyvisa
import time
import tkinter as tk
from tkinter import messagebox, filedialog
import os
from PIL import Image, ImageTk
import threading
import shutil
import ctypes

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
class SignalGenerator:
    def __init__(self, log_callback=None) -> None:
        self.rm = None
        self.inst = None
        self.log = log_callback or (lambda msg: None)

    def connect(self, ip_address):
        self.rm = pyvisa.ResourceManager()
        self.inst = self.rm.open_resource(f'TCPIP0::{ip_address}::inst0::INSTR')
        self.inst.timeout = 10000
        self.inst.read_termination = '\n'
        self.inst.write_termination = '\n'
        self.log(f"[信号源] 已连接到信号发生器")

    def configure(self, waveform="SIN", freq=0.1, volt=0, offset=1):
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
        if self.inst:
            try:
                self.inst.close()
                self.log("[信号源] 已关闭信号发生器连接")
            except Exception:
                pass
        if self.rm:
            try:
                self.rm.close()
            except Exception:
                pass

# ============ 仪器控制类 ============
class LinewidthTester:
    def __init__(self, log_callback=None) -> None:
        self.rm = pyvisa.ResourceManager()
        self.inst = None
        self.log = log_callback or (lambda msg: None)
        self.stop_flag = threading.Event()

    def connect(self, ip_address):
        self.inst = self.rm.open_resource(f'TCPIP0::{ip_address}::inst0::INSTR')
        self.inst.timeout = 10000
        self.log("已连接到频谱仪")

    def configure(self, center_freq, span, rbw, n_db_down):
        self.inst.write("INIT:CONT OFF")  # 关闭连续扫描
        # 添加单位：中心频率使用MHZ，带宽使用MHZ，RBW使用HZ
        self.inst.write(f"FREQ:CENT {center_freq}MHZ")
        self.inst.write(f"FREQ:SPAN {span}KHZ")
        self.inst.write(f"BAND {rbw}HZ")
        self.inst.write("SWE:POIN 2001")  # 设置扫描点数
        self.inst.write(":AVER:COUN 20")
        self.log("设置Count数为20")
        self.n_db_down = n_db_down
        self.log("完成参数设置")

    def measure(self):
        if self.stop_flag.is_set():
            return False
        self.inst.write("INIT;*WAI")  # 开始测量并等待完成
        if self.stop_flag.is_set():
            return False
        self.inst.write("CALC:MARK1 ON")  # 启用 Marker1
        self.inst.write("CALC:MARK:FUNC:NDBD:STAT ON")  # 打开NdBdown
        self.inst.write(f"CALC:MARK1:FUNC:NDBD {self.n_db_down}")  # 设置 N 的值
        self.inst.write("CALC:MARK:MAX:AUTO ON")  # 移动 Marker 到最大峰值
        #self.inst.write("CALC:MARK1:FUNC:EXEC")  # 执行功能计算
        self.log("测量完成")
        return True

    def save_data(self, instr_image_path, instr_trace_csv, pc_shared_folder):
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
            
            return pc_image_path

        except Exception as e:
            self.log(f"保存数据失败: {e}")
            raise

    def close(self):
        if self.inst:
            self.inst.close()
            self.log("连接已关闭")

    def stop(self):
        self.stop_flag.set()
        self.log("测量已停止")

# ============ GUI 控制类 ============
class LineWidthGUI:
    def __init__(self, parent=None):
        self.parent = parent
        
        # --- 核心修改：如果是集成模式，直接使用父控件作为 root ---
        if parent is None:
            self.root = tk.Tk()
            self.root.title("线宽 - 独立模式")
            self.root.geometry("1100x470") 
            self.root.resizable(True, True)
        else:
            self.root = parent # <--- 修改点：直接使用父 Frame
        
        # 初始化参数
        self.params = {
            # 仪器参数（只保存数值，不保存单位）
            '中心频率(MHZ)': '80',
            'RBW(HZ)': '100',
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
        sw = window.winfo_screenwidth()
        sh = window.winfo_screenheight()
        x = (sw - width) // 2
        y = (sh - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")
    
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
        test_params = ['中心频率(MHZ)', 'RBW(HZ)', 'N dB down']
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
        
        self.start_btn = tk.Button(inner_btn_frame, text='开始测试', bg="#4CAF50", fg="#FFFFFF", command=self.start_measurement)
        self.start_btn.pack(side='left', padx=6)
        
        self.stop_btn = tk.Button(inner_btn_frame, text='停止测试', bg="#f44336", fg="#FFFFFF", command=self.stop_measurement, state=tk.DISABLED)
        self.stop_btn.pack(side='left', padx=6)
        
        # 运行日志区 - 放在右侧
        logf = tk.LabelFrame(right_frame, text='运行日志', padx=6, pady=6)
        logf.pack(fill=tk.BOTH, expand=True)
        self.log_box = tk.Text(logf, font=('Arial', 10))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        
    def log(self, msg):
        t = time.strftime('[%H:%M:%S]')
        self.root.after(0, lambda: self._safe_log_append(f"{t} {msg}\n"))
    
    def _safe_log_append(self, text):
        self.log_box.insert(tk.END, text)
        self.log_box.see(tk.END)
    
    def _save_params(self):
        """保存当前输入的参数"""
        for k, e in self.entries.items():
            v = e.get()
            self.params[k] = v
        self.log('[参数] 已更新')
    
    def start_measurement(self):
        """开始测量"""
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
                
                # 保存所有测试结果图片路径和对应的Span值
                all_results = []
                
                for span in span_values:
                    if self.tester.stop_flag.is_set():
                        self.log(f"[停止] 已停止测试，当前完成到Span: {span}")
                        break
                        
                    self.log(f"\n[Span测试] 开始测试Span: {span}")
                    
                    # 配置参数，使用当前Span值
                    self.tester.configure(
                        center_freq=self.params['中心频率(MHZ)'],
                        span=span,
                        rbw=self.params['RBW(HZ)'],
                        n_db_down=self.params['N dB down']
                    )
                    
                    # 执行测量
                    if not self.tester.measure():
                        self.log(f"[Span测试] 测量失败，跳过Span: {span}")
                        continue
                    
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
                    
                    # 保存数据
                    image_path = self.tester.save_data(
                        instr_image_path=instr_image_path,
                        instr_trace_csv=instr_trace_csv,
                        pc_shared_folder=self.params['输出目录']
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
                    signal_gen.connect(self.params['信号发生器IP'])
                    
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
                        center_freq=self.params['中心频率(MHZ)'],
                        span=span,
                        rbw=self.params['RBW(HZ)'],
                        n_db_down=self.params['N dB down']
                    )
                    
                    # 执行测量
                    if self.tester.measure():
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
                        
                        # 保存数据
                        image_path = self.tester.save_data(
                            instr_image_path=instr_image_path,
                            instr_trace_csv=instr_trace_csv,
                            pc_shared_folder=self.params['输出目录']
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
                    
                    # 关闭信号发生器连接
                    signal_gen.close()
                    
                except Exception as e:
                    self.log(f"[错误] 信号发生器控制或额外测试失败：{e}")
                
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
        """停止测量"""
        if self.tester:
            self.tester.stop()
        self.stop_flag.set()
        self.log("[停止] 用户请求停止测量")
    
    def show_results_selection(self, all_results):
        """显示测试结果选择界面，让用户选择需要查看的截图"""
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
        save_btn = tk.Button(btn_frame, text="保存选中图片", font=('Arial', 12), bg="#4CAF50", fg="white", command=save_selected_image)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # 关闭按钮
        close_btn = tk.Button(btn_frame, text="关闭", font=('Arial', 12), bg="#f44336", fg="white", command=win.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        # 初始显示第一张图片
        show_selected_image()
    
    def show_image_popup(self, image_path, span_value=None):
        """显示测量结果截图，单张显示，支持手动保存"""
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
            win.destroy()
        
        # 添加保存按钮
        tk.Button(btn_frame, text="保存图片", font=('Arial', 12), command=_save_img).pack(side=tk.LEFT, padx=10)
        
        # 添加关闭按钮
        tk.Button(btn_frame, text="关闭", font=('Arial', 12), command=_close_window).pack(side=tk.RIGHT, padx=10)
        
        # 显示图片的标签
        img_label = tk.Label(win, image=win.img_tk)
        img_label.pack(padx=6, pady=6, fill=tk.BOTH, expand=True)
        
        # 绑定窗口关闭事件
        win.protocol("WM_DELETE_WINDOW", _close_window)
    
    def run(self):
        """运行GUI"""
        self.root.mainloop()

# ============ 程序入口 ============
if __name__ == '__main__':
    gui = LineWidthGUI()
    gui.run()