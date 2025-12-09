import os
import tkinter as tk
from tkinter import ttk
import multiprocessing
import threading
import sys
import time
from queue import Empty
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

# 为了支持多进程，我们需要动态导入模块
# 这样每个进程都会有自己的导入副本
def import_Rin_FSV3004():
    from zhongzi.Rin_FSV3004 import RinGUI
    return RinGUI

def import_Rin_4051():
    from zhongzi.Rin_4051 import Rin_4051_GUI
    return Rin_4051_GUI

def import_LineWidth():
    from zhongzi.LineWidth import LineWidthGUI
    return LineWidthGUI

def import_TimeDomain():
    from zhongzi.TimeDomain import TimeDomainGUI
    return TimeDomainGUI

def import_SpectrumSNR():
    from zhongzi.SpectrumSNR import SpectrumSNRGUI
    return SpectrumSNRGUI

def import_SingleFrequency():
    from zhongzi.SingleFrequency import SingleFrequencyGUI
    return SingleFrequencyGUI

def import_CT_W():
    from qijian.CT_W import CT_W_GUI
    return CT_W_GUI

def import_CT_P():
    from qijian.CT_P import CT_P_GUI
    return CT_P_GUI

def import_CT_L():
    from qijian.CT_L import CT_L_GUI
    return CT_L_GUI

# 进程入口函数，用于在新进程中启动GUI
# 添加了通信队列参数，用于进程间通信
def start_gui_process(gui_importer, queue=None, process_name="", *args, **kwargs):
    try:
        # 如果有队列，发送进程启动消息
        if queue is not None:
            queue.put({"type": "status", "process": process_name, "status": "started", "timestamp": time.time()})
        
        # 导入GUI类
        gui_class = gui_importer()
        # 创建GUI实例，不传递parent参数，让GUI类自己创建主窗口
        gui = gui_class(None, *args, **kwargs)
        # 运行主循环
        gui.root.mainloop()
        
        # 进程正常结束
        if queue is not None:
            queue.put({"type": "status", "process": process_name, "status": "completed", "timestamp": time.time()})
    except Exception as e:
        # 发送错误消息
        if queue is not None:
            queue.put({"type": "error", "process": process_name, "error": str(e), "timestamp": time.time()})
        raise

class MainApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("PTS - 一体化测试系统")
        # self.root.iconbitmap(r"D:\Coding\Project\DataAutomation\集成测试平台\PreciLasers.ico")
        self.root.iconbitmap(os.path.join(os.path.dirname(__file__), "PreciLasers.ico"))
        self.root.geometry("570x770")  # 增大窗口以容纳状态监控
        self.root.resizable(True, True)
        
        # 创建进程通信队列
        self.queue = multiprocessing.Queue()
        # 存储所有子进程
        self.processes = {}
        # 创建状态监控区域
        self.create_status_monitor()
        # 启动状态监控线程
        self.start_status_monitor_thread()

        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 添加标题
        title_label = ttk.Label(main_frame, text="测试模块选择", font=("SimHei", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 创建状态监控区域
        self.status_frame = ttk.LabelFrame(main_frame, text="进程状态监控", padding="10")
        self.status_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 创建状态列表
        self.status_list = tk.Listbox(self.status_frame, height=5, width=80)
        self.status_list.pack(fill=tk.X, padx=5, pady=5)

        # 创建Notebook（页签）
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # 创建“zhongzi”页签
        seed_frame = ttk.Frame(notebook)
        notebook.add(seed_frame, text="种子")

        # 创建“qijian”页签
        device_frame = ttk.Frame(notebook)
        notebook.add(device_frame, text="器件")

        """zhongzi部门项目按钮"""
        rin_fsv3004_btn = ttk.Button(
            seed_frame, 
            text="Rin_FSV3004", 
            command=self.open_Rin, 
            width=20
        )
        rin_fsv3004_btn.pack(pady=10)

        rin_4051_btn = ttk.Button(
            seed_frame,
            text="Rin_思仪4051",
            command=self.open_Rin_4051,
            width=20,
        )
        rin_4051_btn.pack(pady=10)

        linewidth_btn = ttk.Button(
            seed_frame, 
            text="线宽", 
            command=self.open_LineWidth, 
            width=20
        )
        linewidth_btn.pack(pady=10)

        timedomain_btn = ttk.Button(
            seed_frame, 
            text="时域", 
            command=self.open_TimeDomain, 
            width=20
        )
        timedomain_btn.pack(pady=10)

        spectrumsnr_btn = ttk.Button(
            seed_frame, 
            text="信噪比", 
            command=self.open_SpectrumSNR, 
            width=20
        )
        spectrumsnr_btn.pack(pady=10)

        singlefrequency_btn = ttk.Button(
            seed_frame, 
            text="单频", 
            command=self.open_SingleFrequency, 
            width=20
        )
        singlefrequency_btn.pack(pady=10)

        """qijian部门项目按钮"""
        ct_w_btn = ttk.Button(
            device_frame, 
            text="电流温度_波长", 
            command=self.open_CT_W, 
            width=20
        )
        ct_w_btn.pack(pady=10)

        ct_p_btn = ttk.Button(
            device_frame,
            text="电流温度_功率",
            command=self.open_CT_P,
            width=20
        )
        ct_p_btn.pack(pady=10)

        ct_l_btn = ttk.Button(
            device_frame,
            text="电流温度_线宽",
            command=self.open_CT_L,
            width=20
        )
        ct_l_btn.pack(pady=10)

    def create_status_monitor(self):
        """创建状态监控组件"""
        pass
        
    def start_status_monitor_thread(self):
        """启动状态监控线程"""
        def monitor_queue():
            while True:
                try:
                    # 非阻塞地从队列获取消息
                    message = self.queue.get_nowait()
                    self.handle_message(message)
                except Empty:
                    # 队列为空时休息一下
                    time.sleep(0.1)
                except Exception as e:
                    self.update_status(f"监控线程错误: {str(e)}")
                    break
        
        # 创建并启动监控线程
        self.monitor_thread = threading.Thread(target=monitor_queue, daemon=True)
        self.monitor_thread.start()
        
    def handle_message(self, message):
        """处理来自子进程的消息"""
        msg_type = message.get("type", "unknown")
        process_name = message.get("process", "unknown")
        timestamp = message.get("timestamp", time.time())
        
        if msg_type == "status":
            status = message.get("status", "unknown")
            self.update_status(f"[{time.strftime('%H:%M:%S', time.localtime(timestamp))}] {process_name} - {status}")
        elif msg_type == "error":
            error = message.get("error", "unknown error")
            self.update_status(f"[{time.strftime('%H:%M:%S', time.localtime(timestamp))}] {process_name} - 错误: {error}")
        elif msg_type == "result":
            result = message.get("result", "")
            self.update_status(f"[{time.strftime('%H:%M:%S', time.localtime(timestamp))}] {process_name} - 结果: {result}")
    
    def update_status(self, message):
        """更新状态列表"""
        self.status_list.insert(tk.END, message)
        # 保持最多显示20条消息
        if self.status_list.size() > 20:
            self.status_list.delete(0)
        # 滚动到最后一条消息
        self.status_list.yview(tk.END)
    
    """zhongzi部门项目"""
    def open_Rin(self):
        """1.在新进程中打开RIN分析模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_Rin_FSV3004, self.queue, "Rin_FSV3004")
        )
        process.start()
        self.processes["Rin_FSV3004"] = process
        self.update_status(f"已启动 Rin_FSV3004 进程 (PID: {process.pid})")

    def open_Rin_4051(self):
        """1.5.在新进程中打开RIN分析模块 (4051)"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_Rin_4051, self.queue, "Rin_4051")
        )
        process.start()
        self.processes["Rin_4051"] = process
        self.update_status(f"已启动 Rin_4051 进程 (PID: {process.pid})")

    def open_LineWidth(self):
        """2.在新进程中打开线宽测量模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_LineWidth, self.queue, "LineWidth")
        )
        process.start()
        self.processes["LineWidth"] = process
        self.update_status(f"已启动 LineWidth 进程 (PID: {process.pid})")

    def open_TimeDomain(self):
        """3.在新进程中打开时域分析模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_TimeDomain, self.queue, "TimeDomain")
        )
        process.start()
        self.processes["TimeDomain"] = process
        self.update_status(f"已启动 TimeDomain 进程 (PID: {process.pid})")

    def open_SpectrumSNR(self):
        """4.在新进程中打开信噪比分析模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_SpectrumSNR, self.queue, "SpectrumSNR")
        )
        process.start()
        self.processes["SpectrumSNR"] = process
        self.update_status(f"已启动 SpectrumSNR 进程 (PID: {process.pid})")

    def open_SingleFrequency(self):
        """5.在新进程中打开单频测量模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_SingleFrequency, self.queue, "SingleFrequency")
        )
        process.start()
        self.processes["SingleFrequency"] = process
        self.update_status(f"已启动 SingleFrequency 进程 (PID: {process.pid})")

    """qijian部门项目"""
    def open_CT_W(self):
        """6.在新进程中打开电流温度_波长模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_CT_W, self.queue, "CT_W")
        )
        process.start()
        self.processes["CT_W"] = process
        self.update_status(f"已启动 CT_W 进程 (PID: {process.pid})")

    def open_CT_P(self):
        """7.在新进程中打开电流温度_功率模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_CT_P, self.queue, "CT_P")
        )
        process.start()
        self.processes["CT_P"] = process
        self.update_status(f"已启动 CT_P 进程 (PID: {process.pid})")

    def open_CT_L(self):
        """8.在新进程中打开电流温度_线宽模块"""
        process = multiprocessing.Process(
            target=start_gui_process,
            args=(import_CT_L, self.queue, "CT_L")
        )
        process.start()
        self.processes["CT_L"] = process
        self.update_status(f"已启动 CT_L 进程 (PID: {process.pid})")

if __name__ == "__main__":
    # 在Windows上支持多进程，特别是在打包成可执行文件时
    multiprocessing.freeze_support()
    
    root = tk.Tk()
    app = MainApplication(root)
    root.mainloop()

# pyinstaller package.spec