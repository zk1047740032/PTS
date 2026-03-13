import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import ctypes
import threading
import csv
import time
import traceback
import multiprocessing
from queue import Empty
import docxtpl
# ==========================================
# 动态导入辅助函数
# ==========================================

if os.name == 'nt':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        dpi = ctypes.windll.user32.GetDpiForSystem()
        scaling_factor = dpi / 96.0
    except Exception:
        scaling_factor = 1.0
else:
    scaling_factor = 1.0

# 【修改点 1】：函数签名增加 cmd_queue (命令队列)
def run_module_process(module_name, start_method, msg_queue, cmd_queue):
    """
        运行指定模块的进程，负责初始化 GUI 界面并处理测试命令
    
    该函数根据模块名称导入对应的 GUI 类，创建实例，并监听命令队列以执行测试。
    同时通过消息队列向主进程报告模块状态。
    
    参数:
        module_name (str): 模块名称，用于确定要导入的 GUI 类
        start_method (str): 启动测试的方法名称，当收到 START 命令时会调用该方法
        msg_queue (Queue): 消息队列，用于向主进程发送状态消息
        cmd_queue (Queue): 命令队列，用于接收主进程发送的命令（如 START）
    
    执行流程:
        1. 根据模块名称导入对应的 GUI 类
        2. 向消息队列发送模块启动状态
        3. 创建 GUI 实例并设置窗口标题
        4. 定义内部函数 trigger_test() 用于执行测试
        5. 定义内部函数 check_command_queue() 用于监听命令队列
        6. 启动命令队列监听循环
        7. 运行 GUI 主循环
        8. 向消息队列发送模块完成状态
    
    异常处理:
        - 捕获导入模块失败的异常
        - 捕获执行测试过程中的异常
        - 捕获进程运行过程中的异常
        - 所有异常都会通过消息队列向主进程报告
    """
    try:
        gui_class = None
        # 导入逻辑保持不变
        if module_name == "Rin_FSV3004":
            from zhongzi.Rin_FSV3004 import RinGUI as gui_class
        elif module_name == "Rin_4051":
            from zhongzi.Rin_4051 import Rin_4051_GUI as gui_class
        elif module_name == "线宽_FSV3004":
            from zhongzi.LineWidth_FSV3004 import LineWidth_FSV3004_GUI as gui_class

        elif module_name == "时域":
            from zhongzi.TimeDomain import TimeDomainGUI as gui_class
        elif module_name == "信噪比":
            from zhongzi.SpectrumSNR import SpectrumSNRGUI as gui_class
        elif module_name == "单频":
            from zhongzi.SingleFrequency import SingleFrequencyGUI as gui_class
        elif module_name == "Power":
            from zhongzi.Power import PowerGUI as gui_class
        elif module_name == "CT-波长":
            from qijian.CT_W import CT_W_GUI as gui_class
        elif module_name == "CT-功率":
            from qijian.CT_P import CT_P_GUI as gui_class
        elif module_name == "CT-线宽":
            from qijian.CT_L import CT_L_GUI as gui_class
        
        if not gui_class:
            raise ValueError(f"未知模块: {module_name}")

        msg_queue.put((module_name, "running", f"正在启动 {module_name} 窗口..."))

        app_instance = gui_class(None)
        
        try:
            app_instance.root.title(f"{module_name} [就绪]")
        except:
            pass

        # === 定义执行测试的内部函数 ===
        def trigger_test():
            """
            触发测试的具体逻辑
            
            功能:
                - 检查是否存在指定的启动方法
                - 发送测试开始的消息到消息队列
                - 更新应用窗口标题为运行中状态
                - 执行测试方法
                - 处理测试过程中的异常
            
            异常处理:
                - 捕获执行测试过程中的异常，并通过消息队列上报错误信息
            """
            try:
                if start_method and hasattr(app_instance, start_method):
                    msg_queue.put((module_name, "running", f"{module_name} 测试开始..."))
                    try:
                        app_instance.root.title(f"{module_name} [运行中...]")
                    except: pass
                    
                    method = getattr(app_instance, start_method)
                    method() # 执行测试
                else:
                    msg_queue.put((module_name, "warning", f"未找到启动方法 {start_method}"))
            except Exception as e:
                msg_queue.put((module_name, "error", f"执行错误: {str(e)}"))

        # === 【修改点 2】：监听命令队列 ===
        def check_command_queue():
            """
            监听命令队列
            """
            try:
                # 非阻塞获取命令
                while not cmd_queue.empty():
                    cmd = cmd_queue.get_nowait()
                    if cmd == "START":
                        # 收到主进程的开始命令
                        trigger_test()
            except Empty:
                pass
            finally:
                # 每 200ms 检查一次
                app_instance.root.after(200, check_command_queue)

        # 启动监听循环
        app_instance.root.after(200, check_command_queue)

        # 如果启动时就要求立即测试 (Auto Start)
        if start_method and start_method != "MANUAL_ONLY": 
            # 这里的逻辑稍微调整：如果传入了 start_method，说明是"一键测试"启动的
            # 但为了统一逻辑，建议"一键测试"也通过队列发送 START，或者保留这里的延迟启动
            # 此处保留延迟启动以兼容直接新开进程的情况
            pass 
            # 注意：我在主类中修改了逻辑，如果是"一键测试"启动，会在start后立即发消息
            # 所以这里不需要自动运行，完全依赖 check_command_queue 即可
            # 或者保留 1秒后的自动运行也可以，看你喜好。
            # 为了防止重复，这里我们移除自动运行，全部由主进程发指令控制（更稳健）。

        app_instance.root.mainloop()

        msg_queue.put((module_name, "completed", f"{module_name} 窗口已关闭"))

    except Exception as e:
        msg_queue.put((module_name, "error", f"进程崩溃: {str(e)}"))
        print(f"Process Error: {e}")


# ==========================================
# 配置定义 (保持不变)
# ==========================================
MODULE_MAP = {
    "Rin_FSV3004": {"start_method": "start_rin", "group": "zhongzi"},
    "Rin_4051": {"start_method": "start_test", "group": "zhongzi"},
    "线宽_FSV3004": {"start_method": "start_measurement", "group": "zhongzi"},
    "时域": {"start_method": "start_test", "group": "zhongzi"},
    "信噪比": {"start_method": "start_test", "group": "zhongzi"},
    "单频": {"start_method": "start", "group": "zhongzi"},
    "Power": {"start_method": "start_collect", "group": "zhongzi"},
    "CT-波长": {"start_method": "start_group1", "group": "qijian"},
    "CT-功率": {"start_method": "start_group1", "group": "qijian"},
    "CT-线宽": {"start_method": "start_group1", "group": "qijian"},
}

MODULE_GROUPS = {
    "种子": {
        "通道1": [name for name, info in MODULE_MAP.items() if info["group"] == "zhongzi"],
        "通道2": [],
        "通道3": []
    },
    "器件": [name for name, info in MODULE_MAP.items() if info["group"] == "qijian"],
}

class IntegratedPlatform:
    """
    集成测试平台主类
    
    主要功能和用途:
        - 提供统一的测试项目选择界面，支持多模块并行测试
        - 管理各个测试模块的进程生命周期
        - 监控测试执行状态并记录日志
        - 提供一键测试功能，实现自动化测试流程
        - 支持测试项的全选/清空操作
    
    核心设计理念:
        - 采用多进程架构，每个测试模块运行在独立进程中，提高系统稳定性
        - 使用队列进行进程间通信，实现命令下发和状态上报
        - 模块化设计，通过配置文件定义测试模块，便于扩展
        - 响应式UI设计，提供直观的操作界面和实时状态反馈
    
    关键特性:
        - 支持多模块并行测试，提高测试效率
        - 实时日志记录和状态监控
        - 自动清理资源，确保进程安全退出
        - 容错处理，当命令队列异常时自动重启进程
        - 支持测试项的双击快速打开功能
    
    使用场景:
        - 实验室环境下的多设备并行测试
        - 自动化测试流程中的集中控制
        - 需要实时监控测试状态的场景
        - 对测试结果有详细记录需求的场合
    
    限制条件:
        - 依赖Tkinter GUI库，仅支持桌面环境
        - 测试模块必须符合特定的接口规范
        - 进程间通信依赖multiprocessing.Queue，受系统资源限制
        - 并发测试数量受系统硬件资源限制
    
    与其他类或模块的主要交互关系:
        - 与各测试模块的GUI类交互，通过进程启动和命令队列控制测试执行
        - 与user_guide模块交互，显示操作说明文档
        - 与run_module_process函数交互，启动和管理测试进程
        - 通过MODULE_MAP和MODULE_GROUPS配置定义测试模块信息
    """
    def __init__(self, root):
        """
        初始化集成测试平台
        
        参数:
            root (tk.Tk): Tkinter 根窗口对象
        """
        self.root = root
        self.root.title("PTS - 集成测试平台")
        self.root.geometry("1100x850") # 稍微加大一点
        try:
            self.root.iconbitmap("PreciLasers.ico")
        except:
            pass

        self.check_vars = {}     
        self.processes = {}
        self.cmd_queues = {}
        self.msg_queue = multiprocessing.Queue() 

        self.setup_ui()
        
        self.root.after(100, self.process_queue_messages)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        """
        设置用户界面
        
        该方法创建并配置整个测试平台的用户界面，包括：
        - 左侧控制面板，包含测试项目选择、全选/清空按钮和标签页
        - 右侧日志监控区域，包含进度条、状态标签和日志树视图
        """
        self.style = ttk.Style()
        self.style.theme_use('vista')
        
        # 【修改点 4】：修复 Treeview 行高问题
        self.style.configure("Treeview", rowheight=30, font=("Microsoft YaHei", 10))
        self.style.configure("Treeview.Heading", font=("Microsoft YaHei", 10, "bold"))
        
        # 去掉测试项选项条的背景色
        self.style.configure("TestCheckbutton.TCheckbutton", background="white", foreground="black")
        self.style.map("TestCheckbutton.TCheckbutton", background=[("active", "white")])

        # 主布局
        main_frame = tk.Frame(self.root, bg="white")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 左侧：控制面板 ===
        control_panel = tk.Frame(main_frame, bg="#ffffff", width=340)
        control_panel.pack(side=tk.LEFT, fill=tk.Y, padx=0, pady=0)
        control_panel.pack_propagate(False)

        tk.Label(control_panel, text="测试项目选择", font=("微软雅黑", 14, "bold"), bg="#ffffff").pack(pady=15)

        # 全选/清空按钮
        btn_frame = tk.Frame(control_panel, bg="#ffffff")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_frame, text="全选", command=self.select_all, width=10, cursor="hand2").pack(side=tk.LEFT, padx=1)
        ttk.Button(btn_frame, text="清空", command=self.deselect_all, width=10, cursor="hand2").pack(side=tk.RIGHT, padx=1)

        self.nb = ttk.Notebook(control_panel)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for group_name, module_list in MODULE_GROUPS.items():
            frame = ttk.Frame(self.nb)
            self.nb.add(frame, text=f" {group_name} ")
            
            if group_name == "种子":
                sub_nb = ttk.Notebook(frame)
                sub_nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
                
                for channel_name, channel_modules in module_list.items():
                    channel_frame = ttk.Frame(sub_nb)
                    sub_nb.add(channel_frame, text=f" {channel_name} ")
                    
                    canvas = tk.Canvas(channel_frame, bg="white")
                    scrollbar = ttk.Scrollbar(channel_frame, orient="vertical", command=canvas.yview)
                    scroll_frame = tk.Frame(canvas, bg="white")

                    scroll_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
                    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
                    canvas.configure(yscrollcommand=scrollbar.set)

                    canvas.pack(side="left", fill="both", expand=True)
                    scrollbar.pack(side="right", fill="y")

                    for name in channel_modules:
                        var = tk.BooleanVar()
                        self.check_vars[name] = var
                        cb = ttk.Checkbutton(scroll_frame, text=name, variable=var, 
                                             command=lambda n=name: self.on_test_item_checked(n),
                                             style="TestCheckbutton.TCheckbutton")
                        cb.bind("<Double-1>", lambda e, n=name, w=cb: self.on_test_item_double_click(e, n, w))
                        cb.pack(anchor="w", padx=10, pady=5)
            else:
                canvas = tk.Canvas(frame, bg="white")
                scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
                scroll_frame = tk.Frame(canvas, bg="white")

                scroll_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
                canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)

                canvas.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")

                for name in module_list:
                    var = tk.BooleanVar()
                    self.check_vars[name] = var
                    cb = ttk.Checkbutton(scroll_frame, text=name, variable=var, 
                                         command=lambda n=name: self.on_test_item_checked(n),
                                         style="TestCheckbutton.TCheckbutton")
                    cb.bind("<Double-1>", lambda e, n=name, w=cb: self.on_test_item_double_click(e, n, w))
                    cb.pack(anchor="w", padx=10, pady=5)

        # 底部按钮区
        bottom_frame = tk.Frame(control_panel, bg="#ffffff")
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        # 打开按钮
        self.btn_open = tk.Button(bottom_frame, text="打开", 
                                bg="#1E96E6", fg="white", font=("微软雅黑", 9, "bold"),
                                command=self.open_selected_windows, cursor="hand2")
        self.btn_open.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        # 生成报告
        self.btn_report = tk.Button(bottom_frame, text="生成报告", 
                                    bg="#FF9900", fg="white", font=("微软雅黑", 9, "bold"),
                                    command=self.on_generate_report, cursor="hand2")
        self.btn_report.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        # 一键测试按钮
        self.btn_run = tk.Button(bottom_frame, text="一键测试", 
                                bg="#02BC08", fg="white", font=("微软雅黑", 9, "bold"),
                                command=self.run_selected_tests, cursor="hand2")
        self.btn_run.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=2)

        # === 右侧：日志监控 ===
        right_panel = tk.Frame(main_frame, bg="white")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        tk.Label(right_panel, text="运行状态监控", font=("微软雅黑", 12), bg="white").pack(anchor="w", padx=10, pady=10)
        
        self.progress = ttk.Progressbar(right_panel, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10)
        
        # 添加状态和操作按钮框架
        status_frame = tk.Frame(right_panel, bg="white")
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 状态标签放在左侧
        self.status_label = tk.Label(status_frame, text="就绪", bg="white", fg="#666")
        self.status_label.pack(side=tk.LEFT)
        
        # 操作按钮放在右侧
        button_frame = tk.Frame(status_frame, bg="white")
        button_frame.pack(side=tk.RIGHT)
        
        # 清空日志按钮
        self.btn_clear_log = ttk.Button(button_frame, text="清空日志", 
                                     command=self.clear_logs, width=10, cursor="hand2")
        self.btn_clear_log.pack(side=tk.LEFT, padx=1)
        
        # 说明文档按钮
        self.btn_help = ttk.Button(button_frame, text="说明文档", 
                                 command=self.show_help, width=10, cursor="hand2")
        self.btn_help.pack(side=tk.LEFT, padx=1)

        log_frame = tk.Frame(right_panel)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_tree = ttk.Treeview(log_frame, columns=("Time", "Module", "Message"), show="headings")
        
        # 【修改点 5】：设置具体的列宽，避免挤在一起
        self.log_tree.heading("Time", text="时间")
        self.log_tree.column("Time", width=80, stretch=False, anchor="center")
        
        self.log_tree.heading("Module", text="模块")
        self.log_tree.column("Module", width=120, stretch=False, anchor="w")
        
        self.log_tree.heading("Message", text="消息内容")
        self.log_tree.column("Message", minwidth=200, stretch=True, anchor="w") # 让消息列自动填充剩余空间
        
        vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_tree.yview)
        self.log_tree.configure(yscrollcommand=vsb.set)
        
        self.log_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    # ================= 逻辑控制 =================

    def log(self, module, msg, level="info"):
        """
        记录日志信息
        
        参数:
            module (str): 模块名称
            msg (str): 日志消息内容
            level (str, optional): 日志级别，默认为 "info"
                可选值: "info", "error", "completed", "running"
        """
        timestamp = time.strftime("%H:%M:%S")
        tags = (level,)
        self.log_tree.insert("", "end", values=(timestamp, module, msg), tags=tags)
        self.log_tree.yview_moveto(1)
        
        if level == "error":
            self.log_tree.tag_configure("error", foreground="red")
        elif level == "completed":
            self.log_tree.tag_configure("completed", foreground="#008000") # 深绿色
        elif level == "running":
            self.log_tree.tag_configure("running", foreground="#0000FF")

    def on_test_item_checked(self, module_name):
        """
        测试项勾选状态变化时的处理函数
        
        参数:
            module_name (str): 模块名称
        
        功能:
            - 当测试项被取消勾选时，关闭对应进程并清理资源
            - 仅记录状态，不自动打开窗口
        """
        is_checked = self.check_vars[module_name].get()
        
        if not is_checked:
            # 取消勾选，关闭进程（如果已打开）
            if module_name in self.processes and self.processes[module_name].is_alive():
                # 发送终止信号或直接Terminate
                self.processes[module_name].terminate()
                self.log(module_name, "用户取消勾选，窗口关闭")
                
                # 清理资源
                if module_name in self.processes: del self.processes[module_name]
                if module_name in self.cmd_queues: del self.cmd_queues[module_name]
    
    def on_test_item_double_click(self, event, module_name, widget):
        """
        测试项双击时的处理函数
        
        参数:
            event: Tkinter 事件对象
            module_name (str): 模块名称
            widget: Tkinter 控件对象
        
        功能:
            - 阻止事件传播到默认的单击处理
            - 确保测试项被勾选
            - 打开对应窗口（如果进程不存在或已死亡）
            - 延迟重新启用控件，确保双击事件完全处理完成
        """
        # 阻止事件传播到默认的单击处理
        event.widget.configure(state="disabled")  # 临时禁用控件，防止第二次点击
        
        # 确保测试项被勾选
        self.check_vars[module_name].set(True)
        
        # 打开对应窗口
        if module_name not in self.processes or not self.processes[module_name].is_alive():
            self.start_module_process(module_name, auto_start=False)
        
        # 延迟重新启用控件，确保双击事件完全处理完成
        self.root.after(100, lambda w=widget: w.configure(state="normal"))
    
    def clear_logs(self):
        """
        清空日志区域
        
        功能:
            - 删除日志树视图中的所有日志条目
        """
        # 删除所有日志条目
        for item in self.log_tree.get_children():
            self.log_tree.delete(item)
    
    def show_help(self):
        """
        显示操作说明文档
        
        功能:
            - 创建一个新的窗口显示操作说明文档
            - 从 user_guide 模块导入 USER_GUIDE 内容并显示
            - 提供滚动条和关闭按钮
        """
        # 创建说明文档窗口
        help_window = tk.Toplevel(self.root)
        help_window.title("操作说明")
        help_window.geometry("1500x1000")
        help_window.resizable(False, False)
        
        # 创建滚动文本区域
        text_frame = tk.Frame(help_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        help_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, font=("微软雅黑", 10))
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=help_text.yview)
        
        # 插入说明文本 - 从独立文件引入
        from user_guide import USER_GUIDE
        help_text.insert(tk.END, USER_GUIDE)
        help_text.config(state=tk.DISABLED)  # 设置为只读
        
        # 添加关闭按钮
        close_button = ttk.Button(help_window, text="关闭", 
                               command=help_window.destroy, width=10, cursor="hand2")
        close_button.pack(pady=10)

    def start_module_process(self, name, auto_start=False):
        """
        启动模块进程
        
        参数:
            name (str): 模块名称
            auto_start (bool, optional): 是否自动开始测试，默认为 False
        
        功能:
            - 根据模块名称获取启动方法
            - 创建专属命令队列
            - 启动子进程
            - 如果 auto_start 为 True，立即发送开始测试指令
            - 记录进程启动状态
        """
        start_method = MODULE_MAP[name]["start_method"]
        
        # 创建专属命令队列
        cmd_q = multiprocessing.Queue()
        self.cmd_queues[name] = cmd_q
        
        # 即使这里传入了 start_method，子进程现在也被修改为不会自动运行
        # 而是等待 cmd_q 中的 "START" 指令
        # 为了兼容性，我们在参数里还是传进去，但主要靠下面的 put("START") 控制
        
        p = multiprocessing.Process(
            target=run_module_process,
            args=(name, start_method, self.msg_queue, cmd_q),
            daemon=True
        )
        p.start()
        self.processes[name] = p
        
        if auto_start:
            # 如果是一键启动，立即发送开始指令
            cmd_q.put("START")
            self.log(name, f"进程启动并发送测试指令 (PID: {p.pid})")
        else:
            # 仅打开窗口，不发送指令
            self.log(name, f"窗口已打开，等待测试指令 (PID: {p.pid})")

    def open_selected_windows(self):
        """
        打开所有已勾选测试项的窗口
        
        功能:
            - 获取所有已勾选的测试项
            - 如果没有勾选任何测试项，显示警告
            - 禁用打开按钮并显示正在打开状态
            - 为每个勾选的测试项启动进程（如果进程不存在或已死亡）
            - 稍作延时，避免瞬间并发过高冲击
            - 恢复按钮状态
        """
        selected = [name for name, var in self.check_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("提示", "请先勾选测试项")
            return

        self.btn_open.config(state="disabled", text="正在打开...")
        self.log("SYSTEM", f"准备打开窗口: {', '.join(selected)}")
        
        for name in selected:
            # 只有当进程不存在或已死时，才启动新进程
            if name not in self.processes or not self.processes[name].is_alive():
                self.start_module_process(name, auto_start=False)
            
            # 稍作延时，避免瞬间并发过高冲击
            time.sleep(0.1)
            
        # 恢复按钮
        self.root.after(1000, lambda: self.btn_open.config(state="normal", text="打开"))

    def run_selected_tests(self):
        """
        并发启动测试
        
        功能:
            - 获取所有已勾选的测试项
            - 如果没有勾选任何测试项，显示警告
            - 禁用一键测试按钮并显示正在下发指令状态
            - 对于每个勾选的测试项：
                - 如果窗口已打开（进程存活），通过命令队列发送开始测试指令
                - 如果找不到命令队列，尝试重启进程
                - 如果窗口未打开，启动进程并自动开始测试
            - 稍作延时，避免瞬间并发过高冲击
            - 恢复按钮状态
        """
        selected = [name for name, var in self.check_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("提示", "请先勾选测试项")
            return

        self.btn_run.config(state="disabled", text="正在下发指令...")
        self.log("SYSTEM", f"准备执行任务: {', '.join(selected)}")
        
        for name in selected:
            # 情况1: 窗口已经打开（进程存活）
            if name in self.processes and self.processes[name].is_alive():
                self.log(name, "窗口已存在，发送【开始测试】指令", "running")
                # 【关键逻辑】：通过队列发送指令
                if name in self.cmd_queues:
                    self.cmd_queues[name].put("START")
                else:
                    self.log(name, "错误：找不到命令队列，尝试重启进程", "error")
                    # 容错处理：重启
                    self.processes[name].terminate()
                    self.start_module_process(name, auto_start=True)
            
            # 情况2: 窗口未打开
            else:
                self.start_module_process(name, auto_start=True)
            
            # 稍作延时，避免瞬间并发过高冲击
            time.sleep(0.1)
            
        # 恢复按钮
        self.root.after(1000, lambda: self.btn_run.config(state="normal", text="一键测试"))

    def process_queue_messages(self):
        """
        定时处理消息队列中的消息
        
        功能:
            - 处理消息队列中的所有消息
            - 根据消息类型记录不同级别的日志
            - 当收到 completed 消息时，清理进程和命令队列引用，并取消对应测试项的勾选
            - 根据当前活跃进程数更新状态标签和进度条
            - 每 200ms 调用一次自己，实现定时处理
        """
        try:
            while not self.msg_queue.empty():
                module, type_, msg = self.msg_queue.get_nowait()
                
                if type_ == "running":
                    self.log(module, msg, "running")
                elif type_ == "completed":
                    self.log(module, msg, "completed")
                    # 进程正常退出，清理引用
                    if module in self.processes and not self.processes[module].is_alive():
                        del self.processes[module]
                        if module in self.cmd_queues: del self.cmd_queues[module]
                    # 自动取消勾选（无论进程是否存在于self.processes中）
                    if module in self.check_vars:
                        self.check_vars[module].set(False)
                            
                elif type_ == "error":
                    self.log(module, msg, "error")
                else:
                    self.log(module, msg)
                    
        except Empty:
            pass
        finally:
            active_count = sum(1 for p in self.processes.values() if p.is_alive())
            if active_count > 0:
                self.status_label.config(text=f"当前活跃窗口: {active_count}", fg="blue")
                self.progress.config(mode='indeterminate')
                self.progress.start(20)
            else:
                self.status_label.config(text="所有任务已结束", fg="black")
                self.progress.stop()
                self.progress.config(mode='determinate', value=0)

            self.root.after(200, self.process_queue_messages)

    def get_current_module_list(self):
        """
        获取当前激活页签对应的模块名称列表
        
        返回:
            list: 当前激活页签对应的模块名称列表
        
        功能:
            - 根据当前激活的页签（包括嵌套的子页签），获取对应的模块名称列表
            - 处理嵌套结构（如 "种子" 分组下的通道）和扁平结构（如 "器件" 分组）
            - 捕获异常并返回空列表
        """
        try:
            # 1. 获取一级页签 (如 "种子" 或 "器件")
            current_tab_id = self.nb.select()
            if not current_tab_id: return []
            
            # 注意：IntegratedPlatform 不是控件，需用 self.root.nametowidget
            top_frame = self.root.nametowidget(current_tab_id)
            current_tab_text = self.nb.tab(current_tab_id, "text").strip()

            if current_tab_text not in MODULE_GROUPS:
                return []

            group_data = MODULE_GROUPS[current_tab_text]

            # 2. 判断是否为嵌套结构 (字典即为嵌套，如 "种子")
            if isinstance(group_data, dict):
                # 寻找一级页签下的子 Notebook 控件
                sub_notebook = None
                for child in top_frame.winfo_children():
                    if isinstance(child, ttk.Notebook):
                        sub_notebook = child
                        break
                
                if sub_notebook:
                    # 获取二级页签 (如 "通道1")
                    sub_tab_id = sub_notebook.select()
                    if sub_tab_id:
                        sub_tab_text = sub_notebook.tab(sub_tab_id, "text").strip()
                        # 返回对应通道的模块列表
                        return group_data.get(sub_tab_text, [])
                return []
            
            # 3. 扁平结构 (列表即为扁平，如 "器件")
            elif isinstance(group_data, list):
                return group_data
                
            return []

        except Exception as e:
            print(f"获取模块列表出错: {e}")
            return []

    def select_all(self):
        """
        全选当前可见列表中的模块
        
        功能:
            - 获取当前激活页签对应的模块列表
            - 勾选所有模块的复选框
            - 仅设置状态，不触发 on_test_item_checked，避免全选时误触发不需要的逻辑
        """
        target_modules = self.get_current_module_list()
        
        for name in target_modules:
            if name in self.check_vars:
                # 仅设置状态，不触发 on_test_item_checked (避免全选时误触发不需要的逻辑)
                self.check_vars[name].set(True)
                # 如果需要在勾选时做额外处理，可以解开下面这行，但通常全选只需改变变量
                # self.on_test_item_checked(name)

    def deselect_all(self):
        """
        清空当前可见列表中的模块
        
        功能:
            - 获取当前激活页签对应的模块列表
            - 取消勾选所有模块的复选框
            - 如果模块当前是勾选状态，取消勾选并触发 on_test_item_checked 回调
            - 通过回调关闭正在运行的进程
        """
        target_modules = self.get_current_module_list()
        
        for name in target_modules:
            if name in self.check_vars:
                # 如果当前是勾选状态，才去取消
                if self.check_vars[name].get():
                    self.check_vars[name].set(False)
                    # 取消勾选必须触发回调，因为需要通过它来关闭正在运行的进程
                    self.on_test_item_checked(name)

    def on_generate_report(self):
        """点击生成报告的逻辑"""
        
        # 组装数据字典 (键名必须与 Word 模板中的 {{ 变量名 }} 完全一致)
        report_data = {
            # 项目数据
            "text_date": time.strftime("%Y/%m/%d"),

            # Fig.1 相对强度噪声（Rin）✅
            "img_rin": r"C:\PTS\zhongzi\Rin\FSV3004\Rin.png",
            "text_rin": (lambda: 
                (lambda f: float(next(csv.reader(f))[0]) if f else "未读取到积分数据")(
                    open(r"C:\PTS\zhongzi\Rin\FSV3004\rin_figure2_max.csv", 'r', encoding='utf-8') if os.path.exists(r"C:\PTS\zhongzi\Rin\FSV3004\rin_figure2_max.csv") else None
                )
            )(),

            # Fig.2 PZT（波长计）
            "img_PZT": "暂无",

            # Fig.3 光谱信噪比 ✅
            "img_spectrumSNR": r"C:\PTS\zhongzi\SpectrumSNR\spectrum.bmp",
            "text_SNR": (lambda:
                (lambda f: float(next(csv.reader(f))[0]) if f else "未读取到SNR数据")(
                    open(r"C:\PTS\zhongzi\SpectrumSNR\spectrum_snr.csv", "r", encoding='utf-8') if os.path.exists(r"C:\PTS\zhongzi\SpectrumSNR\spectrum_snr.csv") else None
                )
            )(),

            # Fig.4 线宽
            "img_linewidth": r"C:\PTS\zhongzi\LineWidth\image_1000.png", # ✅
            "text_linewidth": (lambda: # 这里错了，要的是除了2倍根号99后的结果
                (lambda f: float(list(csv.reader(f))[4][0]) if f else "未读取到Ndbdown数据")(
                    open(r"C:\PTS\zhongzi\LineWidth\ndbdown.csv", "r", encoding='utf-8') if os.path.exists(r"C:\PTS\zhongzi\LineWidth\ndbdown.csv") else None
                )
            )(),

            # Fig.5 偏振测试
            "img_polarization": r"暂无",

            # 中心波长（波长计）
            "text_center_wavelength": "{中心波长}",
            "img_center_wavelength": "波长计运行截图",

            # 输出功率 ✅
            "text_power": (lambda: 
                (lambda f: float(next(csv.reader(f))[0]) if f else "未读取到功率数据")(
                    open(r"C:\PTS\zhongzi\Power\power.csv", 'r', encoding='utf-8') if os.path.exists(r"C:\PTS\zhongzi\Power\power.csv") else None
                )
            )(),

            # 功率稳定性
            "img_power_stability": "烤机数据画图",
            "text_RMS": (lambda: 
                (lambda f: float(list(csv.reader(f))[1][1]) if f else "未读取到RMS数据")(
                    open(r"C:\PTS\zhongzi\Power\burnin_result.csv", 'r', encoding='utf-8') if os.path.exists(r"C:\PTS\zhongzi\Power\burnin_result.csv") else None
                )
            )(),
            "text_P2P": "读取整列数据 （最大值-最小值）/平均值",
        }
        
        # 指定模板和输出路径
        template_path = os.path.join(os.getcwd(), "report", "templates", "template_default.docx")
        save_dir = r"C:\PTS\report"
        os.makedirs(save_dir, exist_ok=True)
        #current_sn = "TEST"  # 默认测试序列号
        output_path = os.path.join(save_dir, f"测试报告_{time.strftime('%Y%m%d_%H%M%S')}.docx")

        import threading
        def _generate_task():
            try:
                self.btn_report.config(state="disabled", text="生成中...")
                
                # 【严格按照目录结构导入】
                from report.template_generator import generate_report
                generate_report(template_path, output_path, report_data)
                
                #self.root.after(0, lambda: messagebox.showinfo("成功", f"报告生成完毕！\n路径: {output_path}"))
                self.root.after(0, lambda: self.log("SYSTEM", f"报告生成成功: {output_path}", "completed"))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("错误", f"报告生成失败:\n{str(e)}"))
                self.root.after(0, lambda e=e: self.log("SYSTEM", f"报告生成失败: {str(e)}", "error"))
            finally:
                self.root.after(0, lambda: self.btn_report.config(state="normal", text="生成报告"))

        threading.Thread(target=_generate_task, daemon=True).start()

    def on_close(self):
        """
        关闭应用程序
        
        功能:
            - 终止所有正在运行的子进程
            - 销毁根窗口
            - 退出应用程序
        """
        for name, p in self.processes.items():
            if p.is_alive():
                p.terminate()
        self.root.destroy()
        sys.exit(0)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = IntegratedPlatform(root)
    root.mainloop()
    # pyinstaller package.spec