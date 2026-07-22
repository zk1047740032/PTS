import tkinter as tk
from tkinter import ttk, messagebox
import importlib
import json
import os
import sys
import threading
import time
import traceback
import multiprocessing
from queue import Empty
from utils.LightSwitch import OpticalSwitch
from utils.test_result_dialog import TestResultDialog
from utils.theme import (
    dpix,
    COLOR_BG, COLOR_WHITE, COLOR_CARD_BG, COLOR_CARD_BORDER,
    COLOR_SECTION_BG, COLOR_SECTION_FG,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    COLOR_ENTRY_BORDER, COLOR_ENTRY_BG,
    COLOR_ACCENT, COLOR_ACCENT_HOVER,
    COLOR_SUCCESS, COLOR_WARNING,
    COLOR_LOG_ERROR, COLOR_LOG_COMPLETED, COLOR_LOG_RUNNING, COLOR_LOG_CLICKABLE,
    FONT_FAMILY, FONT_SIZE_TITLE, FONT_SIZE_HEADING,
    FONT_SIZE_BODY, FONT_SIZE_SMALL, FONT_SIZE_CAPTION,
    PADDING_SECTION, PADDING_CARD, PADDING_ROW,
)
from core.config import CFG
# ==========================================
# 动态导入辅助函数
# ==========================================

# DPI 缩放工具由 utils.theme 统一提供 (from utils.theme import dpix)


# ---- 基准尺寸（以 96 DPI / 100% 系统缩放为基准设计） ----
_BASE_WINDOW_W      = 666   # 主窗口宽度
_BASE_WINDOW_H      = 400    # 主窗口高度
_BASE_LEFT_PANEL_W  = 275    # 左侧控制面板宽度
_BASE_RESULT_W      = 370    # 测试结果弹窗宽度
_BASE_RESULT_H      = 410    # 测试结果弹窗高度
_BASE_HELP_W        = 650   # 说明文档窗口宽度
_BASE_HELP_H        = 480    # 说明文档窗口高度
_BASE_SEED_W        = 330    # 种子参数窗口宽度
_BASE_SEED_H        = 540    # 种子参数窗口高度
_BASE_CONFIG_W      = 585    # 配置参数窗口宽度
_BASE_CONFIG_H      = 620    # 配置参数窗口高度
_BASE_NETWORK_W     = 337    # 网络配置窗口宽度
_BASE_NETWORK_H     = 380    # 网络配置窗口高度

# ==========================================
# 模块注册表 —— 新部门只需改这里 + MODULE_MAP + MODULE_GROUPS
# ==========================================
MODULE_REGISTRY = {
    "Rin_FSV3004":    ("path_a.Rin_FSV3004",    "RinGUI"),
    "线宽_FSV3004":   ("path_a.LineWidth_FSV3004", "LineWidth_FSV3004_GUI"),
    "时域":           ("path_a.TimeDomain",      "TimeDomainGUI"),
    "信噪比":         ("path_a.SpectrumSNR",     "SpectrumSNRGUI"),
    "单频":           ("path_a.SingleFrequency", "SingleFrequencyGUI"),
    "功率":           ("path_b.Power",           "PowerGUI"),
    "PZT调制":        ("path_a.WaveLength",      "WaveLengthTestGUI"),
    "相噪":           ("path_a.PhaseNoise",      "PhaseNoiseGUI"),
}

# 【修改点 1】：函数签名增加 cmd_queue (命令队列)
def run_module_process(module_name, start_method, msg_queue, cmd_queue, one_click_test=False):
    """
        运行指定模块的进程，负责初始化 GUI 界面并处理测试命令

    该函数根据模块名称导入对应的 GUI 类，创建实例，并监听命令队列以执行测试。
    同时通过消息队列向主进程报告模块状态。

    参数:
        module_name (str): 模块名称，用于确定要导入的 GUI 类
        start_method (str): 启动测试的方法名称，当收到 START 命令时会调用该方法
        msg_queue (Queue): 消息队列，用于向主进程发送状态消息
        cmd_queue (Queue): 命令队列，用于接收主进程发送的命令（如 START）
        one_click_test (bool): 是否由"一键测试"按钮启动。影响测试结束后窗口是否自动关闭。

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
        # 通过注册表动态导入，新部门只需改 MODULE_REGISTRY
        if module_name not in MODULE_REGISTRY:
            raise ValueError(f"未知模块: {module_name}")
        module_path, class_name = MODULE_REGISTRY[module_name]
        mod = importlib.import_module(module_path)
        gui_class = getattr(mod, class_name)

        # 子进程启动时加载用户配置覆盖到 CFG，确保 GUI 显示的是最新参数
        try:
            from utils.config_params_panel import apply_user_overrides_to_cfg
            apply_user_overrides_to_cfg()
        except Exception:
            pass

        msg_queue.put((module_name, "running", f"正在启动 {module_name} 窗口..."))

        app_instance = gui_class(None)
        app_instance._one_click_test = one_click_test

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
                - 执行测试方法（各模块的 start_xxx 方法为非阻塞，后台线程执行实际测试）
                - 处理测试过程中的异常

            注意:
                - 窗口关闭由各模块在测试线程完成后自行调用 root.destroy()，
                  不在 trigger_test 中统一处理，因为各模块是异步执行的。
            """
            try:
                if start_method and hasattr(app_instance, start_method):
                    msg_queue.put((module_name, "running", f"{module_name} 测试开始..."))
                    try:
                        app_instance.root.title(f"{module_name} [运行中...]")
                    except: pass

                    method = getattr(app_instance, start_method)
                    method() # 启动测试（非阻塞，各模块内创建后台线程）
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
    "Rin_FSV3004": {"start_method": "start_rin", "group": "ch3"},
    "线宽_FSV3004": {"start_method": "start_measurement", "group": "ch3"},
    "时域": {"start_method": "start_test", "group": "ch2"},
    "信噪比": {"start_method": "start_test", "group": "ch3"},
    "单频": {"start_method": "start", "group": "ch3"},
    "功率": {"start_method": "start_collect", "group": "path_b"},
    "PZT调制": {"start_method": "start_test", "group": "ch1"},
    "相噪": {"start_method": "start_test", "group": "ch4"},
}

MODULE_GROUPS = {
    "光路A": {
        "通道1": [name for name, info in MODULE_MAP.items() if info["group"] == "ch1"],
        "通道2": [name for name, info in MODULE_MAP.items() if info["group"] == "ch2"],
        "通道3": [name for name, info in MODULE_MAP.items() if info["group"] == "ch3"],
        "通道4": [name for name, info in MODULE_MAP.items() if info["group"] == "ch4"]
    },
    "光路B": [name for name, info in MODULE_MAP.items() if info["group"] == "path_b"],
}

# ==========================================
# 光开关配置
# ==========================================
OPTICAL_SWITCH_VISA = CFG.usb.optical_switch
CHANNEL_SWITCH_DELAY = CFG.timing.channel_switch_delay_s

class IntegratedPlatform:
    """
    集成测试平台主类
    """
    def __init__(self, root):
        """
        初始化集成测试平台
        
        参数:
            root (tk.Tk): Tkinter 根窗口对象
        """
        self.root = root
        self.root.title("PTS-种子")
        # 窗口居中
        ww, wh = dpix(_BASE_WINDOW_W), dpix(_BASE_WINDOW_H)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - ww) // 2
        y = (sh - wh) // 2
        self.root.geometry(f"{ww}x{wh}+{x}+{y}")
        try:
            self.root.iconbitmap("PreciLasers.ico")
        except:
            pass

        self.check_vars = {}
        self.processes = {}
        self.cmd_queues = {}
        self.msg_queue = multiprocessing.Queue()

        # 加载用户配置覆盖（需在 setup_ui 之前，以便模块窗口读取到覆盖后的 CFG 值）
        self._load_user_config()

        self.setup_ui()

        # 默认显示光路A → 通道3
        self.nb.select(0)            # 光路A（第一个一级标签页）
        self.sub_nb.select(2)        # 通道3（第三个二级标签页，索引从0开始）

        # 初始化光开关（不自动连接，等一键测试时再连接）
        self.optical_switch = OpticalSwitch(
            OPTICAL_SWITCH_VISA,
            log_func=lambda msg: self.log("光开关", msg)
        )

        self.root.after(100, self.process_queue_messages)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        """
        设置用户界面

        该方法创建并配置整个测试平台的用户界面，包括：
        - 左侧控制面板，包含测试项目选择、全选/清空按钮和标签页
        - 右侧日志监控区域，包含进度条、状态标签和日志树视图
        """
        # ---- ttk 主题 ----
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self._configure_ttk_styles()

        # ---- 主布局 ----
        main_frame = tk.Frame(self.root, bg=COLOR_BG)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 左侧：控制面板 ===
        self._build_left_panel(main_frame)

        # === 分隔线 ===
        sep = tk.Frame(main_frame, bg=COLOR_CARD_BORDER, width=1)
        sep.pack(side=tk.LEFT, fill=tk.Y)

        # === 右侧：日志监控 ===
        self._build_right_panel(main_frame)

    # ---------- ttk 样式 ----------

    def _configure_ttk_styles(self):
        """统一配置 ttk 控件样式"""
        style = self.style
        F = FONT_FAMILY
        B = FONT_SIZE_BODY

        # 全局默认
        style.configure(".", font=(F, B), background=COLOR_BG)

        # Notebook（标签页）
        style.configure("TNotebook", background=COLOR_BG, borderwidth=0)
        style.configure("TNotebook.Tab",
                        font=(F, B, "bold"), padding=[18, 6],
                        background=COLOR_SECTION_BG, foreground=COLOR_TEXT_SECONDARY,
                        borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", COLOR_WHITE)],
                  foreground=[("selected", COLOR_ACCENT)])

        # 子 Notebook（通道标签）
        style.configure("Sub.TNotebook", background=COLOR_CARD_BG, borderwidth=0)
        style.configure("Sub.TNotebook.Tab",
                        font=(F, B), padding=[12, 4],
                        background=COLOR_CARD_BG, foreground=COLOR_TEXT_SECONDARY,
                        borderwidth=0)
        style.map("Sub.TNotebook.Tab",
                  background=[("selected", COLOR_ACCENT)],
                  foreground=[("selected", COLOR_WHITE)])

        # Treeview
        style.configure("Treeview",
                        background=COLOR_WHITE, fieldbackground=COLOR_WHITE,
                        rowheight=32, font=(F, B), borderwidth=0)
        style.configure("Treeview.Heading",
                        font=(F, B, "bold"),
                        background=COLOR_SECTION_BG, foreground=COLOR_TEXT_SECONDARY,
                        borderwidth=0, relief="flat")
        style.map("Treeview.Heading",
                  background=[("active", COLOR_SECTION_BG)])

        # Progressbar
        style.configure("TProgressbar",
                        background=COLOR_ACCENT, troughcolor=COLOR_SECTION_BG,
                        borderwidth=0, thickness=6)

        # Scrollbar
        style.configure("TScrollbar",
                        background=COLOR_WHITE, troughcolor=COLOR_BG,
                        borderwidth=0, arrowcolor=COLOR_TEXT_SECONDARY, arrowsize=14)

        # Checkbutton
        style.configure("TestCheckbutton.TCheckbutton",
                        background=COLOR_CARD_BG, foreground=COLOR_TEXT_PRIMARY,
                        font=(F, B))
        style.map("TestCheckbutton.TCheckbutton",
                  background=[("active", COLOR_CARD_BG), ("selected", COLOR_CARD_BG)])

        # Frame（统一背景色）
        style.configure("TFrame", background=COLOR_BG)

    # ---------- 左侧面板 ----------

    def _build_left_panel(self, parent):
        """构建左侧控制面板"""
        panel = tk.Frame(parent, bg=COLOR_BG, width=dpix(_BASE_LEFT_PANEL_W))
        panel.pack(side=tk.LEFT, fill=tk.Y)
        panel.pack_propagate(False)

        # ---- 标题 ----
        header = tk.Frame(panel, bg=COLOR_BG)
        header.pack(fill=tk.X, padx=PADDING_SECTION, pady=(20, 0))

        accent = tk.Frame(header, bg=COLOR_ACCENT, width=4, height=24)
        accent.pack(side=tk.LEFT, padx=(0, 8))
        accent.pack_propagate(False)

        tk.Label(header, text="测试项目选择",
                 font=(FONT_FAMILY, FONT_SIZE_HEADING, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_BG).pack(side=tk.LEFT)

        tk.Label(header, text="单独测试或一键测试",
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_TEXT_MUTED, bg=COLOR_BG).pack(side=tk.LEFT, padx=(8, 0), pady=(6, 0))

        # ---- 全选 / 清空 ----
        btn_row = tk.Frame(panel, bg=COLOR_BG)
        btn_row.pack(fill=tk.X, padx=PADDING_SECTION, pady=(12, 6))

        self._make_secondary_button(btn_row, "全选", self.select_all).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self._make_secondary_button(btn_row, "清空", self.deselect_all).pack(
            side=tk.RIGHT, fill=tk.X, expand=True, padx=(4, 0))

        # ---- 测试项卡片 ----
        card = tk.Frame(panel, bg=COLOR_CARD_BG,
                highlightbackground=COLOR_CARD_BORDER,
                highlightthickness=1,
                height=dpix(170))   # 设一个最小高度
        card.pack(fill=tk.BOTH, expand=True, padx=PADDING_SECTION, pady=(0, 10))
        card.pack_propagate(False)          # 阻止内容撑大/缩小时覆盖 height

        self._build_notebook(card)

        # ---- 底部操作按钮 ----
        bottom = tk.Frame(panel, bg=COLOR_BG)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, padx=PADDING_SECTION, pady=(0, 16))

        self.btn_open = self._make_primary_button(bottom, "打开", self.open_selected_windows)
        self.btn_open.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))

        self.btn_result = self._make_warning_button(bottom, "测试结果", self.show_test_result_dialog)
        self.btn_result.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=3)

        self.btn_run = self._make_success_button(bottom, "一键测试", self.run_selected_tests)
        self.btn_run.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(3, 0))

    def _build_notebook(self, parent):
        """构建测试项标签页"""
        self.nb = ttk.Notebook(parent)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        for group_name, module_list in MODULE_GROUPS.items():
            tab_frame = tk.Frame(self.nb, bg=COLOR_CARD_BG)
            self.nb.add(tab_frame, text=f"  {group_name}  ")

            if group_name == "光路A":
                # 嵌套子标签页
                self.sub_nb = ttk.Notebook(tab_frame, style="Sub.TNotebook")
                self.sub_nb.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

                for channel_name, channel_modules in module_list.items():
                    ch_frame = tk.Frame(self.sub_nb, bg=COLOR_CARD_BG)
                    self.sub_nb.add(ch_frame, text=f"  {channel_name}  ")
                    self._build_checkbox_list(ch_frame, channel_modules)
            else:
                self._build_checkbox_list(tab_frame, module_list)

    def _build_checkbox_list(self, parent, module_names):
        """在可滚动画布中创建复选框列表"""
        canvas = tk.Canvas(parent, bg=COLOR_CARD_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=COLOR_CARD_BG)

        scroll_frame.bind(
            "<Configure>",
            lambda *_: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda *_: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda *_: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for name in module_names:
            var = tk.BooleanVar()
            self.check_vars[name] = var
            cb = ttk.Checkbutton(
                scroll_frame, text=name, variable=var,
                command=lambda n=name: self.on_test_item_checked(n),
                style="TestCheckbutton.TCheckbutton"
            )
            cb.bind("<Double-1>",
                    lambda e, n=name, w=cb: self.on_test_item_double_click(e, n, w))
            cb.pack(anchor="w", padx=14, pady=6)

    # ---------- 右侧面板 ----------

    def _build_right_panel(self, parent):
        """构建右侧日志监控面板"""
        panel = tk.Frame(parent, bg=COLOR_BG)
        panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # ---- 标题行 ----
        header = tk.Frame(panel, bg=COLOR_BG)
        header.pack(fill=tk.X, padx=PADDING_SECTION, pady=(20, 0))

        accent = tk.Frame(header, bg=COLOR_ACCENT, width=4, height=24)
        accent.pack(side=tk.LEFT, padx=(0, 8))
        accent.pack_propagate(False)

        tk.Label(header, text="运行状态监控",
                 font=(FONT_FAMILY, FONT_SIZE_HEADING, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_BG).pack(side=tk.LEFT)

        self.btn_seed_params = self._make_secondary_button(
            header, "种子参数", self.show_seed_params)
        self.btn_seed_params.pack(side=tk.RIGHT, padx=(4, 0))

        self.btn_config_params = self._make_secondary_button(
            header, "配置参数", self.show_config_params)
        self.btn_config_params.pack(side=tk.RIGHT, padx=(0, 4))

        # ---- 进度条 ----
        self.progress = ttk.Progressbar(panel, mode='determinate')
        self.progress.pack(fill=tk.X, padx=PADDING_SECTION, pady=(10, 0))

        # ---- 状态 / 操作行 ----
        status_row = tk.Frame(panel, bg=COLOR_BG)
        status_row.pack(fill=tk.X, padx=PADDING_SECTION, pady=(6, 0))

        self.status_label = tk.Label(
            status_row, text="就绪",
            font=(FONT_FAMILY, FONT_SIZE_SMALL),
            bg=COLOR_BG, fg=COLOR_TEXT_MUTED
        )
        self.status_label.pack(side=tk.LEFT)

        btn_group = tk.Frame(status_row, bg=COLOR_BG)
        btn_group.pack(side=tk.RIGHT)

        self.btn_network = self._make_secondary_button(btn_group, "网络配置", self.show_network_config)
        self.btn_network.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_help = self._make_secondary_button(btn_group, "说明文档", self.show_help)
        self.btn_help.pack(side=tk.LEFT, padx=(4, 0))

        # ---- 日志卡片 ----
        log_card = tk.Frame(panel, bg=COLOR_CARD_BG,
                            highlightbackground=COLOR_CARD_BORDER,
                            highlightthickness=1)
        log_card.pack(fill=tk.BOTH, expand=True,
                      padx=PADDING_SECTION, pady=(6, PADDING_SECTION))

        self.log_tree = ttk.Treeview(
            log_card, columns=("Time", "Module", "Message"),
            show="headings"
        )
        self.log_tree.heading("Time", text="时间")
        self.log_tree.column("Time", width=80, stretch=False, anchor="center")
        self.log_tree.heading("Module", text="模块")
        self.log_tree.column("Module", width=110, stretch=False, anchor="w")
        self.log_tree.heading("Message", text="消息内容")
        self.log_tree.column("Message", minwidth=200, stretch=True, anchor="w")

        vsb = ttk.Scrollbar(log_card, orient="vertical", command=self.log_tree.yview)
        self.log_tree.configure(yscrollcommand=vsb.set)

        self.log_tree.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=2)
        vsb.pack(side="right", fill="y", padx=(0, 2), pady=2)

    # ---------- 按钮工厂 ----------

    def _make_primary_button(self, parent, text, command):
        """蓝色主操作按钮"""
        return tk.Button(
            parent, text=text,
            font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
            bg=COLOR_ACCENT, fg=COLOR_WHITE,
            activebackground=COLOR_ACCENT_HOVER, activeforeground=COLOR_WHITE,
            relief="flat", bd=0, padx=10, pady=6,
            command=command, cursor="hand2"
        )

    def _make_success_button(self, parent, text, command):
        """绿色成功按钮"""
        return tk.Button(
            parent, text=text,
            font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
            bg=COLOR_SUCCESS, fg=COLOR_WHITE,
            activebackground="#2F8A4C", activeforeground=COLOR_WHITE,
            relief="flat", bd=0, padx=10, pady=6,
            command=command, cursor="hand2"
        )

    def _make_warning_button(self, parent, text, command):
        """橙色次级强调按钮"""
        return tk.Button(
            parent, text=text,
            font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
            bg=COLOR_WARNING, fg=COLOR_WHITE,
            activebackground="#E08A3A", activeforeground=COLOR_WHITE,
            relief="flat", bd=0, padx=10, pady=6,
            command=command, cursor="hand2"
        )

    def _make_secondary_button(self, parent, text, command):
        """灰色次级按钮（带柔和边框，区别于背景）"""
        return tk.Button(
            parent, text=text,
            font=(FONT_FAMILY, FONT_SIZE_BODY),
            bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
            activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
            relief="solid", bd=1,
            highlightbackground="#D5D9DF", highlightthickness=1,
            padx=12, pady=5,
            command=command, cursor="hand2"
        )

    # ================= 逻辑控制 =================

    def log(self, module, msg, level="info", file_path=None):
        """
        记录日志信息
        
        参数:
            module (str): 模块名称
            msg (str): 日志消息内容
            level (str, optional): 日志级别，默认为 "info"
                可选值: "info", "error", "completed", "running"
            file_path (str, optional): 文件路径，如果提供则日志可点击打开
        """
        timestamp = time.strftime("%H:%M:%S")
        tags = (level,)
        
        if file_path and os.path.exists(file_path):
            display_msg = f"{msg}"
            item_id = self.log_tree.insert("", "end", values=(timestamp, module, display_msg), tags=tags + ("clickable",))
            self.log_tree.item(item_id, open=False)
        else:
            self.log_tree.insert("", "end", values=(timestamp, module, msg), tags=tags)
        
        self.log_tree.yview_moveto(1)
        
        if level == "error":
            self.log_tree.tag_configure("error", foreground=COLOR_LOG_ERROR)
        elif level == "completed":
            self.log_tree.tag_configure("completed", foreground=COLOR_LOG_COMPLETED)
        elif level == "running":
            self.log_tree.tag_configure("running", foreground=COLOR_LOG_RUNNING)

        if file_path and os.path.exists(file_path):
            self.log_tree.tag_configure("clickable", foreground=COLOR_LOG_CLICKABLE,
                                        font=(FONT_FAMILY, FONT_SIZE_BODY, "underline"))
        
        if file_path:
            self.log_tree.tag_bind("clickable", "<Double-Button-1>", lambda e, fp=file_path: self._open_file(fp))

    def _open_file(self, file_path: str):
        """双击打开文件"""
        try:
            if os.name == 'nt':
                os.startfile(file_path)
            else:
                import subprocess
                subprocess.Popen(['xdg-open', file_path])
        except Exception as e:
            self.log("错误", f"无法打开文件:\n{file_path}\n\n{str(e)}")

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
    
    def show_network_config(self):
        """打开网络配置弹窗"""
        from utils.network_config import NetworkConfigDialog
        NetworkConfigDialog(self.root, width=dpix(_BASE_NETWORK_W), height=dpix(_BASE_NETWORK_H))

    def show_help(self):
        """
        显示操作说明文档

        功能:
            - 创建一个新的窗口显示操作说明文档
            - 从 user_guide 模块导入 USER_GUIDE 内容并显示
            - 提供滚动条和关闭按钮
        """
        from utils.user_guide import USER_GUIDE

        help_window = tk.Toplevel(self.root)
        help_window.title("操作说明")
        help_window.geometry(f"{dpix(_BASE_HELP_W)}x{dpix(_BASE_HELP_H)}")
        help_window.resizable(True, True)
        help_window.configure(bg=COLOR_BG)
        help_window.transient(self.root)

        try:
            help_window.iconbitmap("PreciLasers.ico")
        except Exception:
            pass

        # ---- 标题栏 ----
        header = tk.Frame(help_window, bg=COLOR_BG)
        header.pack(fill=tk.X, padx=PADDING_SECTION, pady=(16, 0))

        accent = tk.Frame(header, bg=COLOR_ACCENT, width=4, height=22)
        accent.pack(side=tk.LEFT, padx=(0, 10))
        accent.pack_propagate(False)

        tk.Label(header, text="操作说明",
                 font=(FONT_FAMILY, FONT_SIZE_HEADING, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_BG).pack(side=tk.LEFT)

        tk.Label(header, text="PTS 频准测试系统使用指南",
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_TEXT_MUTED, bg=COLOR_BG).pack(side=tk.LEFT, padx=(14, 0), pady=(6, 0))

        # ---- 内容卡片 ----
        card = tk.Frame(help_window, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_ENTRY_BORDER,
                        highlightthickness=1, bd=0)
        card.pack(fill=tk.BOTH, expand=True, padx=PADDING_SECTION, pady=(12, 0))

        # 内边距容器
        inner = tk.Frame(card, bg=COLOR_CARD_BG)
        inner.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        scrollbar = ttk.Scrollbar(inner)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        help_text = tk.Text(inner, wrap=tk.WORD, yscrollcommand=scrollbar.set,
                           font=(FONT_FAMILY, FONT_SIZE_BODY),
                           bg=COLOR_CARD_BG, fg=COLOR_TEXT_PRIMARY,
                           relief="flat", bd=0, highlightthickness=0,
                           padx=14, pady=10,
                           insertbackground=COLOR_ACCENT,
                           selectbackground=COLOR_ACCENT,
                           selectforeground=COLOR_WHITE)
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=help_text.yview)

        help_text.insert(tk.END, USER_GUIDE)
        help_text.config(state=tk.DISABLED)

        # ---- 底部按钮栏 ----
        bottom = tk.Frame(help_window, bg=COLOR_BG)
        bottom.pack(fill=tk.X, padx=PADDING_SECTION, pady=(8, 14))

        close_btn = tk.Button(bottom, text="关闭",
                              font=(FONT_FAMILY, FONT_SIZE_BODY),
                              bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
                              activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
                              relief="solid", bd=1, padx=16, pady=5,
                              command=help_window.destroy, cursor="hand2")
        close_btn.pack(side=tk.RIGHT)

        # ---- 居中 ----
        help_window.update_idletasks()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        px = self.root.winfo_x()
        py = self.root.winfo_y()
        dw = help_window.winfo_width()
        dh = help_window.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        help_window.geometry(f"+{x}+{y}")

    def show_seed_params(self):
        """
        显示种子参数查询窗口

        功能:
            - 创建一个新的窗口显示种子激光器参数查询界面
            - 用户配置串口号和地址后，点击查询按钮单次查询参数
        """
        from utils.seed_value import SeedParamsPanel

        seed_window = tk.Toplevel(self.root)
        seed_window.title("种子参数")
        seed_window.geometry(f"{dpix(_BASE_SEED_W)}x{dpix(_BASE_SEED_H)}")
        seed_window.resizable(True, True)
        seed_window.configure(bg=COLOR_BG)
        seed_window.transient(self.root)

        try:
            seed_window.iconbitmap("PreciLasers.ico")
        except Exception:
            pass

        # 嵌入查询面板（含底部按钮）
        SeedParamsPanel(seed_window)

        # 居中于父窗口
        seed_window.update_idletasks()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        px = self.root.winfo_x()
        py = self.root.winfo_y()
        dw = seed_window.winfo_width()
        dh = seed_window.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        seed_window.geometry(f"+{x}+{y}")

    def show_config_params(self):
        """
        显示配置参数面板

        功能:
            - 创建弹窗展示 path_a / path_b 下所有 GUI 程序的默认参数
            - 参数按程序分类，支持在线编辑和保存
            - 保存后自动写入 config/user_config.json，重启生效
        """
        from utils.config_params_panel import ConfigParamsPanel

        config_window = tk.Toplevel(self.root)
        config_window.title("配置参数")
        config_window.geometry(f"{dpix(_BASE_CONFIG_W)}x{dpix(_BASE_CONFIG_H)}")
        config_window.resizable(True, True)

        # 先让窗口壳子显示出来，再构建内容（避免用户感觉卡顿）
        config_window.update_idletasks()

        # 居中于父窗口
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        px = self.root.winfo_x()
        py = self.root.winfo_y()
        dw = config_window.winfo_width()
        dh = config_window.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        config_window.geometry(f"+{x}+{y}")

        # 嵌入配置面板
        ConfigParamsPanel(config_window)

    def _load_user_config(self):
        """
        启动时加载用户配置覆盖

        从 config/user_config.json 读取用户修改过的参数值，
        通过 object.__setattr__ 写回 CFG（绕过 frozen 限制）。
        """
        try:
            from utils.config_params_panel import apply_user_overrides_to_cfg
            apply_user_overrides_to_cfg()
        except Exception:
            pass  # 配置加载失败不影响主流程

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
            args=(name, start_method, self.msg_queue, cmd_q, auto_start),
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
        一键测试：按通道顺序编排测试

        光路A（ch3/ch2/ch1）通过光开关切换，需按通道顺序串行执行，
        同一通道内的多个模块可并行运行。
        光路B不经过光开关，与光路A并行执行。

        执行流程:
            1. 按 MODULE_MAP["group"] 将勾选项分为 ch3/ch2/ch1/光路B 四组
            2. 立即启动 光路B组，无需等待光开关
            3. 后台线程按 ch3 → ch2 → ch1 顺序编排光路A：
                a. 切换光开关到目标通道
                b. 等待稳定延时
                c. 并行启动该通道所有模块
                d. 轮询等待该通道全部完成
        """
        selected = [name for name, var in self.check_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("提示", "请先勾选测试项")
            return

        self.btn_run.config(state="disabled", text="测试中...")
        self.log("SYSTEM", f"准备执行任务: {', '.join(selected)}")

        # 按 group 分组
        channel_groups = {"ch1": [], "ch2": [], "ch3": [], "ch4": [], "path_b": []}
        for name in selected:
            group = MODULE_MAP[name]["group"]
            if group in channel_groups:
                channel_groups[group].append(name)

        # 光路B（path_b）：立即启动，独立并行
        for name in channel_groups.get("path_b", []):
            if name in self.processes and self.processes[name].is_alive():
                if name in self.cmd_queues:
                    self.cmd_queues[name].put("START")
                    self.log(name, "窗口已存在，发送【开始测试】指令", "running")
                else:
                    self.log(name, "错误：找不到命令队列，尝试重启进程", "error")
                    self.processes[name].terminate()
                    self.start_module_process(name, auto_start=True)
            else:
                self.start_module_process(name, auto_start=True)
            time.sleep(0.1)

        # 光路A（ch3/ch2/ch1/ch4）：后台线程按通道顺序编排
        has_optical_channels = any(
            channel_groups.get(k) for k in ["ch1", "ch2", "ch3", "ch4"]
        )
        if has_optical_channels:
            if not self.optical_switch.connect():
                self.log("光开关", "光开关连接失败，通道切换将不可用", "error")
            else:
                self.log("光开关", "光开关连接成功")

        # 传递完整 channel_groups（含 path_b），编排线程在所有测试完成后弹出结果窗口
        threading.Thread(
            target=self._orchestrate_channel_tests,
            args=(channel_groups,),
            daemon=True
        ).start()

    def _orchestrate_channel_tests(self, channel_groups: dict):
        """
        按通道顺序编排光路A的测试（在后台线程中运行）

        光路A各通道通过光开关串行执行，每个通道内的模块并行启动后等待全部完成。
        通道循环结束后等待光路B模块，最后弹出测试结果窗口。

        Args:
            channel_groups: {"ch1": [...], "ch2": [...], "ch3": [...], "ch4": [...], "path_b": [...]}
        """
        channel_map = {"ch1": 1, "ch2": 2, "ch3": 3, "ch4": 4}

        for group_key in ["ch1", "ch2", "ch3", "ch4"]:
            modules = channel_groups.get(group_key, [])
            if not modules:
                continue

            channel_num = channel_map[group_key]
            self.log("SYSTEM", f"========== 开始通道 {channel_num} 测试 ==========", "running")

            # 1. 切换光开关
            try:
                self.optical_switch.set_channel(channel_num)
                self.log("光开关", f"已切换到通道 {channel_num}")
            except Exception as e:
                self.log("光开关", f"通道切换失败: {e}", "error")
                # 切换失败不阻断流程，用户可能手动切换了

            time.sleep(CHANNEL_SWITCH_DELAY)

            # 2. 并行启动该通道所有模块
            for name in modules:
                try:
                    if name in self.processes and self.processes[name].is_alive():
                        if name in self.cmd_queues:
                            self.cmd_queues[name].put("START")
                            self.log(name, "窗口已存在，发送【开始测试】指令", "running")
                        else:
                            self.log(name, "错误：找不到命令队列，尝试重启进程", "error")
                            self.processes[name].terminate()
                            self.start_module_process(name, auto_start=True)
                    else:
                        self.start_module_process(name, auto_start=True)
                except Exception as e:
                    self.log(name, f"启动失败: {e}", "error")
                time.sleep(0.1)

            # 3. 等待该通道所有模块完成
            self._wait_for_modules(modules)
            self.log("SYSTEM", f"========== 通道 {channel_num} 测试完成 ==========", "completed")

        # 等待光路B模块（在 run_selected_tests 中已先行启动）
        path_b_modules = channel_groups.get("path_b", [])
        if path_b_modules:
            self._wait_for_modules(path_b_modules)

        # 全部完成，恢复按钮
        self.root.after(0, lambda: self.btn_run.config(state="normal", text="一键测试"))
        self.log("SYSTEM", "所有通道测试完成", "completed")

        # 弹出测试结果窗口
        self.root.after(0, lambda: TestResultDialog.show(
            self.root, on_generate_report=self.on_generate_report,
            width=dpix(_BASE_RESULT_W), height=dpix(_BASE_RESULT_H)
        ))

    def _wait_for_modules(self, module_names: list):
        """
        轮询等待指定模块全部完成（在后台线程中调用）

        检查逻辑：模块不存在于 self.processes（已被 process_queue_messages 清理）
        或进程不再存活，均视为已完成。

        Args:
            module_names: 要等待的模块名称列表
        """
        while True:
            all_done = True
            for name in module_names:
                p = self.processes.get(name)
                if p and p.is_alive():
                    all_done = False
                    break
            if all_done:
                break
            time.sleep(0.5)

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
                self.status_label.config(text=f"当前活跃窗口: {active_count}", fg=COLOR_LOG_RUNNING)
                self.progress.config(mode='indeterminate')
                self.progress.start(20)
            else:
                self.status_label.config(text="所有任务已结束", fg=COLOR_TEXT_SECONDARY)
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
            - 处理嵌套结构（如 "光路A" 分组下的通道）和扁平结构（如 "光路B" 分组）
            - 捕获异常并返回空列表
        """
        try:
            # 1. 获取一级页签 (如 "光路A" 或 "光路B")
            current_tab_id = self.nb.select()
            if not current_tab_id: return []
            
            # 注意：IntegratedPlatform 不是控件，需用 self.root.nametowidget
            top_frame = self.root.nametowidget(current_tab_id)
            current_tab_text = self.nb.tab(current_tab_id, "text").strip()

            if current_tab_text not in MODULE_GROUPS:
                return []

            group_data = MODULE_GROUPS[current_tab_text]

            # 2. 判断是否为嵌套结构 (字典即为嵌套，如 "光路A")
            if isinstance(group_data, dict):
                # 寻找一级页签下的子 Notebook 控件
                sub_notebook = None
                for child in top_frame.winfo_children():
                    if isinstance(child, ttk.Notebook):
                        sub_notebook = child
                        break
                
                if sub_notebook:
                    # 获取二级页签 (如 "通道3")
                    sub_tab_id = sub_notebook.select()
                    if sub_tab_id:
                        sub_tab_text = sub_notebook.tab(sub_tab_id, "text").strip()
                        # 返回对应通道的模块列表
                        return group_data.get(sub_tab_text, [])
                return []
            
            # 3. 扁平结构 (列表即为扁平，如 "光路B")
            elif isinstance(group_data, list):
                return group_data
                
            return []

        except Exception as e:
            print(f"获取模块列表出错: {e}")
            return []

    def select_all(self):
        """
        全选所有测试项

        功能:
            - 遍历所有模块，勾选全部复选框
            - 仅设置状态，不触发 on_test_item_checked，避免全选时误触发不需要的逻辑
        """
        for name in self.check_vars:
            self.check_vars[name].set(True)

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

    def show_test_result_dialog(self):
        """打开测试结果弹窗"""
        TestResultDialog.show(self.root, on_generate_report=self.on_generate_report,
                              width=dpix(_BASE_RESULT_W), height=dpix(_BASE_RESULT_H))

    def on_generate_report(self, on_done=None, user_fields=None):
        """点击生成报告的逻辑

        Args:
            on_done:     可选回调，报告生成完成后调用（用于调用方恢复按钮状态等）
            user_fields: 可选字典，用户从测试结果弹窗填入的字段（型号/编号/序列号）
        """
        # 数据采集已抽至 report.data_collector 模块
        from report.data_collector import assemble_report_data
        report_data = assemble_report_data()
        if user_fields:
            report_data.update(user_fields)

        # 指定模板和输出路径
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.getcwd()
        template_path = os.path.join(base_path, "report", "templates", "template_default.docx")
        save_dir = str(CFG.dirs.report)
        os.makedirs(save_dir, exist_ok=True)
        output_path = os.path.join(save_dir, f"测试报告_{time.strftime('%Y%m%d_%H%M%S')}.docx")

        import threading
        def _generate_task():
            try:
                from report.template_generator import generate_report
                generate_report(template_path, output_path, report_data)

                self.root.after(0, lambda: self.log("SYSTEM", f"{output_path}", "completed", file_path=output_path))
            except Exception as e:
                self.root.after(0, lambda e=e: self.log("错误", f"报告生成失败:\n{str(e)}"))
                self.root.after(0, lambda e=e: self.log("SYSTEM", f"报告生成失败: {str(e)}", "error"))
            finally:
                if on_done:
                    self.root.after(0, on_done)

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
        self.optical_switch.close()
        self.root.destroy()
        sys.exit(0)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = IntegratedPlatform(root)
    root.mainloop()
    # .venv\Scripts\python.exe -m PyInstaller package.spec