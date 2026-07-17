"""
测试结果弹窗模块

一键测试流程全部完成后弹出的结果汇总窗口。
展示各测试项的测量数据，并可在此窗口内生成报告。

数据来源:
    - 可编辑字段（型号/编号/序列号）：用户手动输入
    - 测量数据字段：从 report.data_collector 读取（与报告模板同源）
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable

from report.data_collector import assemble_report_data
from utils.theme import (
    COLOR_BG, COLOR_CARD_BG, COLOR_CARD_BORDER,
    COLOR_SECTION_BG, COLOR_SECTION_FG,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    COLOR_ENTRY_BORDER, COLOR_ENTRY_BG,
    COLOR_ACCENT, COLOR_ACCENT_HOVER,
)


# ---- 字段定义: (标签文本, data_collector 键名, 是否可编辑, 单位) ----
_FIELDS_EDITABLE = [
    ("激光器型号",   None, True,  ""),
    ("激光器编号",   None, True,  ""),
    ("激光器序列号", None, True,  ""),
]

_FIELDS_MEASURED = [
    ("激光器输出功率",           "text_power",             False, "mW"),
    ("当前工作波长",       "text_center_wavelength", False, "nm"),
    ("快速频率调谐范围",   "text_PZT_range",         False, "nm"),
    ("相对强度噪声(RIN)",                "text_rin",               False, "dBc/Hz"),
    ("激光线宽(100us积分)",               "text_linewidth",         False, "kHz"),
    ("光谱信噪比",         "text_SNR",               False, "dB"),
    ("单频",               None,                     False, ""),
    ("时域",               None,                     False, ""),
]

_PLACEHOLDER = "—"


class TestResultDialog:
    """一键测试结果弹窗

    在一键测试流程（含光路A各通道 + 光路B）全部完成后弹出，
    用于汇总展示本轮测试的各项结果，并可在此窗口内生成报告。

    使用方式:
        TestResultDialog.show(parent, on_generate_report=callback)
    """

    def __init__(self,
                 parent: tk.Tk,
                 test_results: Optional[dict] = None,
                 on_generate_report: Optional[Callable] = None,
                 width: int = 640,
                 height: int = 780):
        self.parent = parent
        self.test_results = test_results or {}
        self.on_generate_report = on_generate_report
        self._width = width
        self._height = height

        self._report_data = assemble_report_data()
        self._entries: dict[str, tk.Entry] = {}

        self._create_dialog()

    # ==================== 界面构建 ====================

    def _create_dialog(self):
        """创建弹窗界面"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("测试结果")
        self.dialog.geometry(f"{self._width}x{self._height}")
        self.dialog.resizable(True, True)
        self.dialog.configure(bg=COLOR_BG)
        self.dialog.transient(self.parent)

        try:
            self.dialog.iconbitmap("PreciLasers.ico")
        except Exception:
            pass

        # ---- 标题栏 ----
        self._build_header()

        # ---- 可滚动内容区 ----
        self._build_scrollable_content()

        # ---- 底部按钮 ----
        self._build_bottom_buttons()

        self._center_on_parent()

    def _build_header(self):
        """标题区域"""
        header_frame = tk.Frame(self.dialog, bg=COLOR_BG)
        header_frame.pack(fill=tk.X, padx=0, pady=(24, 0))

        # 左侧色条 + 标题
        accent_bar = tk.Frame(header_frame, bg=COLOR_ACCENT, width=4, height=28)
        accent_bar.pack(side=tk.LEFT, padx=(30, 10))
        # 禁止色条随布局伸缩
        accent_bar.pack_propagate(False)

        tk.Label(
            header_frame, text="测试结果",
            font=("微软雅黑", 17, "bold"),
            fg="#1A202C", bg=COLOR_BG
        ).pack(side=tk.LEFT)

        tk.Label(
            header_frame,
            text="各项测试数据汇总",
            font=("微软雅黑", 9),
            fg="#A0AEC0", bg=COLOR_BG
        ).pack(side=tk.LEFT, padx=(10, 0), pady=(8, 0))

    def _build_scrollable_content(self):
        """创建带滚动条的数据内容区域"""
        outer = tk.Frame(self.dialog, bg=COLOR_BG)
        outer.pack(fill=tk.BOTH, expand=True, padx=24, pady=(12, 8))

        canvas = tk.Canvas(outer, bg=COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)

        self.content_frame = tk.Frame(canvas, bg=COLOR_BG)
        self.content_frame.bind(
            "<Configure>",
            lambda *_: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        # 让内容区宽度跟随 canvas 宽度（窗口缩放时内容自适应）
        def _on_canvas_resize(event):
            canvas.itemconfig(1, width=event.width)  # window id = 1
        canvas.bind("<Configure>", _on_canvas_resize)

        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda *_: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda *_: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ---- 卡片 1：基本信息 ----
        self._build_card("基本信息", _FIELDS_EDITABLE)
        # 卡片间距
        tk.Frame(self.content_frame, bg=COLOR_BG, height=12).pack()

        # ---- 卡片 2：测试数据 ----
        self._build_card("测试数据", _FIELDS_MEASURED)

    def _build_card(self, title: str, fields: list):
        """创建一个卡片区块

        Args:
            title:  卡片标题
            fields: 字段定义列表, 每项为 (label, data_key, editable, unit)
        """
        # 卡片容器
        card = tk.Frame(self.content_frame, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_CARD_BORDER,
                        highlightthickness=1, highlightcolor=COLOR_CARD_BORDER)
        card.pack(fill=tk.X, pady=0)

        # 卡片标题行
        section_bar = tk.Frame(card, bg=COLOR_SECTION_BG, height=32)
        section_bar.pack(fill=tk.X)
        section_bar.pack_propagate(False)

        tk.Label(
            section_bar, text=f"  {title}",
            font=("微软雅黑", 10, "bold"),
            fg=COLOR_SECTION_FG, bg=COLOR_SECTION_BG,
            anchor="w"
        ).pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 字段网格
        inner = tk.Frame(card, bg=COLOR_CARD_BG)
        inner.pack(fill=tk.X, padx=14, pady=(8, 10))

        for i, (label_text, data_key, editable, unit) in enumerate(fields):
            self._create_field_row(inner, i, label_text, data_key, editable, unit)

        inner.grid_columnconfigure(1, weight=1)

    def _create_field_row(self, parent: tk.Frame, row_idx: int,
                          label_text: str, data_key: Optional[str],
                          editable: bool, unit: str):
        """在给定父容器中创建一行字段

        Args:
            parent:     父容器
            row_idx:    行号
            label_text: 字段中文名
            data_key:   assemble_report_data() 返回字典中的键
            editable:   True 创建 Entry，False 创建只读 Label
            unit:       单位字符串（可为空）
        """
        row_bg = COLOR_CARD_BG if row_idx % 2 == 0 else COLOR_BG

        # 整行背景 frame
        row_frame = tk.Frame(parent, bg=row_bg, height=36)
        row_frame.pack(fill=tk.X)
        row_frame.pack_propagate(False)
        row_frame.grid_columnconfigure(1, weight=1)

        # 标签
        tk.Label(
            row_frame, text=label_text,
            font=("微软雅黑", 10), bg=row_bg,
            fg=COLOR_TEXT_SECONDARY, anchor="e", width=16
        ).grid(row=0, column=0, sticky="e", padx=(8, 8), pady=0)

        if editable:
            self._build_editable_field(row_frame, label_text)
        else:
            self._build_readonly_field(row_frame, data_key, unit, row_bg)

    def _build_editable_field(self, parent: tk.Frame, label_text: str):
        """创建可编辑输入框（带边框包装）"""
        # 边框包装器
        border = tk.Frame(parent, bg=COLOR_ENTRY_BORDER, height=28)
        border.grid(row=0, column=1, sticky="ew", padx=(8, 14), pady=3)

        entry = tk.Entry(
            border, font=("微软雅黑", 10),
            bg=COLOR_ENTRY_BG, fg=COLOR_TEXT_PRIMARY,
            relief="flat", bd=0, highlightthickness=0,
            insertbackground=COLOR_ACCENT
        )
        entry.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self._entries[label_text] = entry

    def _build_readonly_field(self, parent: tk.Frame, data_key: Optional[str],
                              unit: str, row_bg: str):
        """创建只读显示字段"""
        raw = self._report_data.get(data_key) if data_key else None
        has_data = raw is not None and raw != "" and not str(raw).startswith("(未读取")

        if has_data:
            # 格式化数值
            if isinstance(raw, float):
                value_text = f"{raw:.2f}"
            else:
                value_text = str(raw)

            # 数值 + 单位
            value_label = tk.Label(
                parent, text=value_text,
                font=("微软雅黑", 11, "bold"),
                bg=row_bg, fg=COLOR_TEXT_PRIMARY,
                anchor="w"
            )
            value_label.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=0)

            if unit:
                tk.Label(
                    parent, text=f" {unit}",
                    font=("微软雅黑", 9),
                    bg=row_bg, fg="#8895A7",
                    anchor="w"
                ).grid(row=0, column=2, sticky="w", padx=(0, 14), pady=0)
        else:
            # 无数据
            status_dot = "●" if data_key is None else "○"
            status_color = COLOR_TEXT_MUTED
            text = _PLACEHOLDER if data_key is None else "暂无数据"

            tk.Label(
                parent, text=f"{status_dot}  {text}",
                font=("微软雅黑", 10),
                bg=row_bg, fg=status_color,
                anchor="w"
            ).grid(row=0, column=1, columnspan=2, sticky="w", padx=(8, 14), pady=0)

    def _build_bottom_buttons(self):
        """底部按钮栏"""
        bottom_frame = tk.Frame(self.dialog, bg=COLOR_BG)
        bottom_frame.pack(fill=tk.X, padx=24, pady=(4, 18))

        if self.on_generate_report is not None:
            self.btn_generate_report = tk.Button(
                bottom_frame, text="生成报告",
                font=("微软雅黑", 10, "bold"),
                bg=COLOR_ACCENT, fg="white",
                activebackground="#1E5CD6", activeforeground="white",
                relief="flat", bd=0, padx=20, pady=6,
                command=self._on_generate_report_click,
                cursor="hand2"
            )
            self.btn_generate_report.pack(side=tk.LEFT)

        close_btn = tk.Button(
            bottom_frame, text="关闭",
            font=("微软雅黑", 10),
            bg=COLOR_BG, fg="#5A6577",
            activebackground="#EEF1F5", activeforeground="#1A202C",
            relief="flat", bd=0, padx=20, pady=6,
            command=self.dialog.destroy,
            cursor="hand2"
        )
        close_btn.pack(side=tk.RIGHT)

    # ==================== 数据 ====================

    def get_field_values(self) -> dict:
        """获取所有字段的当前值（含用户编辑的可输入字段）

        Returns:
            dict: {标签文本: 当前值字符串}
        """
        all_fields = _FIELDS_EDITABLE + _FIELDS_MEASURED
        result = {}
        for label_text, data_key, editable, *_ in all_fields:
            if editable:
                entry = self._entries.get(label_text)
                result[label_text] = entry.get() if entry else ""
            else:
                raw = self._report_data.get(data_key) if data_key else None
                if raw is not None and raw != "" and not str(raw).startswith("(未读取"):
                    result[label_text] = f"{raw:.2f}" if isinstance(raw, float) else str(raw)
                else:
                    result[label_text] = ""
        return result

    # ==================== 事件处理 ====================

    def _on_generate_report_click(self):
        """点击生成报告"""
        self.btn_generate_report.config(
            state="disabled", text="生成中...",
            bg=COLOR_TEXT_MUTED, activebackground=COLOR_TEXT_MUTED
        )

        def on_done():
            try:
                self.btn_generate_report.config(
                    state="normal", text="生成报告",
                    bg=COLOR_ACCENT, activebackground=COLOR_ACCENT_HOVER
                )
            except tk.TclError:
                pass

        self.on_generate_report(on_done=on_done)

    def _center_on_parent(self):
        """将弹窗居中于父窗口"""
        self.dialog.update_idletasks()
        pw = self.parent.winfo_width()
        ph = self.parent.winfo_height()
        px = self.parent.winfo_x()
        py = self.parent.winfo_y()
        dw = self.dialog.winfo_width()
        dh = self.dialog.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        self.dialog.geometry(f"+{x}+{y}")

    # ==================== 静态工厂 ====================

    @staticmethod
    def show(parent: tk.Tk,
             test_results: Optional[dict] = None,
             on_generate_report: Optional[Callable] = None,
             width: int = 640,
             height: int = 780):
        """快速弹出测试结果窗口"""
        TestResultDialog(parent, test_results,
                         on_generate_report=on_generate_report,
                         width=width, height=height)
