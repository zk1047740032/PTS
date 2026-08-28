"""种子激光器参数查询工具
通过 RS-485 串口直连，点击查询按钮单次查询并显示参数。
"""
import functools
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

import serial
import serial.tools.list_ports

import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from core.config import CFG
from utils.theme import (
    COLOR_BG, COLOR_WHITE, COLOR_CARD_BG, COLOR_CARD_BORDER,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    COLOR_ENTRY_BORDER, COLOR_ENTRY_BG,
    COLOR_ACCENT, COLOR_ACCENT_HOVER, COLOR_SUCCESS,
    COLOR_LOG_ERROR,
    FONT_FAMILY, FONT_SIZE_HEADING, FONT_SIZE_BODY, FONT_SIZE_SMALL, FONT_SIZE_CAPTION,
    PADDING_SECTION,
)

# ===============  DFB 种子激光器 RS-485 串口控制  ===============
class DFBLaserController:
    """协议帧格式: 0x50 | 0x00 | ADDR | 命令 | 数据长度 | [数据] | 和校验 | 异或校验 | 0x0D | 0x0A"""

    CMD_REALTIME = 0xA9   # 实时查询 → 响应 0xB7
    CMD_SYSTEM_INFO = 0xAA  # 系统信息查询 → 响应 0xB0

    def __init__(self, port, addr=None, log_func=print):
        self.port = port
        self.addr = addr if addr is not None else CFG.serial.device_addr
        self.log = log_func
        self.serial = None
        self._lock = threading.Lock()

    def open(self, timeout_s=None):
        if timeout_s is None:
            timeout_s = CFG.serial.timeout_s
        self.serial = serial.Serial(
            port=self.port, baudrate=CFG.serial.baudrate, timeout=timeout_s,
            bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
        )
        self.log(f"串口 {self.port} 已连接")

    def close(self):
        if self.serial and self.serial.is_open:
            self.serial.close()

    @staticmethod
    def _checksum(data: bytes):
        return sum(data) & 0xFF, functools.reduce(lambda a, b: a ^ b, data, 0)

    def _send_frame(self, cmd_byte: int, timeout: float | None = None) -> bytes | None:
        frame = bytes([0x50, 0x00, self.addr, cmd_byte, 0])
        s, x = self._checksum(frame[1:])
        frame += bytes([s, x, 0x0D, 0x0A])

        with self._lock:
            if timeout is not None:
                orig = self.serial.timeout
                self.serial.timeout = timeout
            try:
                self.serial.reset_input_buffer()
                self.serial.write(frame)
                header = self.serial.read(5)
                if len(header) < 5 or header[0] != 0x50:
                    return None
                data_len = header[4]
                rest = self.serial.read(data_len + 4)
                if len(rest) < data_len + 4:
                    return None
                full = header + rest
                recv_s, recv_x = self._checksum(full[1:-4])
                if recv_s != full[-4] or recv_x != full[-3]:
                    return None
                return full
            finally:
                if timeout is not None:
                    self.serial.timeout = orig

    def query_realtime(self, timeout: float | None = None) -> dict | None:
        raw = self._send_frame(self.CMD_REALTIME, timeout=timeout)
        if raw is None or len(raw) < 49:
            return None
        data = raw[5:-4]
        if len(data) < 40:
            return None

        def u2(offset, signed=False):
            return int.from_bytes(data[offset:offset + 2], 'big', signed=signed)

        def u4(offset, signed=False):
            return int.from_bytes(data[offset:offset + 4], 'big', signed=signed)

        return {
            'diode_set_temp':     u2(0, signed=True)   / 1000.0,
            'grating_set_temp':   u2(2)                 / 1000.0,
            'set_current':        u2(4),
            'power_set':          u2(8)                 / 100.0,
            'power_mode':         data[11],
            'current_switch':     data[13],
            'diode_temp':         u2(15, signed=True)  / 1000.0,
            'grating_temp':       u2(18)                / 1000.0,
            'case_temp':          u2(21, signed=True)  / 1000.0,
            'actual_current':     u2(23),
            'wavelength':         u4(30)                / 10000.0,
            'power_pd':           u2(36)                / 100.0,
            'diode_temp2':        u2(38, signed=True)  / 1000.0,
        }

    def query_system_info(self, timeout: float | None = None) -> dict | None:
        """查询系统信息（光栅温度/电流上下限等），命令 0xAA → 响应 0xB0"""
        raw = self._send_frame(self.CMD_SYSTEM_INFO, timeout=timeout)
        if raw is None or len(raw) < 133:
            return None
        data = raw[5:-4]  # 124 字节数据
        if len(data) < 104:
            return None

        def u2(offset):
            return int.from_bytes(data[offset:offset + 2], 'big', signed=False)

        return {
            'grating_temp_upper':  u2(95) / 1000.0,   # 协议偏移 100-101
            'grating_temp_lower':  u2(97) / 1000.0,   # 协议偏移 102-103
            'current_upper':       u2(100),            # 协议偏移 105-106
            'current_lower':       u2(102),            # 协议偏移 107-108
        }

# ===============  参数显示字段定义  ===============
SYS_INFO_LABELS = [
    ('grating_temp_upper',  '光栅温度上限',  ' °C'),
    ('grating_temp_lower',  '光栅温度下限',  ' °C'),
    ('current_upper',       '电流上限',      ' mA'),
    ('current_lower',       '电流下限',      ' mA'),
]

REALTIME_LABELS = [
    ('wavelength',        '波长',              ' nm'),
    ('grating_set_temp',  '光栅设置温度',       ' °C'),
    ('grating_temp',      '当前光栅温度',       ' °C'),
    ('set_current',       '设置电流',           ' mA'),
    ('actual_current',    '当前电流',           ' mA'),
]

# 交替行背景色
_ROW_BG_EVEN = COLOR_CARD_BG   # #FFFFFF
_ROW_BG_ODD  = "#F8F9FB"       # 极浅灰蓝


class SeedParamsPanel:
    """种子参数查询面板 — 嵌入到 Toplevel 弹窗中使用"""

    def __init__(self, parent):
        """
        参数:
            parent: 父容器（通常是 Toplevel 或 Frame）
        """
        self.parent = parent
        self._querying = False

        self._build_ui()

    # ---------- UI 构建 ----------
    def _build_ui(self):
        # 主容器
        main = tk.Frame(self.parent, bg=COLOR_BG)
        main.pack(fill=tk.BOTH, expand=True)

        # ---- 标题栏 ----
        header = tk.Frame(main, bg=COLOR_BG)
        header.pack(fill=tk.X, padx=PADDING_SECTION, pady=(16, 0))

        accent = tk.Frame(header, bg=COLOR_ACCENT, width=4, height=22)
        accent.pack(side=tk.LEFT, padx=(0, 10))
        accent.pack_propagate(False)

        tk.Label(header, text="种子参数",
                 font=(FONT_FAMILY, FONT_SIZE_HEADING, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_BG).pack(side=tk.LEFT)

        tk.Label(header, text="串口查询 DFB 种子激光器参数",
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_TEXT_MUTED, bg=COLOR_BG).pack(side=tk.LEFT, padx=(14, 0), pady=(6, 0))

        # ==== 卡片 1：连接设置 ====
        conn_card = tk.Frame(main, bg=COLOR_CARD_BG,
                             highlightbackground=COLOR_ENTRY_BORDER,
                             highlightthickness=1, bd=0)
        conn_card.pack(fill=tk.X, padx=PADDING_SECTION, pady=(12, 0))

        self._build_card_header(conn_card, "连接设置")

        inner1 = tk.Frame(conn_card, bg=COLOR_CARD_BG)
        inner1.pack(fill=tk.X, padx=10, pady=(2, 8))

        # 串口号行
        row1 = tk.Frame(inner1, bg=_ROW_BG_EVEN)
        row1.pack(fill=tk.X, ipady=1)
        tk.Label(row1, text="串口号", font=(FONT_FAMILY, FONT_SIZE_BODY),
                 fg=COLOR_TEXT_SECONDARY, bg=_ROW_BG_EVEN,
                 width=12, anchor="e").pack(side=tk.LEFT, padx=(10, 8), pady=4)

        self.com_port = ttk.Combobox(row1, width=8, values=[f'COM{i}' for i in range(1, 33)],
                                      font=(FONT_FAMILY, FONT_SIZE_BODY))
        self.com_port.set('COM3')
        self.com_port.pack(side=tk.LEFT)

        self.btn_refresh = tk.Button(row1, text='↻', font=(FONT_FAMILY, FONT_SIZE_BODY),
                                      bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
                                      activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
                                      relief="solid", bd=1, padx=6,
                                      command=self._scan_ports, cursor="hand2")
        self.btn_refresh.pack(side=tk.LEFT, padx=(4, 0))

        # 设备地址行
        row2 = tk.Frame(inner1, bg=_ROW_BG_ODD)
        row2.pack(fill=tk.X, ipady=1)
        tk.Label(row2, text="设备地址", font=(FONT_FAMILY, FONT_SIZE_BODY),
                 fg=COLOR_TEXT_SECONDARY, bg=_ROW_BG_ODD,
                 width=12, anchor="e").pack(side=tk.LEFT, padx=(10, 8), pady=4)

        self.addr_var = tk.StringVar(value=str(CFG.serial.device_addr))
        addr_border = tk.Frame(row2, bg=COLOR_ENTRY_BORDER, height=28)
        addr_border.pack(side=tk.LEFT, pady=3)
        tk.Entry(addr_border, textvariable=self.addr_var, width=6,
                 font=(FONT_FAMILY, FONT_SIZE_BODY),
                 bg=COLOR_ENTRY_BG, fg=COLOR_TEXT_PRIMARY,
                 relief="flat", bd=0, highlightthickness=0,
                 insertbackground=COLOR_ACCENT).pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # 查询按钮 + 状态行（状态左，按钮右）
        btn_row = tk.Frame(inner1, bg=COLOR_CARD_BG)
        btn_row.pack(fill=tk.X, pady=(6, 0))

        # 左侧：状态 + 错误信息
        status_left = tk.Frame(btn_row, bg=COLOR_CARD_BG)
        status_left.pack(side=tk.LEFT, padx=(10, 0))

        self.status_label = tk.Label(status_left, text='● 就绪',
                                      font=(FONT_FAMILY, FONT_SIZE_SMALL),
                                      fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_BG)
        self.status_label.pack(side=tk.LEFT)

        self.error_label = tk.Label(status_left, text='',
                                     font=(FONT_FAMILY, FONT_SIZE_SMALL),
                                     fg=COLOR_LOG_ERROR, bg=COLOR_CARD_BG)
        self.error_label.pack(side=tk.LEFT, padx=(8, 0))

        # 右侧：查询按钮
        self.btn_query = tk.Button(btn_row, text='查询',
                                    font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                                    bg=COLOR_ACCENT, fg=COLOR_WHITE,
                                    activebackground=COLOR_ACCENT_HOVER,
                                    activeforeground=COLOR_WHITE,
                                    relief="flat", bd=0, padx=16, pady=4,
                                    command=self._on_query, cursor="hand2")
        self.btn_query.pack(side=tk.RIGHT, padx=(0, 10))

        # ==== 卡片 2：系统信息 ====
        sys_card = tk.Frame(main, bg=COLOR_CARD_BG,
                            highlightbackground=COLOR_ENTRY_BORDER,
                            highlightthickness=1, bd=0)
        sys_card.pack(fill=tk.X, padx=PADDING_SECTION, pady=(8, 0))

        self._build_card_header(sys_card, "系统信息")

        inner2 = tk.Frame(sys_card, bg=COLOR_CARD_BG)
        inner2.pack(fill=tk.X, padx=10, pady=(2, 8))

        self.value_vars = {}
        for i, (key, label, unit) in enumerate(SYS_INFO_LABELS):
            row_bg = _ROW_BG_EVEN if i % 2 == 0 else _ROW_BG_ODD
            self._create_field_row(inner2, key, label, unit, row_bg)

        # ==== 卡片 3：实时参数 ====
        rt_card = tk.Frame(main, bg=COLOR_CARD_BG,
                           highlightbackground=COLOR_ENTRY_BORDER,
                           highlightthickness=1, bd=0)
        rt_card.pack(fill=tk.BOTH, expand=True, padx=PADDING_SECTION, pady=(8, 0))

        self._build_card_header(rt_card, "实时参数")

        inner3 = tk.Frame(rt_card, bg=COLOR_CARD_BG)
        inner3.pack(fill=tk.X, padx=10, pady=(2, 8))

        for i, (key, label, unit) in enumerate(REALTIME_LABELS):
            row_bg = _ROW_BG_EVEN if i % 2 == 0 else _ROW_BG_ODD
            self._create_field_row(inner3, key, label, unit, row_bg)

        # ---- 底部操作栏 ----
        bottom = tk.Frame(main, bg=COLOR_BG)
        bottom.pack(fill=tk.X, padx=PADDING_SECTION, pady=(8, 14))

        self.update_time_var = tk.StringVar(value='上次查询: --')
        tk.Label(bottom, textvariable=self.update_time_var,
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_TEXT_MUTED, bg=COLOR_BG).pack(side=tk.LEFT, pady=(4, 0))

        self._record_hint = tk.StringVar(value='')
        tk.Label(bottom, textvariable=self._record_hint,
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_SUCCESS, bg=COLOR_BG).pack(side=tk.LEFT, padx=(10, 0), pady=(4, 0))

        close_btn = tk.Button(bottom, text="关闭",
                               font=(FONT_FAMILY, FONT_SIZE_BODY),
                               bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
                               activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
                               relief="solid", bd=1, padx=16, pady=5,
                               command=self.parent.destroy, cursor="hand2")
        close_btn.pack(side=tk.RIGHT, padx=(6, 0))

        record_btn = tk.Button(bottom, text="记录",
                                font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                                bg=COLOR_ACCENT, fg=COLOR_WHITE,
                                activebackground=COLOR_ACCENT_HOVER,
                                activeforeground=COLOR_WHITE,
                                relief="flat", bd=0, padx=16, pady=5,
                                command=self._do_record, cursor="hand2")
        record_btn.pack(side=tk.RIGHT, padx=6)

    def _build_card_header(self, card: tk.Frame, title: str):
        """构建卡片标题栏：左侧强调条 + 标题 + 细分隔线"""
        title_bar = tk.Frame(card, bg=COLOR_CARD_BG)
        title_bar.pack(fill=tk.X)

        accent = tk.Frame(title_bar, bg=COLOR_ACCENT, width=3, height=16)
        accent.pack(side=tk.LEFT, padx=(14, 8), pady=(10, 0))
        accent.pack_propagate(False)

        tk.Label(title_bar, text=title,
                 font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD_BG).pack(side=tk.LEFT, pady=(10, 0))

        sep = tk.Frame(card, bg=COLOR_ENTRY_BORDER, height=1)
        sep.pack(fill=tk.X, padx=14, pady=(6, 2))

    def _create_field_row(self, parent: tk.Frame, key: str, label: str,
                          unit: str, row_bg: str):
        """创建单行字段：标签 + 值 + 单位"""
        row = tk.Frame(parent, bg=row_bg)
        row.pack(fill=tk.X, ipady=1)

        tk.Label(row, text=label + ':', font=(FONT_FAMILY, FONT_SIZE_BODY),
                 fg=COLOR_TEXT_SECONDARY, bg=row_bg,
                 width=14, anchor="e").pack(side=tk.LEFT, padx=(10, 8), pady=4)

        var = tk.StringVar(value='--')
        self.value_vars[key] = var
        tk.Label(row, textvariable=var,
                 font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=row_bg,
                 width=14, anchor="e").pack(side=tk.LEFT, pady=4)

        tk.Label(row, text=unit,
                 font=(FONT_FAMILY, FONT_SIZE_SMALL),
                 fg=COLOR_TEXT_MUTED, bg=row_bg).pack(side=tk.LEFT, padx=(4, 0), pady=4)

    # ---------- 串口扫描 ----------
    def _scan_ports(self):
        """扫描系统可用串口并更新下拉列表"""
        try:
            ports = [p.device for p in serial.tools.list_ports.comports()]
            if not ports:
                ports = [f'COM{i}' for i in range(1, 33)]
        except Exception:
            ports = [f'COM{i}' for i in range(1, 33)]

        self.com_port['values'] = ports
        # 如果当前选中的端口不在列表中，尝试保留或选第一个
        current = self.com_port.get()
        if ports and current not in ports:
            self.com_port.set(ports[0])

    # ---------- 查询逻辑 ----------
    def _on_query(self):
        """点击查询按钮"""
        if self._querying:
            return

        port = self.com_port.get().strip()
        try:
            addr = int(self.addr_var.get())
        except ValueError:
            messagebox.showerror('错误', '地址必须为整数')
            return

        self._querying = True
        self.btn_query.config(state='disabled', text='查询中...')
        self.status_label.config(text='● 查询中...', fg='#F2994A')
        self.error_label.config(text='')

        completed = threading.Event()  # 防止超时回调与工作线程竞态

        def do():
            try:
                # 打开串口 → 系统信息 → 实时参数 → 关闭串口（各环节超时收紧，合计约 2.5s）
                laser = DFBLaserController(port, addr, log_func=lambda msg: None)
                laser.open(timeout_s=0.5)
                sys_info = laser.query_system_info(timeout=1.0)
                realtime = laser.query_realtime(timeout=1.0)
                laser.close()

                if not completed.is_set():
                    self.parent.after(0, self._on_query_success, sys_info, realtime)
            except Exception as e:
                if not completed.is_set():
                    self.parent.after(0, self._on_query_failed, str(e))
            finally:
                if not completed.is_set():
                    completed.set()
                    self.parent.after(0, self._on_query_done)

        def run_with_timeout():
            """看门狗线程：join 工作线程最多 3 秒，超时则强制报失败"""
            worker = threading.Thread(target=do)
            worker.start()
            worker.join(timeout=3.0)
            if not completed.is_set():
                # 3 秒已过，工作线程仍卡在串口操作中 → 强制失败
                completed.set()
                self.parent.after(0, self._on_query_failed, '查询超时（3秒）')
                self.parent.after(0, self._on_query_done)

        threading.Thread(target=run_with_timeout, daemon=True).start()

    def _on_query_success(self, sys_info: dict | None, realtime: dict | None):
        """查询成功，更新系统信息和实时参数"""
        # 综合状态判断
        if sys_info is not None and realtime is not None:
            self.status_label.config(text='● 已连接（查询完成）', fg=COLOR_SUCCESS)
            self.error_label.config(text='')
        elif sys_info is not None:
            self.status_label.config(text='● 部分成功', fg='#F2994A')
            self.error_label.config(text='实时查询无响应')
        else:
            self.status_label.config(text='● 部分成功', fg='#F2994A')
            self.error_label.config(text='系统信息查询无响应')

        # 更新系统信息字段
        for key, _, _ in SYS_INFO_LABELS:
            v = sys_info.get(key) if sys_info else None
            if v is None:
                self.value_vars[key].set('--')
            elif key in ('grating_temp_upper', 'grating_temp_lower'):
                self.value_vars[key].set(f'{v:.3f}')
            elif key in ('current_upper', 'current_lower'):
                self.value_vars[key].set(f'{v:.1f}')
            else:
                self.value_vars[key].set(str(v))

        # 更新实时参数字段
        for key, _, _ in REALTIME_LABELS:
            v = realtime.get(key) if realtime else None
            if v is None:
                self.value_vars[key].set('--')
            elif key == 'wavelength':
                self.value_vars[key].set(f'{v:.4f}')
            elif key in ('grating_set_temp', 'grating_temp'):
                self.value_vars[key].set(f'{v:.3f}')
            elif key in ('set_current', 'actual_current'):
                self.value_vars[key].set(f'{v:.1f}')
            else:
                self.value_vars[key].set(str(v))

        self.update_time_var.set(f'上次查询: {datetime.now():%H:%M:%S}')

    def _on_query_failed(self, err: str):
        """查询失败"""
        self.status_label.config(text='● 查询失败', fg=COLOR_LOG_ERROR)
        self.error_label.config(text=err[:40])

        for key, _, _ in SYS_INFO_LABELS:
            self.value_vars[key].set('--')
        for key, _, _ in REALTIME_LABELS:
            self.value_vars[key].set('--')

    def _on_query_done(self):
        """查询结束，恢复按钮状态"""
        self._querying = False
        self.btn_query.config(state='normal', text='查询')

    # ---------- 记录按钮 ----------
    def _do_record(self):
        """记录当前参数到 CSV 并显示提示"""
        self.record_to_csv()
        self._record_hint.set('记录成功')
        self.parent.after(3000, lambda: self._record_hint.set(''))

    # ---------- 记录到 CSV ----------
    def record_to_csv(self):
        """将当前显示参数覆写写入 CSV 文件（仅一行数据）"""
        import csv
        import os

        save_dir = str(CFG.dirs.seed_value)
        os.makedirs(save_dir, exist_ok=True)
        filepath = os.path.join(save_dir, 'seedvalue.csv')

        fields = SYS_INFO_LABELS + REALTIME_LABELS
        row = {}
        for key, label, _ in fields:
            row[label] = self.value_vars.get(key, tk.StringVar(value='--')).get()

        with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=list(row.keys()))
            writer.writeheader()
            writer.writerow(row)
