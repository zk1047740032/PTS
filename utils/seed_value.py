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

# ===============  DFB 种子激光器 RS-485 串口控制  ===============
class DFBLaserController:
    """协议帧格式: 0x50 | 0x00 | ADDR | 命令 | 数据长度 | [数据] | 和校验 | 异或校验 | 0x0D | 0x0A"""

    CMD_REALTIME = 0xA9   # 实时查询 → 响应 0xB7
    CMD_SYSTEM_INFO = 0xAA  # 系统信息查询 → 响应 0xB0

    def __init__(self, port, addr=CFG.serial.device_addr, log_func=print):
        self.port = port
        self.addr = addr
        self.log = log_func
        self.serial = None
        self._lock = threading.Lock()

    def open(self, timeout_s=CFG.serial.timeout_s):
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
        # 顶部连接栏
        top = ttk.Frame(self.parent, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text='串口号:').pack(side=tk.LEFT)
        self.com_port = ttk.Combobox(top, width=8, values=[f'COM{i}' for i in range(1, 33)])
        self.com_port.set('COM3')
        self.com_port.pack(side=tk.LEFT, padx=4)

        # 刷新串口按钮
        self.btn_refresh = ttk.Button(top, text='↻', width=2, command=self._scan_ports)
        self.btn_refresh.pack(side=tk.LEFT)

        ttk.Label(top, text='地址:').pack(side=tk.LEFT, padx=(8, 0))
        self.addr_var = tk.StringVar(value='100')
        ttk.Entry(top, textvariable=self.addr_var, width=5).pack(side=tk.LEFT, padx=4)

        # 查询按钮
        self.btn_query = ttk.Button(top, text='查询', command=self._on_query, width=8)
        self.btn_query.pack(side=tk.LEFT, padx=(12, 0))

        self.status_label = ttk.Label(top, text='● 就绪', foreground='gray')
        self.status_label.pack(side=tk.LEFT, padx=12)

        self.error_label = ttk.Label(top, text='', foreground='red')
        self.error_label.pack(side=tk.LEFT, padx=4)

        # 分隔线
        ttk.Separator(self.parent, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=8)

        self.value_vars = {}

        # ==== 系统信息 ====
        sys_lf = ttk.LabelFrame(self.parent, text='系统信息', padding=8)
        sys_lf.pack(fill=tk.X, padx=8, pady=(6, 0))

        for row, (key, label, unit) in enumerate(SYS_INFO_LABELS):
            ttk.Label(sys_lf, text=label + ':', font=('', 10)).grid(
                row=row, column=0, sticky=tk.W, pady=4, padx=(0, 10))
            var = tk.StringVar(value='--')
            self.value_vars[key] = var
            ttk.Label(sys_lf, textvariable=var, font=('', 10, 'bold'),
                      width=14, anchor=tk.E).grid(row=row, column=1, sticky=tk.E, pady=4)
            ttk.Label(sys_lf, text=unit, font=('', 9), foreground='gray').grid(
                row=row, column=2, sticky=tk.W, pady=4, padx=(4, 0))

        # ==== 实时参数 ====
        rt_lf = ttk.LabelFrame(self.parent, text='实时参数', padding=8)
        rt_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        for row, (key, label, unit) in enumerate(REALTIME_LABELS):
            ttk.Label(rt_lf, text=label + ':', font=('', 10)).grid(
                row=row, column=0, sticky=tk.W, pady=4, padx=(0, 10))
            var = tk.StringVar(value='--')
            self.value_vars[key] = var
            ttk.Label(rt_lf, textvariable=var, font=('', 10, 'bold'),
                      width=14, anchor=tk.E).grid(row=row, column=1, sticky=tk.E, pady=4)
            ttk.Label(rt_lf, text=unit, font=('', 9), foreground='gray').grid(
                row=row, column=2, sticky=tk.W, pady=4, padx=(4, 0))

        # 底部状态
        ttk.Separator(self.parent, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=8)
        bottom = ttk.Frame(self.parent, padding=6)
        bottom.pack(fill=tk.X)
        self.update_time_var = tk.StringVar(value='上次查询: --')
        ttk.Label(bottom, textvariable=self.update_time_var, font=('', 8),
                  foreground='gray').pack(side=tk.LEFT)

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
        self.status_label.config(text='● 查询中...', foreground='orange')
        self.error_label.config(text='')

        def do():
            try:
                # 打开串口 → 系统信息 → 实时参数 → 关闭串口
                laser = DFBLaserController(port, addr, log_func=lambda msg: None)
                laser.open(timeout_s=1.0)
                sys_info = laser.query_system_info(timeout=1.5)
                realtime = laser.query_realtime(timeout=1.5)
                laser.close()

                self.parent.after(0, self._on_query_success, sys_info, realtime)
            except Exception as e:
                self.parent.after(0, self._on_query_failed, str(e))
            finally:
                self.parent.after(0, self._on_query_done)

        threading.Thread(target=do, daemon=True).start()

    def _on_query_success(self, sys_info: dict | None, realtime: dict | None):
        """查询成功，更新系统信息和实时参数"""
        # 综合状态判断
        if sys_info is not None and realtime is not None:
            self.status_label.config(text='● 已连接（查询完成）', foreground='green')
            self.error_label.config(text='')
        elif sys_info is not None:
            self.status_label.config(text='● 部分成功', foreground='orange')
            self.error_label.config(text='实时查询无响应')
        else:
            self.status_label.config(text='● 部分成功', foreground='orange')
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
        self.status_label.config(text='● 查询失败', foreground='red')
        self.error_label.config(text=err[:40])

        for key, _, _ in SYS_INFO_LABELS:
            self.value_vars[key].set('--')
        for key, _, _ in REALTIME_LABELS:
            self.value_vars[key].set('--')

    def _on_query_done(self):
        """查询结束，恢复按钮状态"""
        self._querying = False
        self.btn_query.config(state='normal', text='查询')

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
