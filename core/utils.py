#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core.utils —— 公共工具函数

集中 path_a / path_b 各测试程序中重复出现的纯逻辑：
    - Windows DPI 自适应
    - VISA 资源地址拼装
    - 可中断睡眠
    - 清空输出目录
    - CSV 单行/双列写入
    - SCPI 截图读回

所有函数无状态、可独立测试，不依赖 tkinter。
"""

from __future__ import annotations

import csv
import os
import shutil
import threading
import time
import ctypes
from typing import Optional, Callable, Iterable, Any

__all__ = [
    "setup_dpi_awareness",
    "visa_address",
    "interruptible_sleep",
    "clear_directory",
    "append_row_csv",
    "write_xy_csv",
    "read_instrument_screenshot",
]


def setup_dpi_awareness() -> float:
    """
    启用 Windows 进程 DPI 感知，解决高 DPI 屏幕下界面模糊问题。

    替代各文件顶部重复的 ctypes.windll.shcore ... 块。
    在 core/__init__.py 导入时调用一次即可，重复调用会被异常吞掉（见下）。

    Returns:
        float: 系统缩放因子（DPI/96.0）。非 Windows 或失败时返回 1.0。
    """
    if os.name != 'nt':
        return 1.0
    try:
        # 设置进程为 "系统 DPI 感知" 级别
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        dpi = ctypes.windll.user32.GetDpiForSystem()
        return dpi / 96.0
    except Exception:
        # 已设置过 DPI 感知时，SetProcessDpiAwareness 会抛异常，此处吞掉
        return 1.0


def visa_address(host: str, *, kind: str = "tcpip_instr") -> str:
    """
    按统一规则构造 VISA 资源地址字符串。

        参数:
            host (str): 主机地址。对于 LAN 类为 IP；对 kind="usb" 为完整 USB 资源串，
                        原样透传。
            kind (str): 地址类型：
                - "tcpip_instr"  -> ``TCPIP::{ip}::INSTR``               (默认，多数频谱仪/信号源，VXI-11)
                - "inst0"        -> ``TCPIP::{ip}::inst0::INSTR``      (R&S FSV3004 等，VXI-11)
                - "socket5025"   -> ``TCPIP::{ip}::5025::SOCKET``      (LXI 原始套接字)
                - "usb"          -> 透传 host 原样作为 USB 资源串        (光开关/功率计等 USB 设备)

        返回:
            str: VISA 资源地址字符串。

        说明:
            收口现有 3 种硬编码地址写法（path_a/SingleFrequency.py 的 SOCKET、
            path_a/Rin_FSV3004.py 的 inst0、path_b/WaveLength.py 的 INSTR），
            避免在每个控制器里重新拼字符串。
    """
    if kind == "tcpip_instr":
        return f"TCPIP::{host}::INSTR"
    if kind == "inst0":
        return f"TCPIP::{host}::inst0::INSTR"
    if kind == "socket5025":
        return f"TCPIP::{host}::5025::SOCKET"
    if kind == "usb":
        return host  # 透传完整 USB 资源串
    raise ValueError(f"未知的 VISA 地址类型: {kind}")


def interruptible_sleep(
    seconds: float,
    stop_flag: threading.Event,
    check_interval: float = 0.5,
) -> bool:
    """
    可中断的睡眠：每隔 check_interval 秒检查一次 stop_flag。

    提取自 path_a/PhaseNoise.py 与 path_b/WaveLength.py 中各 GUI 的
    ``_interruptible_sleep``，将其改为纯函数（stop_flag 作为入参）。

        参数:
            seconds (float): 总睡眠时间（秒）。
            stop_flag (threading.Event): 停止标志事件。置位时立即中断。
            check_interval (float): 检查间隔（秒），默认 0.5。

        返回:
            bool: True 表示正常睡眠完成；False 表示被 stop_flag 中断。
    """
    elapsed = 0.0
    while elapsed < seconds:
        if stop_flag.is_set():
            return False
        sleep_time = min(check_interval, seconds - elapsed)
        time.sleep(sleep_time)
        elapsed += sleep_time
    return True


def clear_directory(dir_path: str, log_func: Optional[Callable[[str], None]] = None) -> None:
    """
    清空指定目录下的所有文件与子目录（保留目录本身）。

    提取自 path_a/LineWidth_FSV3004.py、Rin_FSV3004.py、SingleFrequency.py 中
    重复出现的"清空输出文件夹"循环。逐项 try/except，单文件删除失败不中断整体。

        参数:
            dir_path (str): 要清空的目录路径。不存在则静默返回。
            log_func (callable, optional): 日志回调，删除失败时逐项记录。
    """
    if not os.path.exists(dir_path):
        return
    for item in os.listdir(dir_path):
        fp = os.path.join(dir_path, item)
        try:
            if os.path.isfile(fp) or os.path.islink(fp):
                os.remove(fp)
            elif os.path.isdir(fp):
                shutil.rmtree(fp)
        except Exception as e:
            if log_func:
                log_func(f"[警告] 删除 {fp} 失败: {e}")


def append_row_csv(
    path: str,
    row: Iterable[Any],
    header: Optional[Iterable[Any]] = None,
) -> None:
    """
    向 CSV 文件追加一行；文件不存在且提供了 header 时先写表头。

    提取自 path_a/LineWidth_FSV3004.py 的 ``save_ndbdown_to_csv`` 等场景：
    首次写入带表头，后续只追加数据行。

        参数:
            path (str): CSV 文件路径。所在目录需已存在（由调用方保证）。
            row (iterable): 要追加的数据行。
            header (iterable, optional): 表头。仅当文件不存在时写入。
    """
    file_exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists and header is not None:
            writer.writerow(list(header))
        writer.writerow(list(row))


def write_xy_csv(path: str, xs: Iterable[float], ys: Iterable[float]) -> None:
    """
    将两列数据（频率/幅值）逐行写入 CSV，无表头。

    提取自 path_a/SingleFrequency.py、Rin_FSV3004.py、path_b/WaveLength.py 中
    ``for freq, amp in zip(...): writer.writerow([freq, amp])`` 的重复写法。

        参数:
            path (str): CSV 文件路径。所在目录需已存在。
            xs (iterable): 第一列数据（如频率）。
            ys (iterable): 第二列数据（如幅值）。
    """
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for x, y in zip(xs, ys):
            writer.writerow([x, y])


def read_instrument_screenshot(
    inst: Any,
    instr_temp_path: str,
    pc_save_path: str,
    log_func: Optional[Callable[[str], None]] = None,
    *,
    png_format: str = "PNG",
) -> str:
    """
    通过 SCPI 把仪器屏幕截图读回 PC 本地。

    封装目前散落在 path_a/LineWidth_FSV3004.py 与 Rin_FSV3004.py 的四步序列：
        1. 配置硬拷贝目标为 MMEM、格式为 png，指定仪器本地临时文件名
        2. 触发硬拷贝，等待 *OPC?
        3. ``MMEM:DATA?`` 二进制读回，写入 PC 本地路径
        4. ``MMEM:DEL`` 删除仪器本地临时文件

        参数:
            inst: 已连接的 pyvisa 仪器对象（需具备 write/query/query_binary_values）。
            instr_temp_path (str): 仪器本地临时 png 路径（如 'C:\\...\\_temp.png'）。
            pc_save_path (str): PC 本地保存路径（含文件名）。
            log_func (callable, optional): 日志回调。
            png_format (str): HCOPy 设备语言，默认 "PNG"。

        返回:
            str: PC 本地保存路径。

        说明:
            调用前需确保 pc_save_path 所在目录已存在。本函数不创建目录，
            以保持与现有代码一致的职责边界。
    """
    inst.write("HCOPy:DEST 'MMEM'")
    inst.write("HCOPy:FILE:NAME:AUTO:STATe OFF")
    inst.write(f"HCOPy:DEVice:LANGuage {png_format}")
    inst.write(f"MMEM:NAME '{instr_temp_path}'")
    inst.write("HCOPy:IMM")
    inst.query("*OPC?")

    png_data = inst.query_binary_values(
        f"MMEM:DATA? '{instr_temp_path}'", datatype="B", container=bytearray
    )
    with open(pc_save_path, "wb") as f:
        f.write(bytes(png_data))
    inst.write(f"MMEM:DEL '{instr_temp_path}'")

    if log_func:
        log_func(f"截图已保存到: {pc_save_path}")
    return pc_save_path
