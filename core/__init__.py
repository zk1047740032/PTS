#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core —— path_a / path_b 测试程序通用基类与工具

四层抽象：
    utils          公共工具函数（DPI、地址构造、清目录、CSV、截图读回）
    base_instrument    VisaInstrument        仪器控制器基类
    base_gui           BaseTestGUI           GUI 基类
    base_runner        BaseTestRunner        测试流程编排基类（模板方法）

import 本包即激活 Windows DPI 感知，替代各文件顶部重复的 ctypes 块。
"""

from .utils import (
    setup_dpi_awareness,
    visa_address,
    interruptible_sleep,
    clear_directory,
    append_row_csv,
    write_xy_csv,
    read_instrument_screenshot,
)
from .base_instrument import VisaInstrument
from .base_gui import BaseTestGUI
from .base_runner import BaseTestRunner

__all__ = [
    # utils
    "setup_dpi_awareness",
    "visa_address",
    "interruptible_sleep",
    "clear_directory",
    "append_row_csv",
    "write_xy_csv",
    "read_instrument_screenshot",
    # base classes
    "VisaInstrument",
    "BaseTestGUI",
    "BaseTestRunner",
]

# 导入即激活 DPI 感知（替代各测试文件顶部重复的 ctypes.windll.shcore 块）
setup_dpi_awareness()
