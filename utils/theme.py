"""
PTS 测试平台 — 共享视觉主题

统一的调色板、字体、间距、DPI 缩放工具，供主界面与各弹窗模块引用。
"""

import ctypes
import os

# ==================== DPI 缩放 ====================

if os.name == 'nt':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        _DPI = ctypes.windll.user32.GetDpiForSystem()
        _SCALE = _DPI / 96.0
    except Exception:
        _SCALE = 1.0
else:
    _SCALE = 1.0


def dpix(base_px: int) -> int:
    """将 96 DPI（100% 缩放）基准像素换算为当前屏幕的实际像素。

    用法：所有硬编码的窗口尺寸、控件宽度等都应以 100% 缩放下的值为基准，
         通过此函数转换为实际像素，保证在不同 DPI 的电脑上比例一致。
    """
    return int(round(base_px * _SCALE))


# ==================== 调色板 ====================

# 基础色
COLOR_BG               = "#F5F7FA"   # 页面背景（浅灰）
COLOR_WHITE            = "#FFFFFF"   # 白色

# 卡片 / 容器
COLOR_CARD_BG          = "#FFFFFF"   # 卡片背景
COLOR_CARD_BORDER      = "#FFFFFF"   # 卡片边框

# 分区标题
COLOR_SECTION_BG       = "#EDF0F5"   # 分区标题背景
COLOR_SECTION_FG       = "#4A5568"   # 分区标题文字

# 文字色阶
COLOR_TEXT_PRIMARY     = "#1A202C"   # 主文字
COLOR_TEXT_SECONDARY   = "#5A6577"   # 次级文字（标签）
COLOR_TEXT_MUTED       = "#A0AEC0"   # 弱化文字（占位/说明）

# 输入框
COLOR_ENTRY_BORDER     = "#D2D6DC"   # 输入框边框
COLOR_ENTRY_BG         = "#FFFFFF"   # 输入框背景

# 强调 / 功能色
COLOR_ACCENT           = "#2B6FF2"   # 主强调色（蓝）
COLOR_ACCENT_HOVER     = "#1E5CD6"   # 悬停态
COLOR_SUCCESS          = "#38A169"   # 成功（绿）
COLOR_WARNING          = "#F2994A"   # 警告/次要强调（橙）

# 日志状态色
COLOR_LOG_ERROR        = "#E53E3E"   # 错误
COLOR_LOG_COMPLETED    = "#38A169"   # 完成
COLOR_LOG_RUNNING      = "#2B6FF2"   # 运行中
COLOR_LOG_CLICKABLE    = "#2B6FF2"   # 可点击链接

# ==================== 字体 ====================

FONT_FAMILY            = "微软雅黑"
FONT_SIZE_TITLE        = 17
FONT_SIZE_HEADING      = 14
FONT_SIZE_BODY         = 10
FONT_SIZE_SMALL        = 9
FONT_SIZE_CAPTION      = 8

# ==================== 间距 ====================

PADDING_SECTION        = 24
PADDING_CARD           = 14
PADDING_ROW            = 5
