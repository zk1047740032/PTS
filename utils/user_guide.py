"""
频准测试系统 — 帮助文档渲染

从 Markdown 文件渲染到 tk.Text 控件。
支持: ##/### 标题、**粗体**、列表层级、图片嵌入。
"""

import os
import re
import sys
import tkinter as tk
from PIL import Image, ImageTk

from utils.theme import (
    COLOR_BG, COLOR_CARD_BG,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    COLOR_ACCENT, COLOR_ENTRY_BORDER,
    FONT_FAMILY, FONT_SIZE_HEADING, FONT_SIZE_BODY, FONT_SIZE_SMALL, FONT_SIZE_CAPTION,
)

# ==================== 路径 ====================

# PyInstaller 打包后 __file__ 指向临时目录，需用 sys._MEIPASS
if getattr(sys, "frozen", False):
    _PROJECT_DIR = sys._MEIPASS
else:
    _PROJECT_DIR = os.path.dirname(os.path.dirname(__file__))
HELP_IMAGES_DIR = os.path.join(_PROJECT_DIR, "help_images")

# ==================== 正则 ====================

_IMG_RE   = re.compile(r'^\s*!\[.*?\]\((.+?)\)\s*$')   # ![alt](path)
_H2_RE    = re.compile(r'^## (.+)$')                      # ## 标题
_H3_RE    = re.compile(r'^### (.+)$')                     # ### 子标题
_BULLET_RE = re.compile(r'^(\s*)- (.+)$')                # - /   - 列表
_BOLD_RE  = re.compile(r'\*\*(.+?)\*\*')                  # **粗体**

# ==================== 主渲染函数 ====================


def render_markdown(text_widget, md_filename):
    """解析 Markdown 文件并渲染到 tk.Text 控件。

    Args:
        text_widget: tk.Text 控件实例
        md_filename: Markdown 文件名（相对于项目根目录）
    """
    md_path = os.path.join(_PROJECT_DIR, md_filename)
    if not os.path.exists(md_path):
        text_widget.config(state="normal")
        text_widget.delete("1.0", "end")
        text_widget.insert("end", f"[文档未找到: {md_path}]", "body")
        text_widget.config(state="disabled")
        return

    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    _setup_tags(text_widget)
    _setup_image_holder(text_widget)
    _setup_pixel_scroll(text_widget)

    # 强制布局计算，确保后续能拿到正确的控件宽度
    text_widget.update_idletasks()

    text_widget.config(state="normal")
    text_widget.delete("1.0", "end")

    for line in lines:
        stripped = line.rstrip()

        # 空行
        if not stripped:
            text_widget.insert("end", "\n", "body")
            continue

        # 图片
        m = _IMG_RE.match(stripped)
        if m:
            raw_path = m.group(1).strip().replace("\\", "/")
            if raw_path.startswith("help_images/"):
                img_abs = os.path.join(HELP_IMAGES_DIR, raw_path.split("help_images/", 1)[1])
            else:
                img_abs = os.path.join(_PROJECT_DIR, raw_path)
            _render_md_image(text_widget, img_abs)
            continue

        # h2 标题
        m = _H2_RE.match(stripped)
        if m:
            _insert_h2(text_widget, m.group(1))
            continue

        # h3 子标题
        m = _H3_RE.match(stripped)
        if m:
            _insert_h3(text_widget, m.group(1))
            continue

        # 列表项
        m = _BULLET_RE.match(stripped)
        if m:
            indent = len(m.group(1))
            level = indent // 2
            _insert_bullet(text_widget, level, m.group(2))
            continue

        # 普通正文
        _insert_rich_line(text_widget, stripped, "body")

    text_widget.config(state="disabled")


# ==================== 行内富文本 ====================


def _insert_rich_line(text_widget, text, base_tag):
    """插入一行文本，解析 **粗体** 并分段打标签。"""
    parts = _BOLD_RE.split(text)  # 交替: [普通, 粗体内容, 普通, ...]
    for i, part in enumerate(parts):
        if not part:
            continue
        tag = "bold" if i % 2 == 1 else base_tag
        text_widget.insert("end", part, (base_tag, tag))
    text_widget.insert("end", "\n", base_tag)


# ==================== 各元素渲染 ====================


def _insert_h2(text_widget, text):
    """渲染 h2 标题：蓝色竖线 + 蓝色加粗文字。"""
    # 段前间距
    text_widget.insert("end", "\n", "spacer")
    # 蓝色竖线
    text_widget.insert("end", "  │  ", "h2_bar")
    # 标题文字（带粗体解析）
    _insert_rich_line(text_widget, text, "h2")
    # 下方细线
    text_widget.insert("end", "  " + "─" * 52 + "\n", "h2_rule")


def _insert_h3(text_widget, text):
    """渲染 h3 子标题：深色加粗 + 上间距。"""
    text_widget.insert("end", "\n", "spacer")
    _insert_rich_line(text_widget, text, "h3")


def _insert_bullet(text_widget, level, text):
    """渲染列表项，根据层级选标记和标签。"""
    if level == 0:
        marker, tag = "●", "bullet_l1"
    elif level == 1:
        marker, tag = "◦", "bullet_l2"
    elif level == 2:
        marker, tag = "▪", "bullet_l3"
    else:
        marker, tag = "▸", "bullet_l4"

    # 标记与正文同在一行（无 \\n 分隔），保证换行时不断开
    text_widget.insert("end", f"{marker}  ", (tag, "bullet_marker"))
    _insert_rich_line(text_widget, text, tag)


# ==================== 图片渲染 ====================


def _forward_scroll(event):
    """像素级滚轮滚动（避免嵌入图片时按行滚动造成大幅跳位）。"""
    w = event.widget
    while w is not None:
        if isinstance(w, tk.Text):
            # 每 notch 滚动约 50px，相当于 ~3 行文字
            pixels = int(-1 * (event.delta / 120) * 50)
            w.yview_scroll(pixels, "pixels")
            return
        w = w.master


def _bind_scroll(widget):
    """为控件绑定滚轮事件转发到父级 Text。"""
    widget.bind("<MouseWheel>", _forward_scroll, add="+")


def _setup_pixel_scroll(text_widget):
    """覆盖 Text 控件自身的滚轮行为，统一使用像素级滚动。"""
    if hasattr(text_widget, "_pixel_scroll_ready"):
        return
    text_widget._pixel_scroll_ready = True

    def _on_text_scroll(event):
        pixels = int(-1 * (event.delta / 120) * 50)
        text_widget.yview_scroll(pixels, "pixels")
        return "break"  # 阻止默认的按行滚动

    # 实例级绑定优先于类绑定，return "break" 阻止默认行为
    text_widget.bind("<MouseWheel>", _on_text_scroll)


def _render_md_image(text_widget, image_path):
    """在 tk.Text 中嵌入图片，带边框和自动标题（居中）。"""
    if not os.path.exists(image_path):
        fname = os.path.basename(image_path)
        text_widget.insert("end", f"  [图片未找到: {fname}]\n", "img_placeholder")
        return

    try:
        pil_img = Image.open(image_path)
        max_w = 1000
        if pil_img.width > max_w:
            ratio = max_w / pil_img.width
            h = int(pil_img.height * ratio)
            pil_img = pil_img.resize((max_w, h), Image.LANCZOS)

        photo = ImageTk.PhotoImage(pil_img)
        text_widget._md_images.append(photo)

        img_w = pil_img.width
        img_h = pil_img.height

        # 内容区宽度：优先拿真实宽度，回落 reqwidth，再回落估算
        text_w = text_widget.winfo_width()
        if text_w < 10:
            text_w = text_widget.winfo_reqwidth()
        if text_w < 10:
            from utils.theme import dpix
            text_w = dpix(680) - dpix(100)
        full_w = max(text_w - 32, img_w)

        # ---- 外层容器：撑满内容区宽度，防止被挤压 ----
        outer = tk.Frame(text_widget, bg=COLOR_CARD_BG, width=full_w, height=img_h + 6)
        outer.pack_propagate(False)

        # ---- 图片 + 边框居中放置 ----
        border = tk.Frame(outer, bg=COLOR_ENTRY_BORDER, bd=0)
        inner = tk.Frame(border, bg=COLOR_CARD_BG, bd=0)
        inner.pack(padx=1, pady=1)

        lbl = tk.Label(inner, image=photo, bg=COLOR_CARD_BG, bd=0)
        lbl.image = photo
        lbl.pack()

        border.place(relx=0.5, rely=0.5, anchor="center")

        # 滚轮事件转发给 Text 控件
        for w in (outer, border, inner, lbl):
            _bind_scroll(w)

        text_widget.insert("end", "\n")
        text_widget.window_create("end", window=outer)
        text_widget.insert("end", "\n")

        # ---- 标题（文件名去扩展名），用全宽 Label 居中 ----
        caption = os.path.splitext(os.path.basename(image_path))[0]
        cap_frame = tk.Frame(text_widget, bg=COLOR_CARD_BG, width=full_w, height=30)
        cap_frame.pack_propagate(False)
        cap_lbl = tk.Label(cap_frame, text=f"▲ {caption}",
                           font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                           fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_BG)
        cap_lbl.place(relx=0.5, rely=0.5, anchor="center")

        for w in (cap_frame, cap_lbl):
            _bind_scroll(w)

        text_widget.window_create("end", window=cap_frame)
        text_widget.insert("end", "\n")

    except Exception as e:
        fname = os.path.basename(image_path)
        text_widget.insert("end", f"  [图片加载失败: {fname} — {e}]\n", "img_placeholder")


# ==================== 标签配置 ====================


def _setup_tags(text_widget):
    """配置所有 tk.Text 样式标签。"""
    if hasattr(text_widget, "_md_tags_ready"):
        return
    text_widget._md_tags_ready = True

    # h2 标题
    text_widget.tag_configure("h2",
        font=(FONT_FAMILY, FONT_SIZE_HEADING + 1, "bold"),
        foreground=COLOR_ACCENT,
        spacing1=0, spacing3=0,
    )
    # h2 蓝色竖线
    text_widget.tag_configure("h2_bar",
        font=(FONT_FAMILY, FONT_SIZE_HEADING + 1, "bold"),
        foreground=COLOR_ACCENT,
    )
    # h2 下方细线
    text_widget.tag_configure("h2_rule",
        font=(FONT_FAMILY, FONT_SIZE_CAPTION),
        foreground=COLOR_ENTRY_BORDER,
        spacing1=0, spacing3=8,
    )

    # h3 子标题
    text_widget.tag_configure("h3",
        font=(FONT_FAMILY, FONT_SIZE_BODY + 1, "bold"),
        foreground=COLOR_TEXT_PRIMARY,
        spacing1=12, spacing3=2,
    )

    # 正文
    text_widget.tag_configure("body",
        font=(FONT_FAMILY, FONT_SIZE_BODY),
        foreground=COLOR_TEXT_PRIMARY,
        lmargin1=4, lmargin2=4,
        spacing1=1, spacing3=1,
    )

    # 行内粗体
    text_widget.tag_configure("bold",
        font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
        foreground=COLOR_TEXT_PRIMARY,
    )

    # 间隔
    text_widget.tag_configure("spacer",
        font=(FONT_FAMILY, 2),
    )

    # ---- 列表（lmargin1=lmargin2，等宽缩进，换行不跳位） ----
    text_widget.tag_configure("bullet_l1",
        font=(FONT_FAMILY, FONT_SIZE_BODY),
        foreground=COLOR_TEXT_PRIMARY,
        lmargin1=24, lmargin2=24,
        spacing1=1, spacing3=1,
    )
    text_widget.tag_configure("bullet_l2",
        font=(FONT_FAMILY, FONT_SIZE_BODY),
        foreground=COLOR_TEXT_SECONDARY,
        lmargin1=44, lmargin2=44,
        spacing1=1, spacing3=1,
    )
    text_widget.tag_configure("bullet_l3",
        font=(FONT_FAMILY, FONT_SIZE_BODY),
        foreground=COLOR_TEXT_SECONDARY,
        lmargin1=64, lmargin2=64,
        spacing1=1, spacing3=1,
    )
    text_widget.tag_configure("bullet_l4",
        font=(FONT_FAMILY, FONT_SIZE_BODY),
        foreground=COLOR_TEXT_MUTED,
        lmargin1=80, lmargin2=80,
        spacing1=1, spacing3=1,
    )
    # 列表标记（强调色）
    text_widget.tag_configure("bullet_marker",
        foreground=COLOR_ACCENT,
    )

    # ---- 图片相关 ----
    text_widget.tag_configure("img_line",
        justify="center",
    )
    text_widget.tag_configure("img_caption",
        font=(FONT_FAMILY, FONT_SIZE_CAPTION),
        foreground=COLOR_TEXT_MUTED,
        justify="center",
        spacing1=0, spacing3=10,
    )
    text_widget.tag_configure("img_placeholder",
        font=(FONT_FAMILY, FONT_SIZE_CAPTION),
        foreground=COLOR_TEXT_MUTED,
        background="#EEF2FF",
        lmargin1=4, lmargin2=4,
        spacing1=4, spacing3=4,
        justify="center",
    )


def _setup_image_holder(text_widget):
    """创建图片引用列表（防 GC）。"""
    if not hasattr(text_widget, "_md_images"):
        text_widget._md_images = []


# ==================== 兼容旧版 ====================

USER_GUIDE = """    频准测试系统 (PTS)

    欢迎使用一体化测试系统。下面提供一些基础操作说明……

    如需进一步帮助，请联系开发人员（张珂）。
"""


def render_user_guide(text_widget):
    """已废弃 — 请使用 render_markdown(text_widget, "一键测试操作说明.md")"""
    render_markdown(text_widget, "一键测试操作说明.md")
