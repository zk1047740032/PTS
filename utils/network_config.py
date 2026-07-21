"""
网络配置工具 —— 一键添加仪器子网辅助 IP，支持连通性检测。

解决"换电脑后每次都要手动配静态 IP"的痛点：
    仪器 IP 都在同一子网（如 192.168.7.x），但用户电脑走 DHCP（如 10.x.x.x），
    不在同一子网 → 仪器 ping 不通。本工具通过 netsh 给网卡添加辅助 IP，
    不影响原有 DHCP 上网，一次配置永久生效。

权限说明：
    netsh 添加/删除 IP 需要管理员权限。GUI 中通过 ShellExecuteW + runas
    提权执行子进程。
"""

from __future__ import annotations

import ctypes
import ipaddress
import platform
import re
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional

from core.config import CFG
from utils.theme import (
    COLOR_BG, COLOR_WHITE, COLOR_CARD_BG, COLOR_CARD_BORDER,
    COLOR_SECTION_BG, COLOR_SECTION_FG,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    COLOR_ACCENT, COLOR_ACCENT_HOVER,
    COLOR_SUCCESS, COLOR_WARNING,
    COLOR_ENTRY_BORDER, COLOR_ENTRY_BG,
    FONT_FAMILY, FONT_SIZE_HEADING, FONT_SIZE_BODY,
    FONT_SIZE_SMALL, FONT_SIZE_CAPTION,
    PADDING_SECTION,
)


# ================================================================
# 系统操作（不依赖 GUI）
# ================================================================

def _run_netsh(args: str) -> tuple[int, str, str]:
    """执行 netsh 命令，返回 (returncode, stdout, stderr)。"""
    result = subprocess.run(
        f'netsh {args}',
        shell=True, capture_output=True, text=True
    )
    return result.returncode, result.stdout, result.stderr


def get_active_adapters() -> list[dict[str, str]]:
    """
    获取已连接的以太网 / Wi-Fi 适配器列表。

    Returns:
        [{"name": "以太网", "index": 5, "state": "已连接"}, ...]
    """
    adapters = []
    try:
        rc, out, _ = _run_netsh('interface ip show interfaces')
        if rc != 0:
            return adapters
        for line in out.splitlines():
            m = re.match(
                r'\s*(?P<idx>\d+)\s+\d+\s+(?P<state1>\S+)\s+(?P<state2>\S+)\s+(?P<name>.+)',
                line.strip()
            )
            if not m:
                continue
            state = m.group('state1')
            if state.lower() in ('connected', '已连接') or m.group('state2').lower() in ('connected', '已连接'):
                adapters.append({
                    "name": m.group('name').strip(),
                    "index": m.group('idx'),
                    "state": "已连接",
                })
    except Exception:
        pass
    return adapters


def get_current_ip_addresses() -> list[str]:
    """通过 ipconfig 获取本机所有 IPv4 地址。"""
    try:
        result = subprocess.run('ipconfig', shell=True,
                                capture_output=True, text=True)
        return re.findall(r'IPv4[^:]*:\s*(\d+\.\d+\.\d+\.\d+)', result.stdout)
    except Exception:
        return []


def _extract_instrument_ips() -> list[tuple[str, str]]:
    """从 CFG 中提取所有仪器 IP 地址，返回 (字段名, IP) 列表。

    同一 IP 可对应多个字段（如 FSV3004 同时用于 RIN 和线宽），
    不做去重，每字段一行。
    """
    result = []
    net = CFG.network
    for attr in dir(net):
        if attr.startswith('_'):
            continue
        val = getattr(net, attr, '')
        if isinstance(val, str) and re.match(r'\d+\.\d+\.\d+\.\d+', val):
            result.append((attr, val))
    return result


def _detect_instrument_subnet() -> Optional[str]:
    """
    从仪器 IP 中推断子网。

    取所有仪器 IP 的公共前缀（最常见的 /24 子网）。
    """
    items = _extract_instrument_ips()
    if not items:
        return None
    first_ip = items[0][1]
    prefix = first_ip.rsplit('.', 1)[0] + '.0'
    return f"{prefix}/24"


def _suggest_local_ip(subnet_cidr: str) -> Optional[str]:
    """
    根据子网推荐一个本机 IP，选 .100 作为默认。
    """
    try:
        net = ipaddress.IPv4Network(subnet_cidr, strict=False)
    except Exception:
        return None
    return str(net.network_address + 100)


def ping(ip: str, timeout_ms: int = 800) -> bool:
    """ping 目标 IP，返回是否可达。"""
    param = "-n 1 -w" if platform.system() == "Windows" else "-c 1 -W"
    timeout = str(timeout_ms // 1000) if platform.system() == "Windows" else str(timeout_ms // 1000)
    try:
        result = subprocess.run(
            f"ping {param} {timeout} {ip}",
            shell=True, capture_output=True, text=True,
            timeout=timeout_ms / 1000 + 1
        )
        return result.returncode == 0
    except Exception:
        return False


def add_secondary_ip(adapter: str, ip: str, mask: str = "255.255.255.0") -> tuple[bool, str]:
    """给网卡添加辅助 IP 地址。Returns (success, message)。"""
    rc, out, err = _run_netsh(
        f'interface ip add address "{adapter}" {ip} {mask}'
    )
    if rc == 0:
        return True, f"辅助 IP {ip} 已添加到「{adapter}」"
    if 'already' in err.lower() or '已存在' in err or '存在' in err:
        return True, f"IP {ip} 已存在于「{adapter}」（无需重复添加）"
    return False, f"添加失败: {err.strip() or '未知错误'}"


def remove_secondary_ip(adapter: str, ip: str) -> tuple[bool, str]:
    """从网卡移除辅助 IP 地址。Returns (success, message)。"""
    rc, out, err = _run_netsh(
        f'interface ip delete address "{adapter}" {ip}'
    )
    if rc == 0:
        return True, f"已从「{adapter}」移除 IP {ip}"
    if 'not' in err.lower() or '不存在' in err or '找不到' in err:
        return True, f"IP {ip} 在「{adapter}」上不存在（无需移除）"
    return False, f"移除失败: {err.strip() or '未知错误'}"


# ================================================================
# 网络配置弹窗
# ================================================================

_IP_NAME_MAP = {
    'fsv3004_rin':        'FSV3004 (RIN)',
    'fsv3004_linewidth':  'FSV3004 (线宽)',
    'single_freq_sa':     '频谱仪 (单频)',
    'sig_gen_linewidth':  '信号发生器 (线宽)',
    'sig_gen_wavelength': '信号发生器 (PZT调制)',
    'sig_gen_timedomain': '信号发生器 (时域)',
    'scope':              '示波器 (时域)',
    'osa':                '光谱仪 (信噪比)',
}


class NetworkConfigDialog:
    """网络配置弹窗 —— 显示仪器连通状态，一键配置子网。"""

    def __init__(self, parent: tk.Widget, *, width: int = 400, height: int = 360):
        self.parent = parent
        self._width = width
        self._height = height
        self._instrument_ips = _extract_instrument_ips()
        self._subnet_cidr = _detect_instrument_subnet()
        self._suggested_ip = (
            _suggest_local_ip(self._subnet_cidr)
            if self._subnet_cidr else ""
        )
        self._tree_items: dict[str, str] = {}   # ip → tree item id
        self._scanning_after_id: Optional[str] = None

        self._build_dialog()

    # ==================================================================
    # 窗口骨架
    # ==================================================================

    def _build_dialog(self):
        self.win = tk.Toplevel(self.parent)
        self.win.title("网络配置")
        self.win.geometry(f"{self._width}x{self._height}")
        self.win.resizable(True, True)
        self.win.configure(bg=COLOR_BG)
        self.win.transient(self.parent)
        try:
            self.win.iconbitmap("PreciLasers.ico")
        except Exception:
            pass

        # ---- 标题栏 ----
        header = tk.Frame(self.win, bg=COLOR_BG)
        header.pack(fill=tk.X, padx=PADDING_SECTION, pady=(14, 0))

        accent = tk.Frame(header, bg=COLOR_ACCENT, width=4, height=20)
        accent.pack(side=tk.LEFT, padx=(0, 8))
        accent.pack_propagate(False)

        tk.Label(header, text="仪器子网配置",
                 font=(FONT_FAMILY, FONT_SIZE_HEADING, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_BG).pack(side=tk.LEFT)

        # ---- 主内容卡片（配置 + 仪器表合一） ----
        card = tk.Frame(self.win, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_ENTRY_BORDER,
                        highlightthickness=1, bd=0)
        card.pack(fill=tk.BOTH, expand=True,
                  padx=PADDING_SECTION, pady=(10, 0))

        self._build_config_rows(card)
        self._build_status_table(card)

        # ---- 底部操作栏 ----
        self._build_bottom_bar()

        # ---- 居中 ----
        self.win.update_idletasks()
        pw = self.parent.winfo_width()
        ph = self.parent.winfo_height()
        px = self.parent.winfo_x()
        py = self.parent.winfo_y()
        dw = self.win.winfo_width()
        dh = self.win.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        self.win.geometry(f"+{x}+{y}")

        self._refresh_status()

    # ==================================================================
    # 配置行（卡片上部）
    # ==================================================================

    def _build_config_rows(self, card: tk.Frame):
        """三行参数：适配器 / 子网 / 辅助 IP。"""
        container = tk.Frame(card, bg=COLOR_CARD_BG)
        container.pack(fill=tk.X, padx=14, pady=(10, 6))

        adapters = get_active_adapters()
        adapter_names = [a["name"] for a in adapters] if adapters else ["(未检测到适配器)"]

        self._adapter_var = tk.StringVar(value=adapter_names[0])
        self._ip_var = tk.StringVar(value=self._suggested_ip)

        LABEL_W = 11

        # --- 行1：适配器 ---
        r1 = tk.Frame(container, bg=COLOR_CARD_BG)
        r1.pack(fill=tk.X, pady=2)
        tk.Label(r1, text="网络适配器", width=LABEL_W, anchor="e",
                 font=(FONT_FAMILY, FONT_SIZE_BODY),
                 fg=COLOR_TEXT_SECONDARY, bg=COLOR_CARD_BG).pack(side=tk.LEFT, padx=(0, 8))
        cb = ttk.Combobox(r1, textvariable=self._adapter_var,
                          values=adapter_names, state="readonly",
                          font=(FONT_FAMILY, FONT_SIZE_BODY))
        cb.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- 行2：子网 ---
        r2 = tk.Frame(container, bg=COLOR_CARD_BG)
        r2.pack(fill=tk.X, pady=2)
        tk.Label(r2, text="仪器子网", width=LABEL_W, anchor="e",
                 font=(FONT_FAMILY, FONT_SIZE_BODY),
                 fg=COLOR_TEXT_SECONDARY, bg=COLOR_CARD_BG).pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(r2, text=self._subnet_cidr or "(未检测到)",
                 font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                 fg=COLOR_ACCENT, bg=COLOR_CARD_BG).pack(side=tk.LEFT)

        # --- 行3：辅助 IP ---
        r3 = tk.Frame(container, bg=COLOR_CARD_BG)
        r3.pack(fill=tk.X, pady=2)
        tk.Label(r3, text="本机辅助IP", width=LABEL_W, anchor="e",
                 font=(FONT_FAMILY, FONT_SIZE_BODY),
                 fg=COLOR_TEXT_SECONDARY, bg=COLOR_CARD_BG).pack(side=tk.LEFT, padx=(0, 8))
        tk.Entry(r3, textvariable=self._ip_var,
                 font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                 bg=COLOR_ENTRY_BG, fg=COLOR_TEXT_PRIMARY,
                 relief="solid", bd=1, width=16,
                 highlightbackground=COLOR_ENTRY_BORDER,
                 insertbackground=COLOR_ACCENT).pack(side=tk.LEFT, ipady=1)
        tk.Label(r3, text="（建议值，可改）",
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_BG).pack(side=tk.LEFT, padx=(6, 0))

    # ==================================================================
    # 仪器连通状态表（卡片下部）
    # ==================================================================

    def _build_status_table(self, card: tk.Frame):
        """Treeview 表格显示仪器连通状态。"""
        # 分隔线
        sep = tk.Frame(card, bg=COLOR_ENTRY_BORDER, height=1)
        sep.pack(fill=tk.X, padx=14, pady=(6, 0))

        # 标题行
        title_row = tk.Frame(card, bg=COLOR_CARD_BG)
        title_row.pack(fill=tk.X, padx=14, pady=(6, 2))

        accent = tk.Frame(title_row, bg=COLOR_ACCENT, width=3, height=14)
        accent.pack(side=tk.LEFT, padx=(0, 6))
        accent.pack_propagate(False)

        tk.Label(title_row, text="仪器连通状态",
                 font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD_BG).pack(side=tk.LEFT)

        self._scanning_var = tk.StringVar(value="")
        self._scan_lbl = tk.Label(title_row, textvariable=self._scanning_var,
                                  font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                                  fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_BG)
        self._scan_lbl.pack(side=tk.LEFT, padx=(10, 0))

        # Treeview 表格
        tree_frame = tk.Frame(card, bg=COLOR_CARD_BG)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(2, 8))

        columns = ("instrument", "ip", "status")
        self._tree = ttk.Treeview(tree_frame, columns=columns,
                                  show="headings", height=6)
        self._tree.heading("instrument", text="仪器")
        self._tree.heading("ip", text="IP 地址")
        self._tree.heading("status", text="状态")

        self._tree.column("instrument", width=260, stretch=False)
        self._tree.column("ip", width=180, stretch=False, anchor="center")
        self._tree.column("status", width=140, stretch=False, anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)

        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 样式
        style = ttk.Style()
        style.configure("Network.Treeview",
                        background=COLOR_WHITE, fieldbackground=COLOR_WHITE,
                        rowheight=36, font=(FONT_FAMILY, FONT_SIZE_SMALL),
                        borderwidth=0)
        style.configure("Network.Treeview.Heading",
                        font=(FONT_FAMILY, FONT_SIZE_SMALL, "bold"),
                        background=COLOR_SECTION_BG, foreground=COLOR_TEXT_SECONDARY,
                        borderwidth=0, relief="flat")
        self._tree.configure(style="Network.Treeview")

        # 填充初始数据
        self._populate_tree()

    def _populate_tree(self):
        """向 Treeview 填充仪器行，每字段一行（同 IP 不合并）。"""
        self._tree_items.clear()
        for item in self._tree.get_children():
            self._tree.delete(item)

        if not self._instrument_ips:
            self._tree.insert("", "end", values=("(未检测到仪器 IP)", "", "请在配置参数中设置"))
            return

        for field, ip in self._instrument_ips:
            name = _IP_NAME_MAP.get(field, field)
            iid = self._tree.insert("", "end", values=(name, ip, "待检测"))
            self._tree_items[ip] = iid

    # ==================================================================
    # 底部操作栏
    # ==================================================================

    def _build_bottom_bar(self):
        """底部：提示 + 按钮。"""
        bar = tk.Frame(self.win, bg=COLOR_BG)
        bar.pack(fill=tk.X, padx=PADDING_SECTION, pady=(6, 10))

        # 提示（上方行）
        tip_row = tk.Frame(bar, bg=COLOR_BG)
        tip_row.pack(fill=tk.X, pady=(0, 6))

        self._tip_var = tk.StringVar(value="提示：配置子网需要管理员权限，仅需执行一次")
        tk.Label(tip_row, textvariable=self._tip_var,
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_TEXT_MUTED, bg=COLOR_BG).pack(side=tk.LEFT)

        # 按钮行
        btn_row = tk.Frame(bar, bg=COLOR_BG)
        btn_row.pack(fill=tk.X)

        # 配置子网（主操作，最右）
        tk.Button(btn_row, text="配置子网",
                  font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                  bg=COLOR_ACCENT, fg=COLOR_WHITE,
                  activebackground=COLOR_ACCENT_HOVER,
                  activeforeground=COLOR_WHITE,
                  relief="flat", bd=0, padx=16, pady=5,
                  command=self._configure_subnet, cursor="hand2"
                  ).pack(side=tk.RIGHT, padx=(4, 0))

        # 检测连通性
        tk.Button(btn_row, text="检测连通性",
                  font=(FONT_FAMILY, FONT_SIZE_BODY),
                  bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
                  activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
                  relief="solid", bd=1, padx=12, pady=5,
                  command=self._refresh_status, cursor="hand2"
                  ).pack(side=tk.RIGHT, padx=4)

        # 关闭
        tk.Button(btn_row, text="关闭",
                  font=(FONT_FAMILY, FONT_SIZE_BODY),
                  bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
                  activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
                  relief="solid", bd=1, padx=12, pady=5,
                  command=self.win.destroy, cursor="hand2"
                  ).pack(side=tk.LEFT)

        # 移除配置
        tk.Button(btn_row, text="移除配置",
                  font=(FONT_FAMILY, FONT_SIZE_BODY),
                  bg=COLOR_CARD_BORDER, fg="#C62828",
                  activebackground="#D0D5DB", activeforeground="#B71C1C",
                  relief="solid", bd=1, padx=12, pady=5,
                  command=self._remove_subnet, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=(8, 4))

    # ==================================================================
    # 操作逻辑
    # ==================================================================

    def _refresh_status(self):
        """检测所有仪器 IP 的连通性，更新 Treeview。"""
        self._scanning_var.set("检测中...")
        self._scan_lbl.config(fg=COLOR_TEXT_MUTED)

        all_items = _extract_instrument_ips()
        if all_items != self._instrument_ips:
            self._instrument_ips = all_items
            self._populate_tree()

        # 后台线程串行 ping
        import threading
        def _ping_all():
            results = {}
            for _, ip in self._instrument_ips:
                results[ip] = ping(ip)
            self.win.after(0, lambda: self._apply_ping_results(results))

        threading.Thread(target=_ping_all, daemon=True).start()

    def _apply_ping_results(self, results: dict[str, bool]):
        """在主线程中更新 Treeview 的 ping 结果。"""
        for field, ip in self._instrument_ips:
            iid = self._tree_items.get(ip)
            if not iid:
                continue
            name = _IP_NAME_MAP.get(field, field)
            status_text = "● 可达" if results.get(ip, False) else "✗ 不可达"
            tag = "reachable" if results.get(ip, False) else "unreachable"
            self._tree.item(iid, values=(name, ip, status_text))
            self._tree.item(iid, tags=(tag,))

        self._tree.tag_configure("reachable", foreground=COLOR_SUCCESS)
        self._tree.tag_configure("unreachable", foreground="#E57373")

        self._scanning_var.set("")
        self._scan_lbl.config(fg=COLOR_TEXT_MUTED)

    def _configure_subnet(self):
        """执行子网配置：以管理员权限运行 netsh 添加辅助 IP。"""
        adapter = self._adapter_var.get()
        ip = self._ip_var.get().strip()
        if not ip:
            messagebox.showwarning("提示", "请填写本机辅助 IP 地址", parent=self.win)
            return
        if not re.match(r'^\d+\.\d+\.\d+\.\d+$', ip):
            messagebox.showwarning("提示", f"IP 地址格式无效: {ip}", parent=self.win)
            return

        mask = "255.255.255.0"
        if self._subnet_cidr:
            try:
                net_obj = ipaddress.IPv4Network(self._subnet_cidr, strict=False)
                mask = str(net_obj.netmask)
            except Exception:
                pass

        if not self._is_admin():
            self._run_as_admin(adapter, ip, mask)
        else:
            ok, msg = add_secondary_ip(adapter, ip, mask)
            if ok:
                self._tip_var.set(msg)
                self.win.after(500, self._refresh_status)
            else:
                messagebox.showerror("配置失败", msg, parent=self.win)

    def _run_as_admin(self, adapter: str, ip: str, mask: str):
        """以管理员权限执行 netsh（直接调 netsh.exe，打包后也能用）。"""
        params = f'interface ip add address "{adapter}" {ip} {mask}'
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "netsh", params, None, 1
            )
            self._tip_var.set("已提权执行，完成后请点击「检测连通性」验证")
        except Exception as e:
            messagebox.showerror("提权失败",
                                 f"无法以管理员权限运行: {e}\n\n"
                                 "请右键以管理员身份运行本程序，或手动执行：\n"
                                 f'netsh {params}',
                                 parent=self.win)

    def _remove_subnet(self):
        """移除已配置的辅助 IP。"""
        adapter = self._adapter_var.get()
        ip = self._ip_var.get().strip()
        if not ip:
            messagebox.showwarning("提示", "请填写要移除的 IP 地址", parent=self.win)
            return

        if not self._is_admin():
            params = f'interface ip delete address "{adapter}" {ip}'
            try:
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", "netsh", params, None, 1
                )
                self._tip_var.set("已提权执行移除，完成后请点击「检测连通性」验证")
            except Exception as e:
                messagebox.showerror("提权失败", f"无法以管理员权限运行: {e}", parent=self.win)
        else:
            ok, msg = remove_secondary_ip(adapter, ip)
            if ok:
                self._tip_var.set(msg)
            else:
                messagebox.showerror("移除失败", msg, parent=self.win)

    @staticmethod
    def _is_admin() -> bool:
        """检查当前进程是否以管理员权限运行。"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            return False
