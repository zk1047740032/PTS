"""
配置参数面板 —— 分类展示 path_a / path_b 所有 GUI 程序的默认参数，
支持在线编辑并持久化到 config/user_config.json。

设计原则：
- 参数定义与 CFG 解耦：面板字段手动编排，不依赖 dataclass 反射。
- 保存格式：JSON 扁平 key-value，key 为点分路径（如 "network.fsv3004_rin"）。
- 加载覆盖：主平台启动时读 JSON → object.__setattr__ 写回 CFG。
"""

from __future__ import annotations

import json
import os
from collections import OrderedDict
from pathlib import Path
from tkinter import ttk
import tkinter as tk

from core.config import CFG
from utils.theme import (
    COLOR_BG, COLOR_WHITE, COLOR_CARD_BG, COLOR_CARD_BORDER,
    COLOR_SECTION_BG, COLOR_SECTION_FG,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    COLOR_ACCENT, COLOR_ACCENT_HOVER, COLOR_SUCCESS,
    COLOR_ENTRY_BORDER, COLOR_ENTRY_BG,
    FONT_FAMILY, FONT_SIZE_BODY, FONT_SIZE_SMALL, FONT_SIZE_CAPTION,
    PADDING_SECTION,
)

# ---- 用户配置 JSON 路径 ----
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_USER_CONFIG_PATH = _PROJECT_ROOT / "config" / "user_config.json"

# ---- 窗口基准尺寸 ----
_BASE_CONFIG_W = 820
_BASE_CONFIG_H = 620

# ================================================================
# 参数 Schema 定义
# 结构: OrderedDict {
#     标签页标题: OrderedDict {
#         分组标题: [
#             (key, 中文标签),
#             ...
#         ],
#         ...
#     },
#     ...
# }
#
# key 的格式: "cfg_sub_object.field_name"
# 特殊格式: "cfg_sub_object.tuple_field[index]"  用于元组元素
# ================================================================

PARAM_SCHEMA: OrderedDict[str, OrderedDict[str, list[tuple[str, str]]]] = OrderedDict([

    # ======================== 通用设置 ========================
    ("通用设置", OrderedDict([
        ("仪器 IP 地址", [
            ("network.fsv3004_rin",        "FSV3004 (RIN)"),
            ("network.fsv3004_linewidth",  "FSV3004 (线宽)"),
            ("network.single_freq_sa",     "频谱仪 (单频)"),
            ("network.sig_gen_linewidth",  "信号发生器 (线宽)"),
            ("network.sig_gen_wavelength", "信号发生器 (PZT调制)"),
            ("network.sig_gen_timedomain", "信号发生器 (时域)"),
            ("network.scope",              "示波器 (时域)"),
            ("network.osa",                "光谱仪 (信噪比)"),
        ]),
        ("USB 设备", [
            ("usb.optical_switch", "光开关 VISA 地址"),
        ]),
        ("串口配置", [
            ("serial.port_default", "默认串口"),
            ("serial.baudrate",     "波特率"),
            ("serial.timeout_s",    "超时 (秒)"),
            ("serial.device_addr",  "设备地址"),
        ]),
        ("数据目录", [
            ("dirs.power",             "功率数据"),
            ("dirs.wavelength",        "波长数据"),
            ("dirs.spectrum_snr",      "信噪比数据"),
            ("dirs.rin_fsv3004",       "RIN 数据"),
            ("dirs.line_width",        "线宽数据"),
            ("dirs.single_freq_1um",   "单频 1μm"),
            ("dirs.single_freq_1_5um", "单频 1.5μm"),
            ("dirs.time_domain",       "时域数据"),
            ("dirs.seed_value",        "种子参数"),
            ("dirs.report",            "报告输出"),
        ]),
        ("系统路径", [
            ("sys_paths.wlm_dll",            "波长计 DLL"),
            ("sys_paths.arial_font",         "Arial 字体"),
            ("sys_paths.instrument_temp_png", "线宽截图临时路径"),
            ("sys_paths.rin_temp_png",       "RIN 截图临时路径"),
        ]),
        ("时序参数", [
            ("timing.channel_switch_delay_s", "通道切换延时 (秒)"),
            ("timing.visa_short_ms",          "VISA 短超时 (ms)"),
            ("timing.visa_default_ms",        "VISA 默认超时 (ms)"),
            ("timing.visa_medium_ms",         "VISA 中等超时 (ms)"),
            ("timing.visa_long_ms",           "VISA 长超时 (ms)"),
            ("timing.visa_extra_long_ms",     "VISA 超长超时 (ms)"),
            ("timing.auto_close_short_ms",    "自动关窗 短 (ms)"),
            ("timing.auto_close_long_ms",     "自动关窗 长 (ms)"),
            ("timing.stabilize_1s",           "稳定等待 (秒)"),
            ("timing.stabilize_3s",           "稳定等待 长 (秒)"),
            ("timing.default_retries",        "默认重试次数"),
            ("timing.opc_timeout_ms",         "OPC 超时 (ms)"),
        ]),
    ])),

    # ======================== RIN (path_a/Rin_FSV3004.py) ========================
    ("RIN", OrderedDict([
        ("RIN 测量 — DC / 放大", [
            ("rin.dc_internal", "内部 DC 值 (公式用)"),
            ("rin.dc_initial",  "GUI 默认 DC 值"),
            ("rin.amplification", "放大倍数"),
            ("rin.sweep_points",  "扫描点数"),
        ]),
        ("RIN 测量 — 弛豫振荡峰值搜索", [
            ("rin.relax_start_hz",  "搜索起始频率 (Hz)"),
            ("rin.relax_stop_hz",   "搜索终止频率 (Hz)"),
            ("rin.peak_min_points", "峰值最小点数"),
        ]),
        ("RIN 频段 (segments)", [
            ("rin.segments[0]", "段1: 10–100Hz, RBW=5, Avg=20"),
            ("rin.segments[1]", "段2: 100–1kHz, RBW=5, Avg=20"),
            ("rin.segments[2]", "段3: 1k–10kHz, RBW=30, Avg=20"),
            ("rin.segments[3]", "段4: 10k–100kHz, RBW=30, Avg=20"),
            ("rin.segments[4]", "段5: 100k–1MHz, RBW=30, Avg=20"),
            ("rin.segments[5]", "段6: 1M–10MHz, RBW=30, Avg=20"),
        ]),
        ("背景噪声测量", [
            ("bg_noise.start_freq", "起始频率 (Hz)"),
            ("bg_noise.stop_freq",  "终止频率 (Hz)"),
            ("bg_noise.bandwidth",  "分辨率带宽 (Hz)"),
            ("bg_noise.avg_count",  "平均次数"),
        ]),
    ])),

    # ======================== 线宽 (path_a/LineWidth_FSV3004.py) ========================
    ("线宽", OrderedDict([
        ("频宽选项 (Span)", [
            ("linewidth.span_values", "Span 选项 (kHz, 逗号分隔)"),
        ]),
        ("测量参数", [
            ("linewidth.rbw_hz",         "RBW (Hz)"),
            ("linewidth.ref_level_mv",   "参考电平 (mV)"),
            ("linewidth.m1_position_mhz","M1 位置 (MHz)"),
            ("linewidth.center_freq_mhz","中心频率 (MHz)"),
            ("linewidth.n_db_down",      "N dB down"),
            ("linewidth.sweep_points",   "扫描点数"),
            ("linewidth.avg_count",      "平均次数"),
        ]),
        ("信号发生器额外测试", [
            ("linewidth.extra_span",   "额外测试 Span (kHz)"),
            ("linewidth.extra_freq",   "额外测试 频率 (MHz)"),
            ("linewidth.extra_volt",   "额外测试 电压 (V)"),
            ("linewidth.extra_offset", "额外测试 偏置 (V)"),
            ("linewidth.extra_suffix", "额外测试 文件名后缀"),
        ]),
        ("连接 / 重试", [
            ("linewidth.connect_max_retries",        "连接最大重试次数"),
            ("linewidth.connect_retry_interval_s",   "连接重试间隔 (秒)"),
            ("linewidth.ndbdown_read_timeout_ms",    "N dB down 读取超时 (ms)"),
            ("linewidth.save_data_max_retries",      "保存数据最大重试"),
            ("linewidth.reconnect_delay_s",          "重连延时 (秒)"),
            ("linewidth.signal_stabilize_s",         "信号稳定等待 (秒)"),
        ]),
    ])),

    # ======================== 时域 (path_a/TimeDomain.py) ========================
    ("时域", OrderedDict([
        ("测试参数", [
            ("time_domain.test_frequencies", "测试频率列表 (Hz, 逗号分隔)"),
            ("time_domain.gen_volt",         "信号发生器电压 (V)"),
            ("time_domain.gen_offset",       "信号发生器偏置 (V)"),
            ("time_domain.gen_channel",      "信号发生器通道"),
            ("time_domain.scope_channel",    "示波器通道"),
        ]),
        ("Vpp / 稳定判定", [
            ("time_domain.vpp_stable_count",   "Vpp 稳定判定次数"),
            ("time_domain.vpp_stable_delay_s", "Vpp 稳定间隔 (秒)"),
            ("time_domain.vpp_min_valid",      "Vpp 最小有效值"),
            ("time_domain.vpp_max_valid",      "Vpp 最大有效值"),
            ("time_domain.vpp_default",        "Vpp 默认值"),
        ]),
        ("测量 / 重试", [
            ("time_domain.final_vpp_count",    "最终 Vpp 采集次数"),
            ("time_domain.final_vpp_delay_s",  "最终 Vpp 采集间隔 (秒)"),
            ("time_domain.meas_retries",       "测量重试次数"),
            ("time_domain.meas_retry_delay_s", "测量重试间隔 (秒)"),
            ("time_domain.signal_stabilize_s", "信号稳定等待 (秒)"),
            ("time_domain.gen_settle_s",       "发生器稳定时间 (秒)"),
        ]),
    ])),

    # ======================== 信噪比 (path_a/SpectrumSNR.py) ========================
    ("信噪比", OrderedDict([
        ("测量参数", [
            ("spectrum_snr.center_nm",              "中心波长 (nm)"),
            ("spectrum_snr.span_nm",                "跨度 (nm)"),
            ("spectrum_snr.ref_level_dbm",          "参考电平 (dBm)"),
            ("spectrum_snr.sensitivity",            "灵敏度"),
            ("spectrum_snr.main_peak_exclusion_nm", "主峰排除范围 (nm)"),
        ]),
        ("超时 / 重试", [
            ("spectrum_snr.visa_timeout_s",        "VISA 超时 (秒)"),
            ("spectrum_snr.screenshot_timeout_ms", "截图超时 (ms)"),
            ("spectrum_snr.query_retries",         "查询重试次数"),
            ("spectrum_snr.query_retry_delay_s",   "查询重试间隔 (秒)"),
        ]),
    ])),

    # ======================== 单频 (path_a/SingleFrequency.py) ========================
    ("单频", OrderedDict([
        ("通用", [
            ("single_freq.test_duration_min", "测试时长 (分钟)"),
        ]),
        ("1μm — 温度扫描", [
            ("single_freq.temp_max_1um",      "温度上限 (°C)"),
            ("single_freq.temp_min_1um",      "温度下限 (°C)"),
            ("single_freq.temp_step_1um",     "温度步长 (°C)"),
            ("single_freq.temp_interval_1um", "温度变化间隔 (秒)"),
        ]),
        ("1μm — 电流扫描", [
            ("single_freq.cur_max_1um",      "电流上限 (mA)"),
            ("single_freq.cur_min_1um",      "电流下限 (mA)"),
            ("single_freq.cur_step_1um",     "电流步长 (mA)"),
            ("single_freq.cur_interval_1um", "电流变化间隔 (秒)"),
        ]),
        ("1.5μm — 温度扫描", [
            ("single_freq.temp_max_1_5um",      "温度上限 (°C)"),
            ("single_freq.temp_min_1_5um",      "温度下限 (°C)"),
            ("single_freq.temp_step_1_5um",     "温度步长 (°C)"),
            ("single_freq.temp_interval_1_5um", "温度变化间隔 (秒)"),
        ]),
        ("1.5μm — 电流扫描", [
            ("single_freq.cur_max_1_5um",      "电流上限 (mA)"),
            ("single_freq.cur_min_1_5um",      "电流下限 (mA)"),
            ("single_freq.cur_step_1_5um",     "电流步长 (mA)"),
            ("single_freq.cur_interval_1_5um", "电流变化间隔 (秒)"),
        ]),
        ("频谱扫描 / 峰值检测", [
            ("single_freq.sweep_span_mhz",     "扫描跨度 (MHz)"),
            ("single_freq.sweep_step_mhz",     "扫描步长 (MHz)"),
            ("single_freq.sweep_freq_max_mhz", "最大扫描频率 (MHz)"),
            ("single_freq.rbw_khz",            "RBW (kHz)"),
            ("single_freq.avg_count",          "平均次数"),
            ("single_freq.sweep_time_s",       "扫描时间 (秒)"),
            ("single_freq.fine_repeat_count",  "精细重复次数"),
            ("single_freq.peak_threshold_db",  "峰值阈值 (dB)"),
            ("single_freq.peak_prominence_db", "峰值显著性 (dB)"),
            ("single_freq.peak_guard_points",  "峰值保护点数"),
        ]),
        ("波长稳定等待", [
            ("single_freq.stability_tol_nm",      "容差 (nm)"),
            ("single_freq.stability_consecutive", "连续稳定次数"),
            ("single_freq.stability_max_wait_s",  "最大等待 (秒)"),
            ("single_freq.stability_interval_s",  "检查间隔 (秒)"),
        ]),
        ("超时", [
            ("single_freq.sa_timeout_s",        "频谱仪超时 (秒)"),
            ("single_freq.query_opc_timeout_s", "OPC 超时 (秒)"),
            ("single_freq.sweep_timeout_ms",    "扫描超时 (ms)"),
        ]),
    ])),

    # ======================== 功率 (path_a/Power.py) ========================
    ("功率", OrderedDict([
        ("测量配置", [
            ("power.burnin_header_lines",   "烤机 CSV 表头行数"),
            ("power.burnin_power_column",   "烤机 功率列名"),
            ("power.output_filename",       "输出文件名"),
            ("power.burnin_result_filename","烤机结果文件名"),
            ("power.fallback_save_dir",     "回退保存目录"),
        ]),
    ])),

    # ======================== PZT调制 (path_b/WaveLength.py) ========================
    ("PZT调制", OrderedDict([
        ("1μm 默认参数", [
            ("wavelength.freq_1um",   "频率 (Hz)"),
            ("wavelength.volt_1um",   "幅值 (mVpp)"),
            ("wavelength.offset_1um", "偏置 (mVdc)"),
        ]),
        ("1.5μm 默认参数", [
            ("wavelength.freq_1_5um",   "频率 (Hz)"),
            ("wavelength.volt_1_5um",   "幅值 (mVpp)"),
            ("wavelength.offset_1_5um", "偏置 (mVdc)"),
        ]),
        ("变参测试", [
            ("wavelength.sweep_text",        "变参测试参数 (多行)"),
            ("wavelength.modulation_time_s", "调制时间 (秒)"),
        ]),
        ("信号发生器默认", [
            ("wavelength.sig_gen_default_waveform", "默认波形"),
            ("wavelength.sig_gen_default_freq",     "默认频率 (Hz)"),
            ("wavelength.sig_gen_default_volt",     "默认电压 (Vpp)"),
            ("wavelength.sig_gen_default_offset",   "默认偏置 (Vdc)"),
            ("wavelength.sig_gen_connect_retries",  "连接重试次数"),
            ("wavelength.sig_gen_connect_interval", "连接重试间隔 (秒)"),
        ]),
        ("波长计 — DLL 参数", [
            ("wavelength.exposure_ms",           "曝光时间 (ms)"),
            ("wavelength.wait_event_timeout_ms", "事件等待超时 (ms)"),
            ("wavelength.instantiate_rfc",       "RFC"),
            ("wavelength.instantiate_mode",      "Mode"),
            ("wavelength.instantiate_p1",        "P1"),
            ("wavelength.instantiate_p2",        "P2"),
        ]),
        ("pywinauto 坐标 / 搜索", [
            ("wavelength.wlm_window_title",         "波长计窗口标题 (正则)"),
            ("wavelength.longterm_graph_title",     "长时图表标题 (正则)"),
            ("wavelength.reset_btn_offset_x",       "重置按钮 X 偏移"),
            ("wavelength.top_right_arrow_offset_x", "右上箭头 X 偏移"),
            ("wavelength.top_right_arrow_offset_y", "右上箭头 Y 偏移"),
            ("wavelength.close_panel_offset_x",     "关闭面板 X 偏移"),
            ("wavelength.close_panel_offset_y",     "关闭面板 Y 偏移"),
            ("wavelength.top_chart_found_index",    "上图表 found_index"),
            ("wavelength.bottom_chart_found_index", "下图表 found_index"),
            ("wavelength.reset_panel_found_index",  "重置面板 found_index"),
            ("wavelength.chart_arrow_found_index",  "图表箭头 found_index"),
        ]),
    ])),

    # ======================== 相噪 (path_a/PhaseNoise.py) ========================
    ("相噪", OrderedDict([
        ("文件 / 阈值", [
            ("phase_noise.wavelength_file",          "波长文件路径"),
            ("phase_noise.wavelength_threshold_nm",  "波长阈值 (nm)"),
        ]),
        ("窗口操作 (点击/等待)", [
            ("phase_noise.click_retries",         "点击重试次数"),
            ("phase_noise.click_delay_s",         "点击间隔 (秒)"),
            ("phase_noise.launch_wait_s",         "启动等待 (秒)"),
            ("phase_noise.post_launch_wait_s",    "启动后等待 (秒)"),
            ("phase_noise.click_retry_delay_s",   "重试间隔 (秒)"),
            ("phase_noise.check_interval_s",      "检查间隔 (秒)"),
            ("phase_noise.click_coords",          "点击坐标 (x,y)"),
            ("phase_noise.target_window_class",   "目标窗口类名"),
            ("phase_noise.target_window_title",   "目标窗口标题"),
        ]),
    ])),
])


# ================================================================
# 工具函数：从 CFG 读取 / 写入值
# ================================================================

def _parse_cfg_path(path: str):
    """解析点分路径为 (cfg_object, attr_name)。

    支持 tuple 索引: "rin.segments[0]" → ("rin", "segments", 0)
    """
    import re
    parts = path.split(".")
    # 最后一段可能是 "field[index]" 格式
    m = re.match(r"^(\w+)\[(\d+)\]$", parts[-1])
    index = None
    if m:
        parts[-1] = m.group(1)
        index = int(m.group(2))
    return parts, index


def get_cfg_value(path: str) -> str:
    """从 CFG 读取参数值，转为字符串。"""
    try:
        parts, index = _parse_cfg_path(path)
        obj = CFG
        for p in parts:
            obj = getattr(obj, p)
        if index is not None:
            obj = obj[index]
        # 格式化输出
        if isinstance(obj, tuple):
            return ", ".join(str(x) for x in obj)
        return str(obj)
    except Exception:
        return ""


def set_cfg_value(path: str, value_str: str):
    """将字符串值写回 CFG（通过 object.__setattr__ 绕过 frozen）。"""
    try:
        parts, index = _parse_cfg_path(path)
        obj = CFG
        for p in parts[:-1]:
            obj = getattr(obj, p)
        attr = parts[-1]

        # 获取原始值以推断类型
        old_val = getattr(obj, attr)
        if index is not None:
            old_val = old_val[index]

        # 类型转换
        new_val = _coerce_value(value_str, old_val)

        if index is not None:
            # 修改元组中的元素
            old_tuple = getattr(obj, attr)
            lst = list(old_tuple)
            lst[index] = new_val
            object.__setattr__(obj, attr, tuple(lst))
        else:
            object.__setattr__(obj, attr, new_val)
    except Exception:
        pass  # 静默忽略无法写入的字段


def _coerce_value(value_str: str, old_val):
    """根据旧值类型将字符串转换为对应类型。"""
    if isinstance(old_val, bool):
        return value_str.strip().lower() in ("true", "1", "yes")
    if isinstance(old_val, int):
        try:
            return int(value_str.strip())
        except ValueError:
            return old_val
    if isinstance(old_val, float):
        try:
            return float(value_str.strip())
        except ValueError:
            return old_val
    if isinstance(old_val, tuple):
        # 逗号分隔解析，逐元素匹配原类型（支持异构元组如 (int, int, int, int, str)）
        parts = [x.strip() for x in value_str.split(",") if x.strip()]
        result = []
        for i, part in enumerate(parts):
            if i < len(old_val):
                ref = old_val[i]
                if isinstance(ref, int):
                    try:
                        result.append(int(part))
                        continue
                    except ValueError:
                        pass
                elif isinstance(ref, float):
                    try:
                        result.append(float(part))
                        continue
                    except ValueError:
                        pass
            result.append(part)
        return tuple(result)
    # 默认保持字符串
    return value_str.strip()


# ================================================================
# 用户配置 JSON 读写
# ================================================================

def load_user_overrides() -> dict[str, str]:
    """加载用户覆盖配置。"""
    try:
        if _USER_CONFIG_PATH.exists():
            with open(_USER_CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_user_overrides(data: dict[str, str]):
    """保存用户覆盖配置到 JSON 文件。"""
    _USER_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_USER_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def apply_user_overrides_to_cfg():
    """启动时将 user_config.json 中的覆盖值应用到 CFG。"""
    overrides = load_user_overrides()
    for key, value_str in overrides.items():
        set_cfg_value(key, value_str)


# ================================================================
# 配置参数面板 UI
# ================================================================

class ConfigParamsPanel:
    """配置参数面板 —— 嵌入 Toplevel 使用。"""

    def __init__(self, parent: tk.Widget):
        """
        Args:
            parent: 父容器（通常是 Toplevel）
        """
        self.parent = parent
        # key → tk.StringVar 映射，用于收集用户输入
        self._entry_vars: dict[str, tk.StringVar] = {}
        # key → widget type ("entry" | "text")
        self._widget_types: dict[str, str] = {}

        self._build_ui()

    # ---------- UI 构建 ----------

    def _build_ui(self):
        """构建完整的配置面板 UI。"""
        # 主题样式
        style = ttk.Style()
        style.theme_use("clam")
        self._configure_styles(style)

        # 主容器
        main = tk.Frame(self.parent, bg=COLOR_BG)
        main.pack(fill=tk.BOTH, expand=True)

        # ---- 标题 ----
        header = tk.Frame(main, bg=COLOR_BG)
        header.pack(fill=tk.X, padx=PADDING_SECTION, pady=(16, 0))

        accent = tk.Frame(header, bg=COLOR_ACCENT, width=4, height=22)
        accent.pack(side=tk.LEFT, padx=(0, 10))
        accent.pack_propagate(False)

        tk.Label(header, text="默认参数配置",
                 font=(FONT_FAMILY, 14, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_BG).pack(side=tk.LEFT)

        tk.Label(header, text="修改后点击「保存」生效，重启程序后自动加载",
                 font=(FONT_FAMILY, FONT_SIZE_CAPTION),
                 fg=COLOR_TEXT_MUTED, bg=COLOR_BG).pack(side=tk.LEFT, padx=(14, 0), pady=(6, 0))

        # ---- Notebook 标签页（懒加载：首屏立即构建，其余切换时才构建） ----
        nb_frame = tk.Frame(main, bg=COLOR_BG)
        nb_frame.pack(fill=tk.BOTH, expand=True, padx=PADDING_SECTION, pady=(8, 0))

        self._notebook = ttk.Notebook(nb_frame)
        self._notebook.pack(fill=tk.BOTH, expand=True)

        # 记录每个标签页的占位容器、schema、是否已构建
        self._tab_frames: dict[str, tk.Frame] = {}
        self._tab_sections: dict[str, OrderedDict] = {}
        self._tabs_built: set[str] = set()

        for tab_title, sections in PARAM_SCHEMA.items():
            placeholder = tk.Frame(self._notebook, bg=COLOR_BG)
            self._notebook.add(placeholder, text=f"  {tab_title}  ")
            self._tab_frames[tab_title] = placeholder
            self._tab_sections[tab_title] = sections

        # 首屏立即构建（默认显示第一个标签页）
        first_title = list(PARAM_SCHEMA.keys())[0]
        self._populate_tab_content(self._tab_frames[first_title],
                                   self._tab_sections[first_title])
        self._tabs_built.add(first_title)

        # 切换标签页时按需构建
        self._notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # ---- 底部操作栏 ----
        self._build_bottom_bar(main)

    def _configure_styles(self, style: ttk.Style):
        """配置 ttk 样式。"""
        style.configure("Config.TNotebook", background=COLOR_BG, borderwidth=0)
        style.configure("Config.TNotebook.Tab",
                        font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                        padding=[18, 6],
                        background=COLOR_SECTION_BG, foreground=COLOR_TEXT_SECONDARY,
                        borderwidth=0)
        style.map("Config.TNotebook.Tab",
                  background=[("selected", COLOR_WHITE)],
                  foreground=[("selected", COLOR_ACCENT)])

        # 可滚动区域的滚动条
        style.configure("Config.Vertical.TScrollbar",
                        background=COLOR_BG, troughcolor=COLOR_BG,
                        borderwidth=0, arrowsize=14)

        style.configure("Config.TFrame", background=COLOR_BG)
        style.configure("Config.TLabelframe", background=COLOR_BG,
                        borderwidth=0, relief="flat")
        style.configure("Config.TLabelframe.Label",
                        font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                        foreground=COLOR_TEXT_PRIMARY,
                        background=COLOR_BG)

    def _populate_tab_content(self, container: tk.Frame, sections: OrderedDict):
        """向已有标签页容器中填充可滚动的参数分组内容（懒加载用）。"""
        # 可滚动画布
        canvas = tk.Canvas(container, bg=COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview,
                                  style="Config.Vertical.TScrollbar")
        scroll_frame = tk.Frame(canvas, bg=COLOR_BG)

        scroll_frame.bind(
            "<Configure>",
            lambda *_: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw",
                             tags="scroll_inner")
        # 内容宽度跟随画布，避免横向溢出
        canvas.bind("<Configure>",
                     lambda e: canvas.itemconfig("scroll_inner", width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        # 鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda *_: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda *_: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 内容区 —— 适当留白
        inner = tk.Frame(scroll_frame, bg=COLOR_BG)
        inner.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)

        for section_title, fields in sections.items():
            self._build_section(inner, section_title, fields)

    def _on_tab_changed(self, event):
        """标签页切换时按需懒加载内容。"""
        try:
            current_idx = self._notebook.index("current")
        except Exception:
            return
        tab_titles = list(PARAM_SCHEMA.keys())
        if current_idx >= len(tab_titles):
            return
        tab_title = tab_titles[current_idx]
        if tab_title not in self._tabs_built:
            self._populate_tab_content(self._tab_frames[tab_title],
                                       self._tab_sections[tab_title])
            self._tabs_built.add(tab_title)

    # 交替行背景色（比纯白略深，形成微妙的斑马纹）
    _ROW_BG_EVEN = COLOR_CARD_BG   # #FFFFFF
    _ROW_BG_ODD  = "#F8F9FB"       # 极浅灰蓝

    def _build_section(self, parent: tk.Widget, title: str,
                       fields: list[tuple[str, str]]):
        """构建一个参数分组 —— 卡片式布局，带左侧强调条。"""
        # 卡片容器（白底 + 细边框）
        card = tk.Frame(parent, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_ENTRY_BORDER,
                        highlightthickness=1, bd=0)
        card.pack(fill=tk.X, pady=(0, 10))

        # ---- 分组标题栏（左侧强调条 + 标题） ----
        title_bar = tk.Frame(card, bg=COLOR_CARD_BG)
        title_bar.pack(fill=tk.X)

        accent = tk.Frame(title_bar, bg=COLOR_ACCENT, width=3, height=16)
        accent.pack(side=tk.LEFT, padx=(14, 8), pady=(10, 0))
        accent.pack_propagate(False)

        tk.Label(title_bar, text=title,
                 font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                 fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD_BG
                 ).pack(side=tk.LEFT, pady=(10, 0))

        # 标题与字段之间的细分隔线
        sep = tk.Frame(card, bg=COLOR_ENTRY_BORDER, height=1)
        sep.pack(fill=tk.X, padx=14, pady=(6, 2))

        # ---- 参数字段列表 ----
        fields_container = tk.Frame(card, bg=COLOR_CARD_BG)
        fields_container.pack(fill=tk.X, padx=10, pady=(2, 8))

        for i, (key, label) in enumerate(fields):
            row_bg = self._ROW_BG_ODD if i % 2 else self._ROW_BG_EVEN
            self._create_field_row(fields_container, key, label, row_bg)

    def _create_field_row(self, parent: tk.Widget, key: str, label: str,
                          row_bg: str):
        """创建单行：标签 + 输入框，支持交替背景色。"""
        row = tk.Frame(parent, bg=row_bg)
        row.pack(fill=tk.X, ipady=1)

        # 标签 —— 右对齐，固定宽度
        tk.Label(row, text=label, font=(FONT_FAMILY, FONT_SIZE_BODY),
                 fg=COLOR_TEXT_SECONDARY, bg=row_bg,
                 width=22, anchor="e").pack(side=tk.LEFT, padx=(10, 8), pady=4)

        current_val = get_cfg_value(key)

        # 多行文本参数使用 Text 控件
        if "\n" in current_val or key == "wavelength.sweep_text":
            text_frame = tk.Frame(row, bg=row_bg)
            text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), pady=3)

            text_widget = tk.Text(text_frame, height=4, width=50,
                                  font=(FONT_FAMILY, FONT_SIZE_BODY),
                                  bg=COLOR_ENTRY_BG, fg=COLOR_TEXT_PRIMARY,
                                  relief="solid", bd=1,
                                  highlightbackground=COLOR_ENTRY_BORDER,
                                  padx=6, pady=4)
            text_widget.insert("1.0", current_val)
            text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)

            self._entry_vars[key] = text_widget
            self._widget_types[key] = "text"
        else:
            var = tk.StringVar(value=current_val)
            entry = tk.Entry(row, textvariable=var,
                             font=(FONT_FAMILY, FONT_SIZE_BODY),
                             bg=COLOR_ENTRY_BG, fg=COLOR_TEXT_PRIMARY,
                             relief="solid", bd=1,
                             highlightbackground=COLOR_ENTRY_BORDER,
                             insertbackground=COLOR_ACCENT)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), pady=3,
                       ipady=2)

            self._entry_vars[key] = var
            self._widget_types[key] = "entry"

    def _build_bottom_bar(self, parent: tk.Widget):
        """底部操作栏：保存 / 恢复默认 / 关闭。"""
        bar = tk.Frame(parent, bg=COLOR_BG)
        bar.pack(fill=tk.X, padx=PADDING_SECTION, pady=(8, 14))

        # 状态提示
        self._status_var = tk.StringVar(value="")
        status_label = tk.Label(bar, textvariable=self._status_var,
                                font=(FONT_FAMILY, FONT_SIZE_SMALL),
                                fg=COLOR_TEXT_MUTED, bg=COLOR_BG)
        status_label.pack(side=tk.LEFT, pady=(4, 0))

        # 关闭按钮
        close_btn = tk.Button(bar, text="关闭",
                              font=(FONT_FAMILY, FONT_SIZE_BODY),
                              bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
                              activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
                              relief="solid", bd=1, padx=16, pady=5,
                              command=self.parent.destroy, cursor="hand2")
        close_btn.pack(side=tk.RIGHT, padx=(6, 0))

        # 恢复默认按钮
        reset_btn = tk.Button(bar, text="恢复默认",
                              font=(FONT_FAMILY, FONT_SIZE_BODY),
                              bg=COLOR_CARD_BORDER, fg=COLOR_TEXT_SECONDARY,
                              activebackground="#D0D5DB", activeforeground=COLOR_TEXT_PRIMARY,
                              relief="solid", bd=1, padx=16, pady=5,
                              command=self._reset_defaults, cursor="hand2")
        reset_btn.pack(side=tk.RIGHT, padx=6)

        # 保存按钮
        save_btn = tk.Button(bar, text="保存",
                             font=(FONT_FAMILY, FONT_SIZE_BODY, "bold"),
                             bg=COLOR_ACCENT, fg=COLOR_WHITE,
                             activebackground=COLOR_ACCENT_HOVER,
                             activeforeground=COLOR_WHITE,
                             relief="flat", bd=0, padx=20, pady=6,
                             command=self._save, cursor="hand2")
        save_btn.pack(side=tk.RIGHT, padx=6)

    # ---------- 操作逻辑 ----------

    def _collect_values(self) -> dict[str, str]:
        """收集所有参数的当前值。"""
        data = {}
        for key, widget_or_var in self._entry_vars.items():
            wtype = self._widget_types.get(key, "entry")
            if wtype == "text":
                try:
                    val = widget_or_var.get("1.0", "end-1c")
                except Exception:
                    val = ""
            else:
                val = widget_or_var.get()
            original = get_cfg_value(key)
            # 只保存与 CFG 默认值不同的项
            if val != original:
                data[key] = val
        return data

    def _save(self):
        """收集并保存当前配置到 JSON 文件。"""
        data = self._collect_values()
        save_user_overrides(data)
        # 立即应用到 CFG
        for key, value_str in data.items():
            set_cfg_value(key, value_str)
        self._show_status("配置已保存，重启程序后自动加载", "green")

    def _reset_defaults(self):
        """将所有参数恢复到 CFG 默认值，并清空用户覆盖文件。"""
        for key, widget_or_var in self._entry_vars.items():
            default_val = get_cfg_value(key)
            wtype = self._widget_types.get(key, "entry")
            if wtype == "text":
                try:
                    widget_or_var.delete("1.0", "end")
                    widget_or_var.insert("1.0", default_val)
                except Exception:
                    pass
            else:
                widget_or_var.set(default_val)
        # 清空覆盖文件
        save_user_overrides({})
        self._show_status("已恢复默认值，覆盖配置已清空", "gray")

    def _show_status(self, msg: str, color: str = "gray"):
        """显示状态提示，3 秒后消失。"""
        self._status_var.set(msg)
        self.parent.after(3000, lambda: self._status_var.set(""))
