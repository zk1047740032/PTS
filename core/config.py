"""
项目全局配置文件

将所有可配置参数集中管理：仪器地址、文件路径、测量默认值、时序参数、UI 常量。
各模块从本文件导入 CFG 单例即可获取所有配置默认值。

使用方式:
    from core.config import CFG

    ip = CFG.network.fsv3004_rin
    path = str(CFG.dirs.rin_fsv3004)
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Tuple

# =============================================================================
# 项目根数据目录 —— 所有输出子目录从此派生，迁移项目只需改这一处
# =============================================================================
_PROJECT_ROOT = Path(r"C:\PTS\zhongzi")


# ========================== 仪器 IP 地址 =====================================

@dataclass(frozen=True)
class NetworkInstruments:
    """所有 LAN 连接仪器的 IP 地址。"""
    # 频谱分析仪
    fsv3004_rin:       str = "192.168.7.10"   # Rin_FSV3004, LineWidth_FSV3004
    fsv3004_linewidth: str = "192.168.7.10"   # LineWidth_FSV3004（同一台设备，别名）
    single_freq_sa:    str = "192.168.7.15"   # SingleFrequency 频谱仪

    # 信号发生器
    sig_gen_linewidth:  str = "192.168.7.11"  # LineWidth 信号发生器
    sig_gen_wavelength: str = "192.168.7.11"  # WaveLength 信号发生器（同一台）
    sig_gen_timedomain: str = "192.168.7.13"  # TimeDomain 信号发生器

    # 示波器
    scope:              str = "192.168.7.12"  # TimeDomain 示波器

    # 光谱仪
    osa:                str = "192.168.7.14"  # SpectrumSNR 光谱仪



# ========================== USB / VISA 资源地址 ==============================

@dataclass(frozen=True)
class USBDevices:
    """USB VISA 资源字符串。"""
    optical_switch: str = "USB0::0x0005::0x0012::87104000113::INSTR"
    power_meter:    str = ""  # 功率计 USB 地址（因设备而异，默认为空，需用户在 GUI 中填写）


# ========================== 串口配置 =========================================

@dataclass(frozen=True)
class SerialPortConfig:
    """DFB 激光器 RS-485 串口通信参数。"""
    port_default: str   = "COM1"
    baudrate:     int   = 115200
    timeout_s:    float = 1.5
    device_addr:  int   = 100
    # 协议帧字节
    frame_head1:  int   = 0x50
    frame_head2:  int   = 0x00
    frame_tail1:  int   = 0x0D
    frame_tail2:  int   = 0x0A
    # 串口扫描无结果时的回退列表
    fallback_ports: Tuple[str, ...] = ("COM1", "COM2", "COM3", "COM4")


# ========================== 数据目录 =========================================

@dataclass(frozen=True)
class DataDirectories:
    """项目输出数据目录，均以 _PROJECT_ROOT 为基准派生。"""
    root:          Path = _PROJECT_ROOT
    power:         Path = field(default_factory=lambda: _PROJECT_ROOT / "Power")
    wavelength:    Path = field(default_factory=lambda: _PROJECT_ROOT / "WaveLength")
    spectrum_snr:  Path = field(default_factory=lambda: _PROJECT_ROOT / "SpectrumSNR")
    rin_fsv3004:   Path = field(default_factory=lambda: _PROJECT_ROOT / "Rin" / "FSV3004")
    line_width:    Path = field(default_factory=lambda: _PROJECT_ROOT / "LineWidth")
    single_freq_1um:   Path = field(default_factory=lambda: _PROJECT_ROOT / "SingleFrequency" / "1.0μm")
    single_freq_1_5um: Path = field(default_factory=lambda: _PROJECT_ROOT / "SingleFrequency" / "1.5μm")
    time_domain:   Path = field(default_factory=lambda: _PROJECT_ROOT / "TimeDomain")
    seed_value:    Path = field(default_factory=lambda: _PROJECT_ROOT / "SeedValue")
    report:        Path = Path(r"C:\PTS\report")


# ========================== 系统路径（DLL、字体等） ==========================

@dataclass(frozen=True)
class SystemPaths:
    """外部系统文件路径。"""
    wlm_dll:        str = r"C:\Windows\System32\wlmData.dll"
    arial_font:     str = r"C:\Windows\Fonts\arial.ttf"
    instrument_temp_png: str = r"C:\PTS\zhongzi\LineWidth\_temp.png"
    rin_temp_png:   str = r"C:\PTS\Rin\_temp.png"


# ========================== RIN 测量参数 =====================================

@dataclass(frozen=True)
class RinMeasurementParams:
    """RIN 测量：6 段频段配置表 + DC / 放大常数。

    每段：(start_hz, stop_hz, bandwidth_hz, avg_count, filename)
    """
    segments: Tuple[Tuple[int, int, int, int, str], ...] = (
        (      10,       100,  5, 20, "Rin_1.DAT"),
        (     100,      1000,  5, 20, "Rin_2.DAT"),
        (    1000,     10000, 30, 20, "Rin_3.DAT"),
        (   10000,    100000, 30, 20, "Rin_4.DAT"),
        (  100000,   1000000, 30, 20, "Rin_5.DAT"),
        ( 1000000,  10000000, 30, 20, "Rin_6.DAT"),
    )
    sweep_points:      int   = 2001
    dc_internal:        float = 1.20   # RinTest 默认 DC（公式用）
    dc_initial:         float = 2.40   # GUI 显示默认 DC
    amplification:      int   = 14
    # 弛豫振荡峰值搜索
    relax_start_hz:     int   = 100_000
    relax_stop_hz:      int   = 3_000_000
    target_freqs:       Tuple[int, ...] = (1000, 10000, 100000, 1000000)
    peak_min_points:    int   = 3


@dataclass(frozen=True)
class BackgroundNoiseParams:
    """背景噪声 / 种子光测量默认值。"""
    start_freq:  int = 10
    stop_freq:   int = 100_000_000
    bandwidth:   int = 30
    avg_count:   int = 1


# ========================== 线宽测量参数 =====================================

@dataclass(frozen=True)
class LineWidthParams:
    """线宽测量默认值。"""
    span_values:    Tuple[str, ...] = ("100", "200", "500", "1000", "2000")  # kHz
    rbw_hz:         int   = 100
    ref_level_mv:   int   = 1000
    m1_position_mhz: int  = 80
    center_freq_mhz: int  = 80
    n_db_down:      int   = 20
    sweep_points:   int   = 2001
    avg_count:      int   = 20
    # 信号发生器额外测试
    extra_span:     str   = "500"
    extra_freq:     float = 0.1
    extra_volt:     float = 0.0
    extra_offset:   float = 1.0
    extra_suffix:   str   = "500+1v"
    # 连接
    connect_max_retries: int = 3
    connect_retry_interval_s: int = 2
    ndbdown_read_timeout_ms: int = 5000
    save_data_max_retries: int = 2
    reconnect_delay_s: float = 2.0
    signal_stabilize_s: float = 1.0


# ========================== 信噪比测量参数 ===================================

@dataclass(frozen=True)
class SpectrumSNRParams:
    """频谱 SNR 测量默认值。"""
    center_nm:              float = 1064.0
    span_nm:                float = 150.0
    ref_level_dbm:          float = -4.0
    sensitivity:            str   = "HIGH1"
    main_peak_exclusion_nm: float = 3.0
    visa_timeout_s:         int   = 120
    screenshot_timeout_ms:  int   = 180000
    query_retries:          int   = 3
    query_retry_delay_s:    float = 0.4


# ========================== 单频测量参数 =====================================

@dataclass(frozen=True)
class SingleFrequencyParams:
    """单频测试 — 1μm / 1.5μm 双预设 + 扫描 / 峰值检测参数。"""
    ip_address:         str   = "192.168.7.15"
    serial_port:        str   = "COM1"
    test_duration_min:  float = 30.0

    # ---- 1μm 预设 ----
    temp_max_1um:       float = 60.0
    temp_min_1um:       float = 18.0
    temp_step_1um:      float = 0.1
    temp_interval_1um:  float = 3.0
    cur_max_1um:        float = 600.0
    cur_min_1um:        float = 100.0
    cur_step_1um:       float = 5.0
    cur_interval_1um:   float = 5.0

    # ---- 1.5μm 预设 ----
    temp_max_1_5um:     float = 56.0
    temp_min_1_5um:     float = 20.0
    temp_step_1_5um:    float = 0.1
    temp_interval_1_5um:float = 3.0
    cur_max_1_5um:      float = 1400.0
    cur_min_1_5um:      float = 1400.0
    cur_step_1_5um:     float = 0.0
    cur_interval_1_5um: float = 0.0

    # ---- 扫描 / 峰值检测 ----
    sweep_span_mhz:     float = 500.0
    sweep_step_mhz:     float = 500.0
    sweep_freq_max_mhz: float = 18000.0
    rbw_khz:            float = 30.0
    avg_count:          int   = 2
    sweep_time_s:       float = 1.0
    fine_repeat_count:  int   = 2
    peak_threshold_db:  float = 3.0
    peak_prominence_db: float = 5.0
    peak_guard_points:  int   = 10

    # ---- 波长稳定等待 ----
    stability_tol_nm:       float = 0.001
    stability_consecutive:  int   = 3
    stability_max_wait_s:   float = 300.0
    stability_interval_s:   float = 0.2

    # ---- 超时 ----
    sa_timeout_s:           float = 60.0
    query_opc_timeout_s:    float = 5.0
    sweep_timeout_ms:       int   = 15000


# ========================== 时域测量参数 =====================================

@dataclass(frozen=True)
class TimeDomainParams:
    """时域测量默认值。"""
    test_frequencies:   Tuple[int, ...] = (100, 300)
    gen_volt:           float = 10.0
    gen_offset:         float = 5.0
    gen_channel:        int   = 2
    scope_channel:      str   = "CHAN1"
    # Vpp→放大倍数映射: (min_mv, max_mv) → factor
    scale_map: Tuple[Tuple[Tuple[int, int], int], ...] = (
        ((0, 50), 64),
        ((50, 100), 32),
        ((100, 200), 16),
        ((200, 400), 8),
        ((400, 800), 4),
        ((800, 1600), 2),
        ((1600, float('inf')), 1),
    )
    vpp_stable_count:       int   = 5
    vpp_stable_delay_s:     float = 0.5
    vpp_min_valid:          float = 0.01
    vpp_max_valid:          float = 10.0
    vpp_default:            float = 0.1
    final_vpp_count:        int   = 3
    final_vpp_delay_s:      float = 0.3
    meas_retries:           int   = 5
    meas_retry_delay_s:     float = 0.8
    signal_stabilize_s:     float = 3.0
    gen_settle_s:           float = 1.5
    invalid_sentinels:      Tuple[str, ...] = ("9.91E+37", "NAN")


# ========================== PZT / 波长调制参数 ===============================

@dataclass(frozen=True)
class WaveLengthParams:
    """PZT 调制测试 — 1μm / 1.5μm 双预设 + DLL / 时序。"""
    sig_gen_ip:         str = "192.168.7.11"

    # ---- 1μm 预设 ----
    freq_1um:           str = "1"
    volt_1um:           str = "2"
    offset_1um:         str = "100"

    # ---- 1.5μm 预设 ----
    freq_1_5um:         str = "1.5"
    volt_1_5um:         str = "4"
    offset_1_5um:       str = "100"

    # ---- 变参测试（共用） ----
    sweep_text:         str = (
        "0.1Hz, 1Vpp, 0.5Vdc\n"
        "0.1Hz, 2Vpp, 1Vdc\n"
        "0.1Hz, 5Vpp, 2.5Vdc\n"
        "0.1Hz, 10Vpp, 5Vdc"
    )
    modulation_time_s:  str = "100"

    # ---- 信号发生器默认 ----
    sig_gen_default_waveform: str  = "RAMP"
    sig_gen_default_freq:     float = 0.1
    sig_gen_default_volt:     float = 10.0
    sig_gen_default_offset:   float = 5.0
    sig_gen_connect_retries:  int   = 3
    sig_gen_connect_interval: int   = 2

    # ---- 波长计 ----
    exposure_ms:              int   = 2
    wait_event_timeout_ms:    int   = 1000
    # WlmController 实例化参数 (RFC, Mode, P1, P2)
    instantiate_rfc:  int = 0
    instantiate_mode: int = 1
    instantiate_p1:   int = 0
    instantiate_p2:   int = 0

    # ---- pywinauto 坐标 / 搜索 ----
    wlm_window_title:          str = ".*Wavelength Meter.*"
    longterm_graph_title:      str = r".*WLM LongTerm graph.*"
    pywinauto_backend:         str = "win32"
    reset_btn_offset_x:        int = 40
    top_right_arrow_offset_x:  int = 12
    top_right_arrow_offset_y:  int = 12
    close_panel_offset_x:      int = 12
    close_panel_offset_y:      int = 12
    top_chart_found_index:     int = 1
    bottom_chart_found_index:  int = 0
    reset_panel_found_index:   int = 1
    chart_arrow_found_index:   int = 1

    # ---- 输出文件名 ----
    pzt_range_filename:       str = "PZT_range.csv"
    wavelength_csv_filename:  str = "wavelength.csv"
    wavelength_screenshot:    str = "波长.png"
    initial_signal_screenshot: str = "初始信号.png"


# ========================== 功率测量参数 =====================================

@dataclass(frozen=True)
class PowerParams:
    """功率计测量默认值。"""
    candidate_commands: Tuple[str, ...] = (
        "READ?", "MEAS:POW?", "POW:READ?", "READ:POWER?", "READ:POW?"
    )
    burnin_header_lines: int  = 14
    burnin_power_column: str  = "Power (W)"
    output_filename:     str  = "power.csv"
    burnin_result_filename: str = "burnin_result.csv"
    fallback_save_dir:   str  = "."


# ========================== 相噪测量参数 =====================================

@dataclass(frozen=True)
class PhaseNoiseParams:
    """相位噪声测试默认值。"""
    program_1um:            str = ""
    program_1_5um:          str = ""
    wavelength_file:        str = r"C:\PTS\zhongzi\WaveLength\wavelength.csv"
    wavelength_threshold_nm: float = 1500.0
    click_retries:          int   = 3
    click_delay_s:          float = 0.5
    launch_wait_s:          float = 3.0
    post_launch_wait_s:     float = 2.0
    click_retry_delay_s:    float = 1.0
    check_interval_s:       float = 0.5
    click_coords:           Tuple[int, int] = (296, 675)
    target_window_class:    str   = "SunAwtFrame"
    target_window_title:    str   = "LaserNoiseMeasurement_Manufactory"


# ========================== 时序参数 =========================================

@dataclass(frozen=True)
class TimingConfig:
    """全局时序常量（timeout、delay、retry）。"""
    # VISA 超时
    visa_short_ms:          int   = 5000
    visa_default_ms:        int   = 10000
    visa_medium_ms:         int   = 60000
    visa_long_ms:           int   = 120000
    visa_extra_long_ms:     int   = 180000

    # 自动关窗
    auto_close_short_ms:    int   = 2000
    auto_close_long_ms:     int   = 3000

    # 信号稳定等待
    stabilize_short_s:      float = 0.1
    stabilize_hover_s:      float = 0.1
    stabilize_0_2s:         float = 0.2
    stabilize_0_3s:         float = 0.3
    stabilize_0_5s:         float = 0.5
    stabilize_1s:           float = 1.0
    stabilize_1_5s:         float = 1.5
    stabilize_2s:           float = 2.0
    stabilize_3s:           float = 3.0

    # 通道切换
    channel_switch_delay_s: float = 1.5

    # 重试
    default_retries:        int   = 3
    default_retry_interval: int   = 2

    # OPC
    opc_timeout_ms:         int   = 15000

    # GUI 轮询
    gui_realtime_poll_ms:   int   = 1500

    # 波长计
    wlm_connect_timeout_s:  float = 5.0
    wlm_start_timeout_s:    float = 15.0
    wlm_wait_retries:       int   = 5
    wlm_wait_interval_s:    float = 0.5
    wlm_popup_render_s:     float = 1.0
    wlm_post_read_delay_s:  float = 1.0


# ========================== UI 常量 ==========================================

@dataclass(frozen=True)
class UIConfig:
    """独立模块窗口几何、颜色、字体、matplotlib 默认值。
    注意：main_platform 的主题常量在 utils/theme.py，不在此处。
    """
    # 窗口几何
    geo_power:          str = "1275x510"
    geo_phase_noise:    str = "1140x420"
    geo_spectrum_snr:   str = "1250x370"
    geo_rin:            str = "1170x630"
    geo_rin_center_h:   int = 330
    geo_linewidth:      str = "1150x550"
    geo_single_freq:    str = "1320x950"
    geo_time_domain:    str = "1320x420"
    geo_wavelength:     str = "1200x820"
    geo_result_sel:     str = "2100x1300"
    geo_image_popup:    str = "1800x1600"

    # 按钮颜色
    btn_green:        str = "#4CAF50"
    btn_blue:         str = "#1D74C0"
    btn_orange:       str = "#f4a236"
    btn_red:          str = "#f44336"
    btn_dark_green:   str = "#28862B"
    btn_amber:        str = "#FFA000"
    btn_orange2:      str = "#FF9800"
    btn_browse_bg:    str = "#F0F0F0"
    btn_white_fg:     str = "#FFFFFF"

    # 字体
    font_arial:            str = "Arial"
    font_times_new_roman:  str = "Times New Roman"
    font_simhei:           str = "SimHei"
    font_consolas:         str = "Consolas"
    font_ms_yahei:         str = "Microsoft YaHei"

    # matplotlib 默认
    mpl_figsize_rin:        Tuple[int, int] = (10, 8)
    mpl_dpi:                int = 300
    mpl_dpi_high:           int = 600
    mpl_curve_color:        str = "#085cab"
    mpl_curve_linewidth:    float = 2.0
    mpl_spine_linewidth:    float = 2.5
    mpl_figsize_single:     Tuple[int, int] = (12, 6)
    mpl_linewidth_single:   float = 1.2
    mpl_color_single:       str = "yellow"


# ========================== 顶层配置聚合 =====================================

@dataclass(frozen=True)
class Config:
    """根配置对象 —— 一行导入获取所有配置。"""
    network:      NetworkInstruments   = field(default_factory=NetworkInstruments)
    usb:          USBDevices           = field(default_factory=USBDevices)
    serial:       SerialPortConfig     = field(default_factory=SerialPortConfig)
    dirs:         DataDirectories      = field(default_factory=DataDirectories)
    sys_paths:    SystemPaths          = field(default_factory=SystemPaths)
    rin:          RinMeasurementParams = field(default_factory=RinMeasurementParams)
    bg_noise:     BackgroundNoiseParams = field(default_factory=BackgroundNoiseParams)
    linewidth:    LineWidthParams      = field(default_factory=LineWidthParams)
    spectrum_snr: SpectrumSNRParams    = field(default_factory=SpectrumSNRParams)
    single_freq:  SingleFrequencyParams = field(default_factory=SingleFrequencyParams)
    time_domain:  TimeDomainParams     = field(default_factory=TimeDomainParams)
    wavelength:   WaveLengthParams     = field(default_factory=WaveLengthParams)
    power:        PowerParams          = field(default_factory=PowerParams)
    phase_noise:  PhaseNoiseParams     = field(default_factory=PhaseNoiseParams)
    timing:       TimingConfig         = field(default_factory=TimingConfig)
    ui:           UIConfig             = field(default_factory=UIConfig)


# 模块级单例
CFG = Config()
