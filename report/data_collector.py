"""
报告数据采集模块

负责从各个测试子系统的输出文件中读取数据，
组装成模板渲染所需的字典。

数据来源:
    - Rin/         相对强度噪声
    - WaveLength/   PZT 波长计
    - SpectrumSNR/  光谱信噪比
    - LineWidth/    线宽
    - Power/        输出功率 & 功率稳定性
"""

import csv
import os
import time


def _read_csv_cell(filepath, row_idx, col_idx, cast=float, fallback=""):
    """
    从 CSV 文件中读取指定单元格的值。

    :param filepath: CSV 文件的绝对路径
    :param row_idx:  行索引 (0-based)
    :param col_idx:  列索引 (0-based)
    :param cast:     类型转换函数, 默认 float
    :param fallback: 文件不存在或读取失败时的回退文本
    :return:         转换后的值, 或回退文本
    """
    if not os.path.exists(filepath):
        return fallback
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            rows = list(csv.reader(f))
            if row_idx < len(rows) and col_idx < len(rows[row_idx]):
                return cast(rows[row_idx][col_idx])
            return fallback
    except (OSError, ValueError, IndexError):
        return fallback

def _get_waverange():
    """
    判断波长调谐范围
    """
    wavelength = _read_csv_cell(
        r"C:\PTS\zhongzi\SeedValue\seedvalue.csv",
        row_idx=1, col_idx=4, cast=float, fallback="(未读取到电流数据)",
    )
    if wavelength < 1200: return f"{wavelength-0.35:.2f} ~ {wavelength+0.35:.2f}"
    else: return f"{wavelength-0.5:.2f} ~ {wavelength+0.5:.2f}"

def _get_hot_waverange():
    """
    判断波长热调谐范围
    """
    wavelength = _read_csv_cell(
        r"C:\PTS\zhongzi\SeedValue\seedvalue.csv",
        row_idx=1, col_idx=4, cast=float, fallback="未读取到电流数据",
    )
    if wavelength < 1200: return 0.7
    else: return 1

def _check_single_freq_result():
    """
    检查单频测试输出目录是否有文件。
    测试开始时清空输出目录，只有扫出单频（检测到峰值）才会写入 CSV/PNG。
    因此输出目录为空 → "无"，非空 → "有"。
    """
    from core.config import CFG
    dirs = [str(CFG.dirs.single_freq_1um), str(CFG.dirs.single_freq_1_5um)]
    for d in dirs:
        if os.path.isdir(d):
            try:
                if any(os.path.isfile(os.path.join(d, f)) for f in os.listdir(d)):
                    return "有"
            except OSError:
                pass
    return "无"


def _get_time_domain_vpp_values():
    """
    从 vpp.csv 读取各频率对应的 Vpp 值 (mV)。

    文件格式:
        频率(Hz),Vpp(mV)
        100,523.00
        300,487.00

    :return: dict, {freq_str: vpp_mv_float}，如 {"100": 523.0, "300": 487.0}
             文件不存在或读取失败时返回空 dict
    """
    vpp_csv = r"C:\PTS\zhongzi\TimeDomain\vpp.csv"
    if not os.path.exists(vpp_csv):
        return {}
    try:
        with open(vpp_csv, "r", encoding="utf-8") as f:
            rows = list(csv.reader(f))
            result = {}
            for row in rows[1:]:  # 跳过表头
                if len(row) >= 2:
                    result[row[0]] = float(row[1])
            return result
    except (OSError, ValueError, IndexError):
        return {}


def assemble_report_data():
    """
    组装报告模板所需的全部数据。

    键名必须与 Word 模板中的 {{ 变量名 }} 完全一致。

    :return: dict, 可直接传给 template_generator.generate_report()
    """
    return {
        # ---- 项目信息 ----
        "text_date": time.strftime("%Y/%m/%d"),

        # ---- 设备信息（由用户在测试结果弹窗中填入）----
        "text_model":  "",
        "text_number": "",
        "text_serial": "",

        # ---- 测试数据 ----
        "text_current": _read_csv_cell(
            r"C:\PTS\zhongzi\SeedValue\seedvalue.csv",
            row_idx=1, col_idx=8, cast=float, fallback="(未读取到电流数据)",
        ),

        "text_temp": f"{_read_csv_cell(
            r"C:\PTS\zhongzi\SeedValue\seedvalue.csv",
            row_idx=1, col_idx=6, cast=float, fallback=0,
        ):.2f}",

        "temp_min": _read_csv_cell(
            r"C:\PTS\zhongzi\SeedValue\seedvalue.csv",
            row_idx=1, col_idx=1, cast=float, fallback="(未读取到温度下限)",
        ),

        "temp_max": _read_csv_cell(
            r"C:\PTS\zhongzi\SeedValue\seedvalue.csv",
            row_idx=1, col_idx=0, cast=float, fallback="(未读取到温度上限)",
        ),

        "text_waverange": _get_waverange(),

        "text_hot_waverange": _get_hot_waverange(),


        # ---- Fig.1 相对强度噪声 (Rin) ----
        "img_rin": r"C:\PTS\zhongzi\Rin\FSV3004\Rin.png",
        "text_rin": _read_csv_cell(
            r"C:\PTS\zhongzi\Rin\FSV3004\rin_figure2_max.csv",
            row_idx=0, col_idx=0, cast=float, fallback="(未读取到积分数据)",
        ),

        # ---- Fig.2 PZT（波长计） ----
        "img_PZT": r"C:\PTS\zhongzi\WaveLength\10v.png",
        "text_PZT_range": _read_csv_cell(
            r"C:\PTS\zhongzi\WaveLength\PZT_range.csv",
            row_idx=1, col_idx=2, cast=float, fallback="(未读取到PZT数据)",
        ),
        "text_center_wavelength": _read_csv_cell(
            r"C:\PTS\zhongzi\WaveLength\wavelength.csv",
            row_idx=0, col_idx=0, cast=float, fallback="(未读取到波长数据)",
        ),

        # ---- Fig.3 光谱信噪比 ----
        "img_spectrumSNR": r"C:\PTS\zhongzi\SpectrumSNR\spectrum.bmp",
        "text_SNR": _read_csv_cell(
            r"C:\PTS\zhongzi\SpectrumSNR\spectrum_snr.csv",
            row_idx=0, col_idx=0, cast=float, fallback="(未读取到SNR数据)",
        ),

        # ---- Fig.4 线宽 ----
        "img_linewidth": r"C:\PTS\zhongzi\LineWidth\image_1000.png",
        "text_linewidth": _read_csv_cell(
            r"C:\PTS\zhongzi\LineWidth\ndbdown.csv",
            row_idx=1, col_idx=0, cast=float, fallback="(未读取到Ndbdown数据)",
        ),

        # ---- 单频 ----
        "text_single_freq": _check_single_freq_result(),

        # ---- 时域 ----
        "text_time_domain1": _get_time_domain_vpp_values().get("100", "(未读取到时域数据)"),
        "text_time_domain2": _get_time_domain_vpp_values().get("300", "(未读取到时域数据)"),

        # ---- Fig.5 偏振测试 ----
        "img_polarization": "暂无",

        # ---- 输出功率 ----
        "text_power": _read_csv_cell(
            r"C:\PTS\zhongzi\Power\power.csv",
            row_idx=0, col_idx=0, cast=lambda v: f"{float(v) / 0.03:.2f}", fallback="(未读取到功率数据)",
        ),

        # ---- 功率稳定性 ----
        "img_power_stability": "烤机数据画图",
        "text_RMS": _read_csv_cell(
            r"C:\PTS\zhongzi\Power\burnin_result.csv",
            row_idx=1, col_idx=1, cast=float, fallback="(未读取到RMS数据)",
        ),
        "text_P2P": "读取整列数据 （最大值-最小值）/平均值",
    }
