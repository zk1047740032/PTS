"""Headless RIN 测试服务 — 复用仪器层 + 测试逻辑，去掉 tkinter 依赖"""

import csv
import os
import sys
import threading
import traceback
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import matplotlib
matplotlib.use('Agg')  # 非交互后端，无需 tkinter
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

# 确保能 import core
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from core import clear_directory, write_xy_csv, read_instrument_screenshot
from core.config import CFG


# =====================================================================
# 复用原 FSV3004 仪器控制（直接 import path_a 文件里的类）
# =====================================================================
def _import_fsv3004():
    """动态导入 FSV3004Instrument，避免顶层 tkinter 导入冲突"""
    import importlib

    # 阻止原脚本的 matplotlib.use('TkAgg') 执行（我们用 Agg 后端）
    _original_use = matplotlib.use
    matplotlib.use = lambda backend: None

    # 抑制原脚本的 tkinter 依赖：提供虚拟的 tk/messagebox/filedialog
    sys.modules.setdefault('tkinter', _FakeTk())
    sys.modules.setdefault('tkinter.messagebox', _FakeMessageBox())
    sys.modules.setdefault('tkinter.filedialog', _FakeModule())
    sys.modules.setdefault('tkinter.simpledialog', _FakeModule())
    sys.modules.setdefault('PIL.Image', _FakeModule())
    sys.modules.setdefault('PIL.ImageTk', _FakeModule())
    sys.modules.setdefault('matplotlib.backends.backend_tkagg', _FakeModule())

    spec = importlib.util.spec_from_file_location(
        "Rin_FSV3004_module",
        _PROJECT_ROOT / "path_a" / "Rin_FSV3004.py",
    )
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
    finally:
        matplotlib.use = _original_use  # 恢复

    return mod.FSV3004Instrument


class _FakeTk:
    """虚拟 tkinter 模块，拦截 tk.Toplevel / tk.Button 等创建"""
    Toplevel = type('_FakeToplevel', (), {'title': lambda *a: None, 'winfo_exists': lambda *a: False})()
    Frame = type('_FakeFrame', (), {})()
    LabelFrame = type('_FakeLabelFrame', (), {})()
    Label = type('_FakeLabel', (), {})()
    Button = type('_FakeButton', (), {})()
    Entry = type('_FakeEntry', (), {})()
    Text = type('_FakeText', (), {})()
    END = 'end'
    DISABLED = 'disabled'
    NORMAL = 'normal'
    WORD = 'word'
    TOP = 'top'
    LEFT = 'left'
    RIGHT = 'right'
    BOTTOM = 'bottom'
    BOTH = 'both'
    X = 'x'
    Y = 'y'
    HORIZONTAL = 'horizontal'
    VERTICAL = 'vertical'
    NSEW = 'nsew'
    StringVar = lambda *a: type('_FakeSV', (), {'get': lambda: '', 'set': lambda v: None})()
    BooleanVar = lambda *a: type('_FakeBV', (), {'get': lambda: False, 'set': lambda v: None})()


class _FakeMessageBox:
    showinfo = lambda *a, **kw: None
    showwarning = lambda *a, **kw: None
    showerror = lambda *a, **kw: None
    askyesno = lambda *a, **kw: False
    askfloat = lambda *a, **kw: None


class _FakeModule:
    """通用假模块，属性访问返回无操作函数"""
    def __getattr__(self, name):
        return lambda *a, **kw: None


# =====================================================================
# Headless RIN 测试运行器
# =====================================================================
class HeadlessRinRunner:
    """
    RIN 测试的 headless 版本。
    复用 path_a/Rin_FSV3004.py 的 FSV3004Instrument 和数据处理逻辑，
    去掉所有 tkinter GUI 依赖，改为通过回调推送进度 + 直接保存结果。
    """

    def __init__(
        self,
        dc_value: float = 2.4,
        ip_address: str | None = None,
        save_path: str | None = None,
        progress_callback=None,
        stop_flag: threading.Event | None = None,
    ):
        self.dc_value = dc_value / 2.0  # DC 输入值 → 内部使用值
        self.ip_address = ip_address or CFG.network.fsv3004_rin
        self.save_path = save_path or str(CFG.dirs.rin_fsv3004)
        self.progress_callback = progress_callback or (lambda msg, data=None: None)
        self.stop_flag = stop_flag or threading.Event()
        self.instrument = None

        # 数据处理中间结果
        self.dx: list = []
        self.dy: list = []
        self.ddx: list = []
        self.ddy: list = []
        self.RIN_power: list = []
        self.file_paths: list[str] = [
            str(CFG.dirs.rin_fsv3004 / seg[4]) for seg in CFG.rin.segments
        ]

        # 结果
        self.result: dict = {}
        self.rin_png_path: str = ""

    # -----------------------------------------------------------------
    # 主流程
    # -----------------------------------------------------------------
    def run(self):
        """执行完整 RIN 测量流程"""
        try:
            self._log("正在导入仪器控制模块...")
            FSV3004Instrument = _import_fsv3004()

            # 1. 清理输出目录
            self._log("[初始化] 正在清理数据文件夹...")
            local_dir = self.save_path
            os.makedirs(local_dir, exist_ok=True)
            clear_directory(local_dir, log_func=self._log_static)
            self._log("[初始化] 文件夹清理完成")

            # 2. 连接仪器
            self._log(f"[连接] 正在连接 FSV3004 ({self.ip_address})...")
            self.instrument = FSV3004Instrument(log_func=self._log)
            ok = self.instrument.connect(ip_address=self.ip_address)
            if not ok:
                self._log("[错误] 无法连接仪器，测试终止")
                self.result = {"status": "error", "error": "连接仪器失败"}
                return self.result
            self._log("[连接] 成功连接 FSV3004")

            # 3. 配置仪器
            self._log("[配置] 正在配置 RIN 测量模式...")
            self.instrument.configure_rin_mode()

            # 4. 6 段测量
            self._log("[测试] 开始 6 段测量...")
            measurement_params = CFG.rin.segments
            for idx, (start, stop, bw, avg, fname) in enumerate(measurement_params):
                if self.stop_flag.is_set():
                    self._log("[测试] 用户终止测试")
                    self.result = {"status": "stopped"}
                    return self.result

                self._log(f"[测试] 段 {idx+1}/6: {start}Hz - {stop}Hz, BW={bw}Hz")
                self._send_progress("measure", {
                    "segment": idx + 1, "total_segments": 6,
                    "start_freq": start, "stop_freq": stop,
                    "progress": (idx) / 6 * 100,
                })

                freqs, amps = self.instrument.measure_rin_segment(start, stop, bw, avg)
                if freqs is None:
                    self._log(f"[错误] 段 {idx+1} 测量失败")
                    continue

                filepath = os.path.join(self.save_path, fname)
                with open(filepath, 'w', newline='') as f:
                    writer = csv.writer(f)
                    for freq, amp in zip(freqs, amps):
                        writer.writerow([freq, amp])
                self._log(f"[保存] {fname}")

            # 5. 关闭仪器（在数据处理前关闭以释放资源）
            try:
                self.instrument.close()
            except Exception:
                pass
            self._log("[仪器] 已断开连接")

            if self.stop_flag.is_set():
                self.result = {"status": "stopped"}
                return self.result

            # 6. 数据处理
            self._log("[处理] 正在分析数据...")
            self._send_progress("analyze", {"progress": 80})
            self._process_files()
            self._send_progress("analyze", {"progress": 90})

            # 7. 生成图表
            self._log("[可视化] 正在生成 RIN 图表...")
            self._send_progress("visualize", {"progress": 95})
            self._visualize_headless()

            # 8. 组装结果
            self._log("[完成] RIN 测试完成")
            self._send_progress("complete", {"progress": 100})

            self.result = {
                "status": "completed",
                "rin_png": self.rin_png_path,
                "peak_info": self._result_peaks,
                "target_points": self._result_points,
                "integrated_rms_max": self._result_rms_max,
                "output_dir": self.save_path,
            }
            return self.result

        except Exception as e:
            self._log(f"[异常] {e}\n{traceback.format_exc()}")
            self.result = {"status": "error", "error": str(e)}
            return self.result
        finally:
            if self.instrument:
                try:
                    self.instrument.close()
                except Exception:
                    pass

    # -----------------------------------------------------------------
    # 数据处理（原 RinAnalyzer 逻辑）
    # -----------------------------------------------------------------
    def _read_csv(self, file_path):
        try:
            with open(file_path, 'r') as f:
                sample = f.read(1024)
                f.seek(0)
                dialect = csv.Sniffer().sniff(sample)
                reader = csv.reader(f, dialect)
                fx, fy = [], []
                for row in reader:
                    if len(row) >= 2:
                        try:
                            fx.append(float(row[0].replace(',', '.')))
                            fy.append(float(row[1].replace(',', '.')))
                        except ValueError:
                            continue
            if len(fy) < 2001:
                self._log(f"警告: 数据点数非2001，实际 {len(fy)}")
                return [], []
            return fx, fy
        except Exception as e:
            self._log(f"读取文件失败 {file_path}: {e}")
            return [], []

    def _process_files(self):
        self.dx, self.dy = [], []
        for fp in self.file_paths:
            if not (os.path.exists(fp) and os.path.getsize(fp) > 0):
                self._log(f"文件不存在: {fp}")
                self.dx.append([])
                self.dy.append([])
                continue
            fx, fy = self._read_csv(fp)
            self.dx.append(fx)
            self.dy.append(fy)
            self._log(f"读取: {os.path.basename(fp)} ({len(fy)} 点)")

        self.ddx, self.ddy = [], []
        amplification = CFG.rin.amplification
        dc_value = self.dc_value
        rows_per_file = 2001

        for j in range(len(self.dx)):
            if not self.dx[j]:
                continue
            for i in range(min(rows_per_file, len(self.dx[j]))):
                self.ddx.append(self.dx[j][i])
                scale_factor = np.sqrt(5) if j < 2 else np.sqrt(30)
                v_noise = self.dy[j][i]
                if v_noise <= 0:
                    self.ddy.append(float('-inf'))
                else:
                    rin_val = 20 * np.log10(v_noise / (dc_value * amplification * scale_factor))
                    self.ddy.append(rin_val)

        if self.ddx and self.ddy:
            self.RIN_power = self._compute_rin_power(self.ddx, self.ddy)

    def _compute_rin_power(self, x, y):
        power = []
        segment_length = 6
        for k in range(1, len(x) // segment_length + 1):
            sub_x = x[:k * segment_length]
            sub_y_exp = [np.power(10, v / 10.0) if np.isfinite(v) else 0 for v in y[:k * segment_length]]
            integral = sum(
                (sub_x[i] - sub_x[i - 1]) * (sub_y_exp[i] + sub_y_exp[i - 1]) / 2.0
                for i in range(1, len(sub_x))
            )
            power.append(np.sqrt(integral))
        return power

    # -----------------------------------------------------------------
    # 可视化（headless matplotlib — 不弹 tk 窗口）
    # -----------------------------------------------------------------
    def _visualize_headless(self):
        if not self.ddx or not self.ddy or not self.RIN_power:
            self._log("没有可视化数据")
            return

        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8),
            gridspec_kw={'height_ratios': [3, 2]})

        # 图1: RIN 曲线
        ax1.plot(self.ddx, self.ddy, color="#085cab", linewidth=2)
        ax1.set_xscale('log')
        ax1.margins(x=0)
        ax1.tick_params(axis='both', which='major', labelsize=20, pad=5, length=12, width=3, direction='in')
        ax1.set_ylabel('RIN (dBc/Hz)', fontsize=18, fontweight='bold')
        ax1.grid(True, which='both', lw=2, linestyle='--', alpha=1)
        for spine in ax1.spines.values():
            spine.set_linewidth(2.5)
        for label in ax1.get_xticklabels() + ax1.get_yticklabels():
            label.set_fontname('Times New Roman')
            label.set_fontsize(20)
            label.set_fontweight('bold')
        finite_ddy = [v for v in self.ddy if np.isfinite(v)]
        if finite_ddy:
            ax1.set_ylim(np.floor(min(finite_ddy) / 10) * 10, np.ceil(max(finite_ddy) / 10) * 10)
        ax1.yaxis.set_major_locator(MaxNLocator(nbins=6, integer=True))

        # 图2: RMS 积分曲线
        adjusted_power = [p * 100 for p in self.RIN_power]
        plot_len = len(adjusted_power)
        ax2.plot(self.ddx[::6][:plot_len], adjusted_power, color="#085cab", linewidth=2)
        ax2.set_xscale('log')
        ax2.margins(x=0)
        y2_min, y2_max = np.min(adjusted_power), np.max(adjusted_power)
        if np.isclose(y2_min, y2_max):
            y2_min -= 1
            y2_max += 1
        y2_mid = (y2_min + y2_max) / 2
        ax2.set_yticks([y2_min, y2_mid, y2_max])
        ax2.set_yticklabels([f"{y2_min:.3f}%", f"{y2_mid:.3f}%", f"{y2_max:.3f}%"])
        ax2.tick_params(axis='both', which='major', labelsize=15, pad=5, length=12, width=3, direction='in')
        ax2.set_xlabel('Frequency(Hz)', fontsize=18, fontweight='bold')
        ax2.set_ylabel('Integrated RMS', fontsize=18, fontweight='bold')
        ax2.grid(True, which='both', lw=2, linestyle='--', alpha=1)
        for spine in ax2.spines.values():
            spine.set_linewidth(2.5)
        for label in ax2.get_xticklabels() + ax2.get_yticklabels():
            label.set_fontname('Times New Roman')
            label.set_fontsize(20)
            label.set_fontweight('bold')

        plt.tight_layout()
        plt.subplots_adjust(hspace=0.15)

        # 保存 PNG
        png_path = os.path.join(self.save_path, "Rin.png")
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
        plt.close(fig)
        self.rin_png_path = png_path
        self._log(f"[保存] 图表已保存: {png_path}")

        # 保存图二 y 轴最大值
        max_path = os.path.join(self.save_path, "rin_figure2_max.csv")
        with open(max_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([f"{y2_max:.3f}"])
        self._result_rms_max = f"{y2_max:.3f}%"

        # 驰豫振荡峰检测
        self._detect_peaks()
        self._compute_target_points()

    def _detect_peaks(self):
        relax_start = CFG.rin.relax_start_hz
        relax_stop = CFG.rin.relax_stop_hz
        freqs = np.array(self.ddx)
        ys = np.array(self.ddy)
        mask = (freqs >= relax_start) & (freqs <= relax_stop) & np.isfinite(ys)
        highest_rin = float('nan')
        peak_freq = float('nan')

        if np.any(mask):
            masked_freqs = freqs[mask]
            masked_ys = ys[mask]
            if len(masked_ys) >= 3:
                center_vals = masked_ys[1:-1]
                is_peak = (center_vals > masked_ys[:-2]) & (center_vals > masked_ys[2:])
                if np.any(is_peak):
                    peak_indices = np.where(is_peak)[0] + 1
                    best_idx = peak_indices[np.argmax(masked_ys[peak_indices])]
                    highest_rin = float(masked_ys[best_idx])
                    peak_freq = float(masked_freqs[best_idx])
                else:
                    max_idx = np.argmax(masked_ys)
                    highest_rin = float(masked_ys[max_idx])
                    peak_freq = float(masked_freqs[max_idx])
            else:
                max_idx = np.argmax(masked_ys)
                highest_rin = float(masked_ys[max_idx])
                peak_freq = float(masked_freqs[max_idx])

        self._result_peaks = {
            "start_hz": relax_start,
            "stop_hz": relax_stop,
            "rin_dbc_hz": round(highest_rin, 3) if np.isfinite(highest_rin) else None,
            "freq_hz": round(peak_freq, 0) if np.isfinite(peak_freq) else None,
        }
        self._log(f"驰豫振荡峰 ({relax_start}-{relax_stop} Hz): "
                  f"{self._result_peaks['rin_dbc_hz']} dBc/Hz @ {self._result_peaks['freq_hz']} Hz")

    def _compute_target_points(self):
        points = []
        for tx in CFG.rin.target_freqs:
            if len(self.ddx) == 0:
                continue
            idx = np.argmin(np.abs(np.array(self.ddx) - tx))
            x_val = self.ddx[idx]
            y_val = self.ddy[idx]
            points.append({
                "freq_hz": round(x_val, 0),
                "rin_dbc_hz": round(y_val, 3) if np.isfinite(y_val) else None,
            })
        self._result_points = points

    # -----------------------------------------------------------------
    # 回调辅助
    # -----------------------------------------------------------------
    def _log(self, msg: str):
        self.progress_callback(msg, {"type": "log"})

    def _log_static(self, msg: str):
        """静态日志（被 clear_directory 等返回值为 None 的回调使用）"""
        self.progress_callback(msg, {"type": "log"})

    def _send_progress(self, phase: str, data: dict):
        self.progress_callback(f"阶段: {phase}", {"type": "progress", "phase": phase, **data})
