# 添加状态稳定检测
"""
CT Power Tuning GUI - PM100D 版本
- 保持原有 GUI、激光上位机控制与两组测试流程不变
- 不再使用光谱仪（OSA），改为通过 USB (VISA) 连接 Thorlabs PM100D 功率计（或兼容设备）
- 每步采集单点功率值 (W)，保存为 summary CSV，并绘制 温度/电流 vs 功率 曲线
Requirements:
    pip install pyvisa numpy matplotlib pillow pywinauto
Notes:
    - 在 GUI 的 "USB 资源地址" 中填写实际的 VISA 资源字符串，例如:
      USB0::0x1313::0x8078::PM100D_serial::INSTR
    - PM100D 常见读数命令尝试顺序: "READ?", "MEAS:POW?", "POW:READ?" 等，程序会做多种尝试以兼容不同固件。
"""
from __future__ import annotations

import os
import time
import threading
import csv
import traceback
from typing import Optional, Any, Dict, List, Tuple

import pyvisa
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog
import matplotlib
import matplotlib.ticker as mticker

# PIL (for image popup)
try:
    from PIL import Image, ImageTk
except Exception:
    raise ImportError("请先安装 Pillow: pip install pillow")

matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
# 设置中文字体
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题
# pywinauto 控制激光上位机（保留原逻辑）
try:
    from pywinauto.application import Application
    from pywinauto import timings
    PYW_AVAILABLE = True
except Exception:
    Application = None
    timings = None
    PYW_AVAILABLE = False

# -------------------------
# Helpers
# -------------------------
def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)
    return path

# -------------------------
# LaserController (复用原来实现)
# -------------------------
class LaserController:
    def __init__(self, exe_path: str = r"C:\PTS\CTTuning\Preci_Semi\Preci-Seed.exe",
                 window_title: str = r"Preci-Semi-Seed", log_func=print):
        self.exe_path = exe_path
        self.window_title = window_title
        self.app = None
        self.win = None
        self.log = log_func
        self.connected = False

    def connect(self, timeout: float = 10.0, attach_timeout: float = 3.0) -> bool:
        if not PYW_AVAILABLE:
            raise RuntimeError("pywinauto 未安装或不可用，无法控制激光器上位机。")
        if self.connected and self.win is not None:
            try:
                if self.win.exists() and self.win.is_visible():
                    self.log("[Laser] 已复用现有窗口句柄")
                    return True
            except Exception:
                self.log("[Laser] 现有句柄无效，重新连接")
                self.app = None
                self.win = None
                self.connected = False
        try:
            self.log("[Laser] 尝试附加到运行中的 Preci-Semi-Seed 窗口...")
            self.app = Application(backend="uia").connect(title_re=self.window_title, timeout=attach_timeout)
            self.win = self.app.window(title_re=self.window_title)
            timings.wait_until_passes(3, 0.5, lambda: self.win.exists() and self.win.is_visible())
            self.win.set_focus()
            self.connected = True
            self.log("[Laser] 附加成功")
            return True
        except Exception as e_attach:
            self.log(f"[Laser] 未找到运行实例: {e_attach}，尝试启动 exe：{self.exe_path}")
        try:
            self.log("[Laser] 启动 exe...")
            self.app = Application(backend="uia").start(cmd_line=f'"{self.exe_path}"')
            start_time = time.time()
            deadline = start_time + timeout
            while time.time() < deadline:
                try:
                    self.app.connect(title_re=self.window_title, timeout=1)
                    self.win = self.app.window(title_re=self.window_title)
                    timings.wait_until_passes(5, 0.5, lambda: self.win.exists() and self.win.is_visible())
                    self.win.set_focus()
                    self.connected = True
                    self.log("[Laser] 启动并连接成功")
                    return True
                except Exception:
                    time.sleep(0.3)
            raise RuntimeError("启动后未检测到窗口（超时）")
        except Exception as e_start:
            self.log(f"[Laser][错误] 启动或连接失败: {e_start}")
            self.connected = False
            raise

    def get_current_mA(self) -> Optional[float]:
        try:
            ctrl = self.win.child_window(auto_id="Label_current", control_type="Text")
            txt = ctrl.window_text()
            return float(txt)
        except Exception as e:
            self.log(f"[Laser] 读取电流失败: {e}")
            return None

    def set_current_mA(self, val_mA: float):
        try:
            edit = self.win.child_window(auto_id="textBox_Current", control_type="Edit")
            edit.set_edit_text(f"{val_mA:.2f}")
            btn = self.win.child_window(title="Set", control_type="Button")
            btn.click()
            self.log(f"[Laser] 已设置电流: {val_mA:.2f} mA")
        except Exception as e:
            self.log(f"[Laser] 设置电流失败: {e}")
            raise

    def get_temperature_C(self) -> Optional[float]:
        try:
            ctrl = self.win.child_window(auto_id="Label_Temperature", control_type="Text")
            txt = ctrl.window_text()
            return float(txt)
        except Exception as e:
            self.log(f"[Laser] 读取温度失败: {e}")
            return None

    def set_temperature_C(self, val_C: float):
        try:
            edit = self.win.child_window(auto_id="TextBox_Temperature", control_type="Edit")
            edit.set_edit_text(f"{val_C:.2f}")
            btn = self.win.child_window(title="Set", control_type="Button")
            btn.click()
            self.log(f"[Laser] 已设置温度: {val_C:.2f} °C")
        except Exception as e:
            self.log(f"[Laser] 设置温度失败: {e}")
            raise

# -------------------------
# PowerMeterController (针对 Thorlabs PM100D 等 USB 功率计)
# -------------------------
class PowerMeterController:
    """
    简单且健壮的功率计控制器。
    通过 pyvisa 打开 USB (VISA) 资源，并尝试读取单点功率（单位 W）。
    兼容性：为增加兼容性，会按顺序尝试多个常见的读数命令。
    """
    def __init__(self, resource: str, log_func=print, timeout_ms: int = 5000):
        self.rm = pyvisa.ResourceManager()
        self.inst = None
        self.resource = resource
        self.log = log_func
        self.timeout_ms = timeout_ms

    def connect(self):
        try:
            self.inst = self.rm.open_resource(self.resource)
            self.inst.timeout = int(self.timeout_ms)
            # 一些设备需要设置为 ASCII/readable format，但 PM100D 通常直接支持 READ?
            self.log(f"[PM] 已连接: {self.resource}")
        except Exception as e:
            self.log(f"[PM] 连接失败: {e}")
            raise

    def query_idn(self) -> str:
        try:
            return self.inst.query("*IDN?").strip()
        except Exception as e:
            self.log(f"[PM] *IDN? 失败: {e}")
            return ""

    def _try_query_float(self, cmd: str) -> Optional[float]:
        try:
            resp = self.inst.query(cmd).strip()
            if resp == "":
                return None
            # 清理常见格式并解析第一个浮点数
            token = resp.split()[0].replace(",", "")
            return float(token)
        except Exception as e:
            self.log(f"[PM] 命令 '{cmd}' 读取失败: {e}")
            return None

    def read_power(self) -> float:
        """
        尝试按优先级读取功率值（单位 W）。
        返回浮点功率（W），若失败抛出 RuntimeError。
        常见命令尝试顺序：
            READ?
            MEAS:POW?
            POW:READ?
            READ:POW?
        """
        if self.inst is None:
            raise RuntimeError("功率计未连接")
        candidates = ["READ?", "MEAS:POW?", "POW:READ?", "READ:POWER?", "READ:POW?"]
        last_errs = []
        for cmd in candidates:
            val = self._try_query_float(cmd)
            if val is not None:
                self.log(f"[PM] 命令 '{cmd}' 返回: {val} (W)")
                return float(val)
            else:
                last_errs.append(cmd)
        # 退回到 raw read（某些驱动下）
        try:
            raw = self.inst.read()
            if raw is not None and raw.strip() != "":
                try:
                    tok = raw.strip().split()[0].replace(",", "")
                    val = float(tok)
                    self.log(f"[PM] raw read 返回: {val} (W)")
                    return float(val)
                except Exception:
                    pass
        except Exception as e:
            self.log(f"[PM] raw read 失败: {e}")

        raise RuntimeError(f"无法从功率计读取功率，尝试的命令: {last_errs}")

# -------------------------
# TestRunner (改为读取功率并保存/绘图)
# -------------------------
class TestRunner:
    def __init__(self, laser: Optional[LaserController], pm: Optional[PowerMeterController], log_func=print):
        self.laser = laser
        self.pm = pm
        self.log = log_func
        self._stop = False

    def stop(self):
        self._stop = True
        self.log("[Runner] 停止信号已设置")

    def _float_range(self, start: float, stop: float, step: float) -> List[float]:
        if step == 0:
            raise ValueError("step cannot be 0")
        out = []
        t = start
        step_magnitude = abs(step)
        if start < stop:
            while t <= stop + 1e-9:
                out.append(round(t, 6))
                t += step_magnitude
        else:
            while t >= stop - 1e-9:
                out.append(round(t, 6))
                t -= step_magnitude
        return out

    def _append_summary(self, save_path: str, current_mA: float, temperature: Optional[float], power_w: float, test_group: int = 0, summary_filename: str = None):
        """
        汇总文件列： Current_mA, Temperature_C, Power_W
        """
        if os.path.isdir(save_path) or save_path.endswith(os.sep):
            out_dir = save_path
        else:
            out_dir = os.path.dirname(save_path) or "."
        ensure_dir(out_dir)

        if summary_filename:
            # 检查文件名是否已经包含.csv后缀，如果没有则添加
            if not summary_filename.lower().endswith('.csv'):
                summary_filename += '.csv'
            summary_fn = os.path.join(out_dir, summary_filename)
        elif test_group == 1:
            summary_fn = os.path.join(out_dir, "Test1_summary.csv")
        elif test_group == 2:
            summary_fn = os.path.join(out_dir, "Test2_summary.csv")
        else:
            summary_fn = os.path.join(out_dir, f"ct_power_summary_{time.strftime('%Y%m%d')}.csv")

        header_needed = not os.path.exists(summary_fn)
        with open(summary_fn, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if header_needed:
                w.writerow(["Current_mA", "Temperature_C", "Power_mW"])
            temp_str = f"{temperature:.2f}" if temperature is not None else "N/A"
            # 将W转换为mW（乘以1000）
            power_mw = float(power_w) * 1000
            w.writerow([f"{current_mA:.2f}", temp_str, f"{float(power_mw):.2f}"])
        self.log(f"[Runner] 汇总: {summary_fn} -> {current_mA:.2f} mA, {temp_str}, {power_mw:.2f} mW")

    def _plot_xy_curve(self, x, y, xlabel, ylabel, title, out_dir, prefix, invert_x=False, save_csv=False, extra_cols=None):
        ensure_dir(out_dir)
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        fig_path = os.path.join(out_dir, f"{prefix}_{timestamp}.png")

        plt.figure(figsize=(20, 10))
        plt.plot(x, y, marker='o', linestyle='-', linewidth=2)
        if invert_x:
            plt.gca().invert_xaxis()
        plt.xlabel(xlabel, fontsize=20)
        plt.ylabel(ylabel, fontsize=20)
        plt.title(title, fontsize=22)

        ax = plt.gca()
        ax.ticklabel_format(style='plain', axis='y')
        ax.yaxis.get_major_formatter().set_scientific(False)
        ax.yaxis.get_major_formatter().set_useOffset(False)
        # y 显示为科学计数或普通小数，保留合适位数
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.2f'))

        ax.xaxis.get_major_formatter().set_scientific(False)
        ax.xaxis.get_major_formatter().set_useOffset(False)
        plt.xticks(fontsize=16)
        plt.yticks(fontsize=16)
        plt.grid(True, linestyle='--', alpha=0.7, which='major')
        ax.minorticks_on()
        ax.grid(True, axis='x', linestyle=':', alpha=0.5, which='minor')
        # 设置x轴刻度步进
        if "Temperature" in xlabel or "group1" in prefix:
            # 图一：温度步进为5
            x_min, x_max = min(x), max(x)
            plt.xticks(np.arange(round(x_min), round(x_max) + 5, 5))
        elif "Current" in xlabel or "group2" in prefix:
            # 图二：电流步进为50
            x_min, x_max = min(x), max(x)
            plt.xticks(np.arange(round(x_min), round(x_max) + 50, 50))

        plt.tight_layout()
        plt.savefig(fig_path, dpi=300)
        plt.close()
        self.log(f"[Runner] 图像保存到 {fig_path}")

        # 可选择保存 csv（每行 x,y 以及 extra_cols）
        if save_csv:
            csv_path = os.path.join(out_dir, f"{prefix}_{timestamp}.csv")
            header = [xlabel, ylabel]
            extra_keys = []
            if extra_cols:
                extra_keys = list(extra_cols.keys())
                header += extra_keys
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(header)
                for i, xv in enumerate(x):
                    row = [xv, y[i]]
                    for k in extra_keys:
                        col = extra_cols.get(k, [])
                        row.append(col[i] if i < len(col) else "")
                    writer.writerow(row)
            self.log(f"[Runner] 数据 CSV 保存到 {csv_path}")

        return fig_path

    def run_group1(self, start_temp: float, end_temp: float, step: float, save_path: str = "./data", delay_s: float = 0.8, summary_filename: str = None, current_mA: float = None):
        """
        组1：在固定电流下，扫描温度；每步读取功率并汇总。
        """
        self._stop = False
        try:
            # 新增：检查并删除已存在的同名文件
            if summary_filename:
                if os.path.isdir(save_path) or save_path.endswith(os.sep):
                    out_dir = save_path
                else:
                    out_dir = os.path.dirname(save_path) or "."
                
                # 处理文件名，确保包含.csv扩展名
                if not summary_filename.lower().endswith('.csv'):
                    summary_filename += '.csv'
                
                # 构建完整文件路径
                file_path = os.path.join(out_dir, summary_filename)
                
                # 如果文件已存在，删除它
                if os.path.exists(file_path):
                    os.remove(file_path)
                    self.log(f"[Runner] 已删除同名文件: {file_path}")

            current_for_temp = 360.0
            if self.laser:
                # 优先使用传入的电流值，如果没有则读取当前电流
                if current_mA is not None:
                    try:
                        self.laser.set_current_mA(current_mA)
                        self.log(f"[Runner] 组1: 设置电流为 {current_mA:.2f} mA")
                        # 等待电流稳定
                        time.sleep(1.0)  # 简单延时等待电流稳定
                        current_for_temp = current_mA
                    except Exception as e:
                        self.log(f"[Runner] 组1: 设置电流失败，将使用当前电流值: {e}")
                        v = self.laser.get_current_mA()
                        if v is not None:
                            current_for_temp = v
                else:
                    v = self.laser.get_current_mA()
                    if v is not None:
                        current_for_temp = v
            
            temps = self._float_range(start_temp, end_temp, step)
            self.log(f"[Runner] 组1: 电流 {current_for_temp} mA 温度扫描 {start_temp}->{end_temp} step {step} 共 {len(temps)} 步, 等待 {delay_s}s")
            
            # 稳定参数设置
            stability_threshold = 0.1  # 稳定阈值，摄氏度
            max_wait_time = delay_s * 5  # 最大等待时间
            check_interval = 0.5  # 检查间隔
            
            for t in temps:
                if self._stop:
                    self.log("[Runner] 收到停止信号，结束组1")
                    break
                if self.laser:
                    try:
                        self.laser.set_temperature_C(t)
                        # 新增：等待温度稳定
                        self.log(f"[Runner] 设置温度为 {t}°C，等待稳定...")
                        wait_time = 0
                        stable = False
                        
                        # 先等待一段时间让温度开始变化
                        time.sleep(delay_s * 0.5)
                        
                        # 循环检查温度是否稳定
                        while wait_time < max_wait_time and not stable and not self._stop:
                            current_temp = self.laser.get_temperature_C()
                            if current_temp is not None:
                                temp_diff = abs(current_temp - t)
                                self.log(f"[Runner] 当前温度: {current_temp:.2f}°C, 目标: {t:.2f}°C, 差值: {temp_diff:.2f}°C")
                                
                                if temp_diff <= stability_threshold:
                                    stable = True
                                    self.log(f"[Runner] 温度已稳定在 {t}°C")
                                else:
                                    time.sleep(check_interval)
                                    wait_time += check_interval
                            else:
                                # 无法读取温度时，退化为简单延时
                                time.sleep(check_interval)
                                wait_time += check_interval
                        
                        if not stable and not self._stop:
                            self.log(f"[Runner] 温度在 {max_wait_time}s 内未完全稳定，继续测量")
                    except Exception as e:
                        self.log(f"[Runner] 设置温度失败: {e}")
                        # 设置失败时也等待一段时间
                        time.sleep(delay_s)
                else:
                    # 未连接激光控制器时，使用简单延时
                    time.sleep(delay_s)
                    
                try:
                    if not self.pm:
                        raise RuntimeError("未配置功率计 (PowerMeterController)")
                    power = self.pm.read_power()
                except Exception as e:
                    self.log(f"[Runner] 组1 读取功率失败 (temp {t}°C): {e}")
                    continue
                try:
                    self._append_summary(save_path, current_for_temp, t, power, test_group=1, summary_filename=summary_filename)
                    power_mw = float(power) * 1000
                    self.log(f"[Runner] 组1 {current_for_temp}mA, {t:.2f}°C -> Power {power_mw:.2f} mW")
                except Exception as e:
                    self.log(f"[Runner] 组1 写入汇总失败: {e}")
        except Exception as e:
            self.log(f"[Runner] 组1 出错: {e}\n{traceback.format_exc()}")
        self.log("[Runner] 组1 流程完成")

    def plot_group1_power_vs_temperature(self, out_dir, summary_filename=None):
        try:
            filename = summary_filename if summary_filename else "Test1_summary.csv"
            # 修复：自动处理没有.csv扩展名的情况
            if not filename.endswith('.csv'):
                filename += '.csv'
            file_path = os.path.join(out_dir, filename)
            if not os.path.exists(file_path):
                self.log(f"[Runner] {filename} 文件不存在: {file_path}")
                return None

            temps, powers = [], []
            with open(file_path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader)
                for row in reader:
                    try:
                        temp = float(row[1])
                        # 注意：如果CSV文件已经保存为mW单位，则不需要再次转换
                        # 如果是处理历史数据（之前以W保存的），则需要乘以1000
                        pwr = float(row[2])  # 假设CSV已保存为mW
                        temps.append(temp)
                        powers.append(pwr)
                    except Exception:
                        continue
            if temps:
                # 让温度按从高到低绘图
                uniq = {}
                for t, p in zip(temps, powers):
                    uniq[t] = p
                temps_sorted = sorted(uniq.keys(), reverse=True)
                powers_sorted = [uniq[t] for t in temps_sorted]
                return self._plot_xy_curve(
                    temps_sorted, powers_sorted,
                    xlabel="温度(°C)", ylabel="功率 (mW)",  # 修改Y轴标签
                    title=f"{self.laser.get_current_mA() if self.laser else 360:.2f} mA 下温度-功率关系",
                    out_dir=out_dir, prefix="温度功率关系图", invert_x=True, save_csv=False
                )
            else:
                self.log("[Runner] 组1 没有采集到有效功率数据，请检查 CSV 内容")
                return None
        except Exception as e:
            self.log(f"[Runner] 组1 绘制失败: {e}\n{traceback.format_exc()}")
            return None

    def run_group2(self, start_mA: float, step_mA: float, stop_mA: float, temp_C: float,
                   save_path: str = "./data", delay_s: float = 0.6, summary_filename: str = None):
        """
        组2：固定温度，扫描电流（从 start_mA 递减到 stop_mA），每步读取功率。
        """
        self._stop = False
        try:
            # 新增：检查并删除已存在的同名文件
            if summary_filename:
                if os.path.isdir(save_path) or save_path.endswith(os.sep):
                    out_dir = save_path
                else:
                    out_dir = os.path.dirname(save_path) or "."
                
                # 处理文件名，确保包含.csv扩展名
                if not summary_filename.lower().endswith('.csv'):
                    summary_filename += '.csv'
                
                # 构建完整文件路径
                file_path = os.path.join(out_dir, summary_filename)
                
                # 如果文件已存在，删除它
                if os.path.exists(file_path):
                    os.remove(file_path)
                    self.log(f"[Runner] 已删除同名文件: {file_path}")

            if self.laser:
                try:
                    self.laser.set_temperature_C(temp_C)
                    self.log(f"[Runner] 组2: 设置温度为 {temp_C:.2f} °C")
                    
                    # 新增：等待温度稳定
                    temp_stability_threshold = 0.1  # 温度稳定阈值，°C
                    temp_max_wait_time = 60  # 温度最大等待时间，秒
                    temp_check_interval = 2  # 温度检查间隔，秒
                    
                    self.log(f"[Runner] 组2: 等待温度稳定，阈值: {temp_stability_threshold}°C, 最大等待时间: {temp_max_wait_time}s")
                    temp_wait_time = 0
                    temp_stable = False
                    
                    # 循环检查温度是否稳定
                    while temp_wait_time < temp_max_wait_time and not temp_stable and not self._stop:
                        current_temp = self.laser.get_temperature_C()
                        if current_temp is not None:
                            temp_diff = abs(current_temp - temp_C)
                            self.log(f"[Runner] 当前温度: {current_temp:.2f}°C, 目标: {temp_C:.2f}°C, 差值: {temp_diff:.2f}°C")
                            
                            if temp_diff <= temp_stability_threshold:
                                temp_stable = True
                                self.log(f"[Runner] 温度已稳定在 {temp_C:.2f}°C")
                            else:
                                time.sleep(temp_check_interval)
                                temp_wait_time += temp_check_interval
                        else:
                            # 无法读取温度时，退化为简单延时
                            time.sleep(temp_check_interval)
                            temp_wait_time += temp_check_interval
                    
                    if not temp_stable and not self._stop:
                        self.log(f"[Runner] 温度在 {temp_max_wait_time}s 内未完全稳定，继续测量")
                except Exception as e:
                    self.log(f"[Runner] 组2: 设置温度失败 {e}")

            start_curr = float(start_mA)
            step_mag = abs(float(step_mA))
            stop_curr = float(stop_mA)
            if step_mag == 0:
                self.log("[Runner] group2_step_mA 不能为 0，已跳过组2")
                return
            currents = []
            c = start_curr
            while c >= stop_curr - 1e-9:
                currents.append(round(c, 6))
                c -= step_mag
            self.log(f"[Runner] 组2: 电流 {start_curr} -> {stop_curr} step {step_mag} 共 {len(currents)} 步, 等待 {delay_s}s")

            vals_curr = []
            vals_power = []
            stability_threshold = 1.0  # 电流稳定阈值，mA
            max_wait_time = delay_s * 3  # 最大等待时间
            check_interval = 0.3  # 检查间隔
            
            for cur in currents:
                if self._stop:
                    self.log("[Runner] 收到停止信号，提前结束组2")
                    break
                try:
                    if self.laser:
                        try:
                            self.laser.set_current_mA(cur)
                            # 新增：等待电流稳定
                            self.log(f"[Runner] 设置电流为 {cur}mA，等待稳定...")
                            wait_time = 0
                            stable = False
                            
                            # 循环检查电流是否稳定
                            while wait_time < max_wait_time and not stable and not self._stop:
                                current_current = self.laser.get_current_mA()
                                if current_current is not None:
                                    curr_diff = abs(current_current - cur)
                                    self.log(f"[Runner] 当前电流: {current_current:.2f}mA, 目标: {cur:.2f}mA, 差值: {curr_diff:.2f}mA")
                                    
                                    if curr_diff <= stability_threshold:
                                        stable = True
                                        self.log(f"[Runner] 电流已稳定在 {cur}mA")
                                    else:
                                        time.sleep(check_interval)
                                        wait_time += check_interval
                                else:
                                    # 无法读取电流时，退化为简单延时
                                    time.sleep(check_interval)
                                    wait_time += check_interval
                            
                            if not stable and not self._stop:
                                self.log(f"[Runner] 电流在 {max_wait_time}s 内未完全稳定，继续测量")
                        except Exception as e:
                            self.log(f"[Runner] 设置电流 {cur} mA 失败: {e}")
                            time.sleep(delay_s)  # 设置失败时也等待一段时间
                    else:
                        self.log(f"[Runner] 未配置 LaserController，跳过设置电流 {cur} mA")
                        time.sleep(delay_s)  # 未配置时使用简单延时

                    time.sleep(delay_s * 0.5)  # 额外小延时，确保系统稳定

                    if not self.pm:
                        raise RuntimeError("未配置功率计 (PowerMeterController)")
                    power = self.pm.read_power()

                    try:
                        self._append_summary(save_path, cur, temp_C, power, test_group=2, summary_filename=summary_filename)
                    except Exception as e:
                        self.log(f"[Runner] 组2 写入汇总失败: {e}")

                    vals_curr.append(cur)
                    vals_power.append(power)
                    power_mw = float(power) * 1000
                    self.log(f"[Runner] 组2 {int(cur)}mA @ {temp_C:.2f}°C -> Power {power_mw:.2f} mW")
                except Exception as e:
                    self.log(f"[Runner] 组2 电流 {cur} mA 处理失败: {e}")
                    continue

            if vals_curr:
                vals_power_mw = [float(p) * 1000 for p in vals_power]
                self._plot_xy_curve(
                    vals_curr, vals_power_mw,
                    xlabel="电流 (mA)", ylabel="功率 (mW)",
                    title=f"{temp_C:.2f}°C 下电流-功率关系",
                    out_dir=save_path, prefix="电流功率关系图", invert_x=False, save_csv=False,
                    extra_cols={"Temperature_C": [f"{temp_C:.2f}"] * len(vals_curr)}
                )
            else:
                self.log("[Runner] 组2 没有采集到任何功率数据，跳过作图")
        except Exception as e:
            self.log(f"[Runner] 组2 出错: {e}\n{traceback.format_exc()}")

# -------------------------
# GUI (大部分继承原结构，但把 OSA -> PowerMeter 转换)
# -------------------------
class CT_P_GUI:
    def __init__(self, parent=None):
        self.parent = parent
        
        # --- 核心修改：如果是集成模式，直接使用父控件作为 root ---
        if parent is None:
            self.root = tk.Tk()
            self.root.title("CT_P - 独立模式")
            # 假设 set_center() 只有在独立模式下需要
            if hasattr(self, 'set_center'):
                self.set_center(1510, 1090) 
            self.root.resizable(True, True)
            try:
                self.root.iconbitmap(r'PreciLasers.ico')
            except:
                pass
        else:
            self.root = parent # <--- 修改点：直接使用父 Frame

        # defaults
        self.params = {
            "usb_resource": "",            # 用于存放 VISA 资源字符串
            "current_mA": 360.0,
            "t_start": 36.0,
            "t_stop": 15.0,
            "t_step": 1.0,
            "laser_exe_path": r"C:\PTS\qijian\上位机软件\Preci_Semi\Preci-Seed.exe",
            "save_path": r"C:\PTS\qijian\CT_P",
            # group2 specific
            "group2_temp_C": 25.0,
            "group2_start_mA": 400.0,
            "group2_stop_mA": 0.5,
            "group2_step_mA": 5.0,
            # delays
            "group1_delay_s": 5,
            "group2_delay_s": 2,
            # filenames
            "group1_summary_filename": "Test1_summary",
            "group2_summary_filename": "Test2_summary"
        }

        self.param_labels = {
            "usb_resource": "USB 资源 (VISA)",
            "laser_exe_path": "软件路径",
            "save_path": "保存路径",
            "current_mA": "电流 (mA)",
            "t_start": "初始温度 (℃)",
            "t_stop": "终止温度 (℃)",
            "t_step": "温度步进 (℃)",
            "group2_temp_C": "组2 固定温度 (℃)",
            "group2_start_mA": "组2 初始电流 (mA)",
            "group2_stop_mA": "组2 终止电流 (mA)",
            "group2_step_mA": "组2 步进电流 (mA)",
            "group1_delay_s": "组1 稳定时间 (秒)",
            "group2_delay_s": "组2 稳定时间 (秒)",
            "group1_summary_filename": "组1文件名",
            "group2_summary_filename": "组2文件名"
        }

        self.create_widgets()
        self.laser: Optional[LaserController] = None
        self.pm: Optional[PowerMeterController] = None
        self.runner: Optional[TestRunner] = None
        self.runner_thread: Optional[threading.Thread] = None
        self.group1_running = False
        self.group2_running = False

    def set_center(self, width: int, height: int):
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        posx = (screenwidth - width) // 2
        posy = (screenheight - height) // 2
        self.root.geometry(f'{width}x{height}+{posx}+{posy}')

    def create_widgets(self):
        # 创建一个主框架来容纳参数设置和日志框，使用左右布局
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 左侧：参数设置
        param_frame = tk.LabelFrame(main_container, text="参数设置", padx=8, pady=8)
        param_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.entries: Dict[str, tk.Entry] = {}

        connect_frame = tk.LabelFrame(param_frame, text="连接与地址", padx=8, pady=8)
        connect_frame.pack(fill=tk.X, padx=6, pady=4)

        self._add_param_entry(connect_frame, "usb_resource", "USB资源:", self.params.get("usb_resource", ""), row=0)
        self._add_param_entry(connect_frame, "save_path", "保存路径:", self.params.get("save_path", "./data"), row=1)
        self._add_param_entry(connect_frame, "laser_exe_path", "软件路径:", self.params.get("laser_exe_path", ""), row=2)

        connect_buttons = tk.Frame(connect_frame)
        connect_buttons.grid(row=6, column=0, columnspan=3, pady=4)

        # 新增“地址”按钮
        def list_visa_resources():
            try:
                rm = pyvisa.ResourceManager()
                res = rm.list_resources()
                if not res:
                    self.log("[Diag] 未发现任何可用的 VISA 资源。")
                else:
                    self.log("[Diag] 可用 VISA 资源列表：")
                    for r in res:
                        self.log(f"    - {r}")
            except Exception as e:
                self.log(f"[Diag] VISA 资源枚举失败: {e}")

        self.btn_list_resources = tk.Button(connect_buttons, text="地址",
                                            command=list_visa_resources,
                                            bg="#1D74C0", fg="#FFFFFF", width=8)
        self.btn_list_resources.pack(side=tk.LEFT, padx=4)

        self.btn_connect = tk.Button(connect_buttons, text="连接", command=self.diag_connect_and_query, bg="#1D74C0", fg="#FFFFFF", width=8)
        self.btn_connect.pack(side=tk.LEFT, padx=4)

        self.btn_open_laser = tk.Button(connect_buttons, text="上位机", command=self.open_laser_software, bg="#1D74C0", fg="#FFFFFF", width=8)
        self.btn_open_laser.pack(side=tk.RIGHT, padx=4)

        # Group1
        group1_frame = tk.LabelFrame(param_frame, text="第一组测试", padx=6, pady=6)
        group1_frame.pack(fill="x", padx=6, pady=4)

        self._add_param_entry(group1_frame, "t_start", "初始温度:", self.params.get("t_start", 36.0), row=0)
        self._add_param_entry(group1_frame, "t_stop", "终止温度:", self.params.get("t_stop", 15.0), row=1)
        self._add_param_entry(group1_frame, "t_step", "步进温度:", self.params.get("t_step", 1.0), row=2)
        self._add_param_entry(group1_frame, "current_mA", "固定电流:", self.params.get("current_mA", 360.0), row=3)
        self._add_param_entry(group1_frame, "group1_delay_s", "稳定时间:", self.params.get("group1_delay_s", 5), row=4)
        self._add_param_entry(group1_frame, "group1_summary_filename", "保存文件名:", self.params.get("group1_summary_filename", "Test1_summary.csv"), row=5)

        group1_buttons = tk.Frame(group1_frame)
        group1_buttons.grid(row=6, column=0, columnspan=3, pady=4)
        self.btn_group1_start = tk.Button(group1_buttons, text="开始测试", command=self.start_group1, bg="#4CAF50", fg="#FFFFFF", width=12)
        self.btn_group1_start.pack(side=tk.LEFT, padx=4)
        self.btn_group1_stop = tk.Button(group1_buttons, text="停止测试", command=self.stop_group1, bg="#f44336", fg="#FFFFFF", width=12)
        self.btn_group1_stop.pack(side=tk.LEFT, padx=4)

        # Group2
        group2_frame = tk.LabelFrame(param_frame, text="第二组测试", padx=6, pady=6)
        group2_frame.pack(fill="x", padx=6, pady=4)

        self._add_param_entry(group2_frame, "group2_start_mA", "初始电流:", self.params.get("group2_start_mA", 400.0), row=0)
        self._add_param_entry(group2_frame, "group2_stop_mA", "终止电流:", self.params.get("group2_stop_mA", 0.5), row=1)
        self._add_param_entry(group2_frame, "group2_step_mA", "步进电流:", self.params.get("group2_step_mA", 5.0), row=2)
        self._add_param_entry(group2_frame, "group2_temp_C", "测试温度:", self.params.get("group2_temp_C", 25.0), row=3)
        self._add_param_entry(group2_frame, "group2_delay_s", "稳定时间:", self.params.get("group2_delay_s", 2), row=4)
        self._add_param_entry(group2_frame, "group2_summary_filename", "保存文件名:", self.params.get("group2_summary_filename", "Test2_summary.csv"), row=5)

        group2_buttons = tk.Frame(group2_frame)
        group2_buttons.grid(row=6, column=0, columnspan=3, pady=4)
        self.btn_group2_start = tk.Button(group2_buttons, text="开始测试", command=self.start_group2, bg="#4CAF50", fg="#FFFFFF", width=12)
        self.btn_group2_start.pack(side=tk.LEFT, padx=4)
        self.btn_group2_stop = tk.Button(group2_buttons, text="停止测试", command=self.stop_group2, bg="#f44336", fg="#FFFFFF", width=12)
        self.btn_group2_stop.pack(side=tk.LEFT, padx=4)

        # 右侧：日志框
        log_frame = tk.LabelFrame(main_container, text="运行日志", padx=6, pady=6)
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _add_param_entry(self, parent, key, label, default="", row=0, browse=None):
        tk.Label(parent, text=label, anchor="e", width=14).grid(row=row, column=0, sticky="e", padx=4, pady=4)
        ent = tk.Entry(parent, width=30)
        ent.insert(0, str(self.params.get(key, default)))
        ent.grid(row=row, column=1, padx=4, pady=4)
        self.entries[key] = ent
        if browse == "file":
            tk.Button(parent, text="浏览", command=lambda k=key: self.browse_file(k)).grid(row=row, column=2, padx=4, pady=4)
        if browse == "dir":
            tk.Button(parent, text="选择目录", command=lambda k=key: self.browse_savefile(k)).grid(row=row, column=2, padx=4, pady=4)
        return ent

    def log(self, msg: str):
        t = time.strftime("[%H:%M:%S]")
        try:
            self.log_box.insert(tk.END, f"{t} {msg}\n")
            self.log_box.see(tk.END)
            self.root.update_idletasks()
        except Exception:
            pass
        print(f"{t} {msg}")

    def open_laser_software(self):
        p = self.get_params()
        try:
            exe_path = p["laser_exe_path"]
            if not exe_path:
                messagebox.showerror("错误", "请先设置软件路径")
                return

            def _open_laser_thread():
                try:
                    self.laser = LaserController(exe_path=exe_path, window_title=r"Preci-Semi-Seed", log_func=self.log)
                    self.laser.connect()
                    self.log("[上位机] 已成功打开或连接到上位机软件")
                except Exception as e:
                    error_msg = f"[错误] 打开上位机软件失败: {e}"
                    self.log(error_msg)
                    self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                    self.laser = None

            thread = threading.Thread(target=_open_laser_thread, daemon=True)
            thread.start()
        except Exception as e:
            self.log(f"[错误] 准备打开上位机软件失败: {e}")
            messagebox.showerror("错误", f"准备打开上位机软件失败: {e}")

    def browse_file(self, param_key: str):
        filename = filedialog.askopenfilename(title="选择激光控制软件 (exe)", filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")])
        if filename:
            self.entries[param_key].delete(0, tk.END)
            self.entries[param_key].insert(0, filename)

    def browse_savefile(self, param_key: str):
        if messagebox.askyesno("选择", "选择保存目录？(否 = 选择具体文件名)"):
            dirname = filedialog.askdirectory(title="选择保存目录")
            if dirname:
                self.entries[param_key].delete(0, tk.END)
                self.entries[param_key].insert(0, dirname)
        else:
            filename = filedialog.asksaveasfilename(title="选择保存 文件", defaultextension=".csv", filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
            if filename:
                self.entries[param_key].delete(0, tk.END)
                self.entries[param_key].insert(0, filename)

    def get_params(self) -> Dict[str, Any]:
        p = {}
        for k in self.params.keys():
            try:
                if k in self.entries:
                    val = self.entries[k].get()
                    if k in ("laser_exe_path", "usb_resource", "save_path", "group1_summary_filename", "group2_summary_filename"):
                        p[k] = val
                    else:
                        p[k] = float(val)
                else:
                    p[k] = self.params[k]
            except Exception:
                p[k] = float(self.params[k]) if k not in ("laser_exe_path", "usb_resource", "save_path", "group1_summary_filename", "group2_summary_filename") else self.params[k]
        return p

    def show_image_popup(self, img_path, title="测试完成 - 截图预览"):
        win = tk.Toplevel(self.root)
        win.title(title)
        try:
            img = Image.open(img_path)
        except Exception as e:
            self.log(f"[错误] 无法打开图片: {e}")
            messagebox.showerror("错误", f"无法打开图片: {e}")
            return
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        max_w, max_h = int(sw * 0.8), int(sh * 0.8)
        disp_img = img
        if img.width > max_w or img.height > max_h:
            scale = min(max_w / img.width, max_h / img.height)
            new_size = (int(img.width * scale), int(img.height * scale))
            disp_img = img.resize(new_size, Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(disp_img)
        win.img = img
        win.disp_img = disp_img
        win.img_tk = img_tk

        btn_frame = tk.Frame(win)
        btn_frame.pack(side=tk.TOP, pady=8)

        def save_img():
            save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG 文件", "*.png"), ("所有文件", "*.*")], title="保存图片")
            if save_path:
                win.img.save(save_path)
                messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")

        save_btn = tk.Button(btn_frame, text="保存图片", command=save_img)
        save_btn.pack()

        lbl = tk.Label(win, image=win.img_tk)
        lbl.pack(padx=6, pady=6)

    # 诊断并连接功率计
    def diag_connect_and_query(self):
        usb_res = self.entries["usb_resource"].get().strip()
        if not usb_res:
            messagebox.showerror("错误", "请在诊断面板填写 USB 资源 (VISA地址)")
            return
        try:
            pm = PowerMeterController(resource=usb_res, log_func=self.log)
            pm.connect()
            idn = pm.query_idn()
            self.log(f"[Diag] 连接成功: IDN='{idn}'")
            self.pm = pm
        except Exception as e:
            self.log(f"[Diag] 连接/查询失败: {e}")
            messagebox.showerror("错误", f"诊断失败: {e}")

    def start_group1(self):
        p = self.get_params()
        self.btn_group1_start.config(state=tk.DISABLED)
        self.btn_group1_stop.config(state=tk.NORMAL)
        self.group1_running = True

        try:
            if not self.laser:
                self.laser = LaserController(exe_path=p["laser_exe_path"], window_title=r"Preci-Semi-Seed", log_func=self.log)
                try:
                    self.laser.connect()
                except Exception as e:
                    self.log(f"[错误] 激光控制软件连接失败: {e}")
                    if not messagebox.askyesno("警告", "激光控制软件连接失败，是否继续仅使用功率计?"):
                        self.btn_group1_start.config(state=tk.NORMAL)
                        self.btn_group1_stop.config(state=tk.DISABLED)
                        self.group1_running = False
                        return
                    else:
                        self.laser = None

            if not self.pm:
                usb_res = p["usb_resource"]
                if not usb_res:
                    raise RuntimeError("未填写 USB 资源地址")
                self.pm = PowerMeterController(resource=usb_res, log_func=self.log)
                self.pm.connect()

            if not self.runner:
                self.runner = TestRunner(self.laser, self.pm, log_func=self.log)
            else:
                self.runner._stop = False

            def target():
                try:
                    self.runner.run_group1(
                        start_temp=p["t_start"],
                        end_temp=p["t_stop"],
                        step=p["t_step"],
                        save_path=p["save_path"],
                        delay_s=p["group1_delay_s"],
                        summary_filename=p["group1_summary_filename"],
                        current_mA=p["current_mA"]  # 传递电流参数
                    )
                    img_path = self.runner.plot_group1_power_vs_temperature(
                        p["save_path"],
                        summary_filename=p["group1_summary_filename"]
                    )
                    if img_path and os.path.exists(img_path):
                        self.root.after(0, lambda: self.show_image_popup(img_path, "第一组测试完成 - 截图预览"))
                except Exception as e:
                    self.log(f"[线程异常] {e}\n{traceback.format_exc()}")
                finally:
                    try:
                        self.btn_group1_start.config(state=tk.NORMAL)
                        self.btn_group1_stop.config(state=tk.DISABLED)
                        self.group1_running = False
                    except Exception:
                        pass

            self.runner_thread = threading.Thread(target=target, daemon=True)
            self.runner_thread.start()
            self.log("[主] 第一组测试线程已启动")
        except Exception as e:
            self.log(f"[错误] 启动第一组测试失败: {e}")
            messagebox.showerror("错误", f"启动第一组测试失败: {e}")
            self.btn_group1_start.config(state=tk.NORMAL)
            self.btn_group1_stop.config(state=tk.DISABLED)
            self.group1_running = False

    def stop_group1(self):
        if self.runner and self.group1_running:
            try:
                self.runner.stop()
                self.log("[主] 第一组测试停止信号已发送")
            except Exception as e:
                self.log(f"[错误] 停止第一组测试失败: {e}")
        else:
            self.log("[主] 没有正在运行的第一组测试")

    def start_group2(self):
        p = self.get_params()
        self.btn_group2_start.config(state=tk.DISABLED)
        self.btn_group2_stop.config(state=tk.NORMAL)
        self.group2_running = True

        try:
            if not self.laser:
                self.laser = LaserController(exe_path=p["laser_exe_path"], window_title=r"Preci-Semi-Seed", log_func=self.log)
                try:
                    self.laser.connect()
                except Exception as e:
                    self.log(f"[错误] 激光控制软件连接失败: {e}")
                    if not messagebox.askyesno("警告", "激光控制软件连接失败，是否继续仅使用功率计?"):
                        self.btn_group2_start.config(state=tk.NORMAL)
                        self.btn_group2_stop.config(state=tk.DISABLED)
                        self.group2_running = False
                        return
                    else:
                        self.laser = None

            if not self.pm:
                usb_res = p["usb_resource"]
                if not usb_res:
                    raise RuntimeError("未填写 USB 资源地址")
                self.pm = PowerMeterController(resource=usb_res, log_func=self.log)
                self.pm.connect()

            if not self.runner:
                self.runner = TestRunner(self.laser, self.pm, log_func=self.log)
            else:
                self.runner._stop = False

            def target():
                try:
                    self.runner.run_group2(
                        start_mA=p["group2_start_mA"],
                        step_mA=p["group2_step_mA"],
                        stop_mA=p["group2_stop_mA"],
                        temp_C=p["group2_temp_C"],
                        save_path=p["save_path"],
                        delay_s=p["group2_delay_s"],
                        summary_filename=p["group2_summary_filename"]
                    )
                    # 找到最新保存的第二组图像并弹窗（同原逻辑）
                    import glob
                    pattern = os.path.join(p["save_path"], "电流功率关系图_*.png")
                    files = glob.glob(pattern)
                    if files:
                        files.sort(key=os.path.getmtime, reverse=True)
                        img_path = files[0]
                        if img_path and os.path.exists(img_path):
                            self.root.after(0, lambda: self.show_image_popup(img_path, "第二组测试完成 - 截图预览"))
                    else:
                        self.log("[Runner] 未找到第二组测试图像，请检查保存路径或命名。")
                except Exception as e:
                    self.log(f"[线程异常] {e}\n{traceback.format_exc()}")
                finally:
                    try:
                        self.btn_group2_start.config(state=tk.NORMAL)
                        self.btn_group2_stop.config(state=tk.DISABLED)
                        self.group2_running = False
                    except Exception:
                        pass

            self.runner_thread = threading.Thread(target=target, daemon=True)
            self.runner_thread.start()
            self.log("[主] 第二组测试线程已启动")
        except Exception as e:
            self.log(f"[错误] 启动第二组测试失败: {e}")
            messagebox.showerror("错误", f"启动第二组测试失败: {e}")
            self.btn_group2_start.config(state=tk.NORMAL)
            self.btn_group2_stop.config(state=tk.DISABLED)
            self.group2_running = False

    def stop_group2(self):
        if self.runner and self.group2_running:
            try:
                self.runner.stop()
                self.log("[主] 第二组测试停止信号已发送")
            except Exception as e:
                self.log(f"[错误] 停止第二组测试失败: {e}")
        else:
            self.log("[主] 没有正在运行的第二组测试")

    def single_scan(self):
        p = self.get_params()
        try:
            if not self.pm:
                usb_res = p["usb_resource"]
                if not usb_res:
                    raise RuntimeError("未填写 USB 资源地址")
                self.pm = PowerMeterController(resource=usb_res, log_func=self.log)
                self.pm.connect()
            # 读取单次功率
            power = self.pm.read_power()
            self.log(f"[单次] 读取功率: {power:.6e} W")

            save_base = p["save_path"]
            if os.path.isdir(save_base) or save_base.endswith(os.sep):
                fig_dir = save_base
            else:
                fig_dir = os.path.dirname(save_base) or "."
            ensure_dir(fig_dir)
            fig_path = os.path.join(fig_dir, f"single_power_{time.strftime('%Y%m%d_%H%M%S')}.png")

            plt.figure(figsize=(8, 4))
            # 单点绘制为折线（单点用点表示）
            plt.plot([0], [power], marker='o')
            plt.xlabel("Sample")
            plt.title("Single Power Reading")
            plt.ylabel("Power (W)")
            plt.tight_layout()
            plt.savefig(fig_path)
            plt.close()
            self.log(f"[单次] 图像保存到 {fig_path}")

            csv_fn = os.path.join(fig_dir, f"single_power_{time.strftime('%Y%m%d_%H%M%S')}.csv")
            with open(csv_fn, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Timestamp", "Power_W"])
                w.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), f"{power:.9e}"])
            self.log(f"[单次] CSV 保存到 {csv_fn}")

            # 弹出图片
            if os.path.exists(fig_path):
                self.show_image_popup(fig_path, "单次功率读取")
        except Exception as e:
            self.log(f"[错误] 单次读取失败: {e}\n{traceback.format_exc()}")
            messagebox.showerror("错误", f"单次读取失败: {e}")

    def run(self):
        if self.root.winfo_exists():
            self.root.mainloop()

if __name__ == "__main__":
    gui = CT_P_GUI()
    gui.run()
