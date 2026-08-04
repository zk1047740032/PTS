"""
2026.3.10：添加了读取上位机软件的类，之后要用再写
"""
from __future__ import annotations

import os
import time
import threading
import csv
import statistics
from typing import Optional, Dict, Any
import pyvisa
import tkinter as tk
from tkinter import filedialog

import sys, pathlib
# 确保能 import core（脚本独立运行时不以包形式组织）
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from core import (
    VisaInstrument,
    BaseTestGUI,
    clear_directory,
)
from core.config import CFG

# ===============  上位机控制（pywinauto）  ===============
try:
    from pywinauto.application import Application
    from pywinauto import timings
    PYW_AVAILABLE = True
except Exception:
    PYW_AVAILABLE = False

class PowerMeterController(VisaInstrument):
    """
    功率计控制器（继承 VisaInstrument，复用 ResourceManager/inst/close 等通用行为）。

    注意：本控制器的 connect() 与基类语义不同——原实现不设置读/写终止符且
    连接失败时抛异常（调用方用 try/except 处理），故保留子类 connect() 原样，
    不调用 super().connect()，以保证功率计通信行为与原版一致（零行为变更）。
    self.resource 为业务字段，对应基类 self.address。
    """
    def __init__(self, resource: str, log_func=print, timeout_ms: int = CFG.timing.visa_short_ms):
        """
        Initialize a PowerMeterController.

        Args:
            resource (str): VISA resource string identifying the power meter (e.g., 'USB::0x1234::0x5678::INSTR').
            log_func (callable): Logging function used to report status and errors. Defaults to print.
            timeout_ms (int): Timeout in milliseconds for VISA operations. Defaults to 5000.

        The controller creates a PyVISA ResourceManager and will open the
        specified instrument resource when :meth:`connect` is called.  The
        instance is stored in ``self.inst`` for subsequent commands.
        """
        super().__init__(log_func=log_func, address=resource, timeout_ms=timeout_ms)
        self.resource = resource  # 业务字段，对应基类 self.address

    def connect(self):
        # 优先 @py 纯 Python 后端，打不开资源时回退默认后端（NI-VISA ivi）
        tried = []
        for backend_tag in ("@py", ""):
            try:
                self.rm = (pyvisa.ResourceManager(backend_tag)
                           if backend_tag else pyvisa.ResourceManager())
                self.inst = self.rm.open_resource(self.resource)
                self.inst.timeout = int(self.timeout_ms)
                self.log(f"[PM] 已连接: {self.resource}")
                return
            except Exception as e:
                tried.append(f"{backend_tag or 'default'}: {e}")
                if self.rm:
                    try:
                        self.rm.close()
                    except Exception:
                        pass
                    self.rm = None
        err_msg = f"[PM] 连接失败，所有后端都打不开 {self.resource}: {' | '.join(tried)}"
        self.log(err_msg)
        raise RuntimeError(err_msg)

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
            token = resp.split()[0].replace(",", "")
            return float(token)
        except Exception as e:
            self.log(f"[PM] 命令 '{cmd}' 读取失败: {e}")
            return None

    def read_power(self) -> float:
        if self.inst is None:
            raise RuntimeError("功率计未连接")
        candidates = list(CFG.power.candidate_commands)
        last_errs = []
        for cmd in candidates:
            val = self._try_query_float(cmd)
            if val is not None:
                self.log(f"[PM] 命令 '{cmd}' 返回: {val} (W)")
                return float(val)
            else:
                last_errs.append(cmd)
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


class PowerCollector:
    def __init__(self, pm: Optional[PowerMeterController], log_func=print):
        self.pm = pm
        self.log = log_func

    def collect_and_save(self, save_path: str) -> bool:
        try:
            if not self.pm:
                raise RuntimeError("未配置功率计 (PowerMeterController)")

            power_w = self.pm.read_power()
            power_mw = float(power_w) * 1000
            timestamp = time.strftime('%Y-%m-%d %H:%M:%S')

            self.log(f"[Collector] 读取功率: {power_w:.6f} W = {power_mw:.3f} mW")

            if os.path.isdir(save_path) or save_path.endswith(os.sep):
                out_dir = save_path
            else:
                out_dir = os.path.dirname(save_path) or CFG.power.fallback_save_dir

            os.makedirs(out_dir, exist_ok=True)

            csv_filename = os.path.join(out_dir, f"power.csv")

            with open(csv_filename, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                # writer.writerow(["Power_mW"])
                writer.writerow([f"{power_mw:.2f}"])

            self.log(f"[Collector] 数据已保存到: {csv_filename}")
            return True
        except Exception as e:
            self.log(f"[Collector] 采集失败: {e}")
            return False


class PowerGUI(BaseTestGUI):
    """
    功率测量 GUI（继承 BaseTestGUI，复用窗口构造、线程安全日志(log)、
    auto_close、run 等通用能力，子类保留自身布局与业务按钮）。
    """
    def __init__(self, parent=None):
        super().__init__(parent, title="功率",
                         geometry="1275x510", icon="PreciLasers.ico")

        self.params = {
            "usb_resource": str(CFG.usb.power_meter),
            "save_path": str(CFG.dirs.power),
            "burnin_data": str(CFG.dirs.power / "burnin_data.csv"),
            "burnin_output": str(CFG.dirs.power)
        }

        self.param_labels = {
            "usb_resource": "USB 资源 (VISA)",
            "save_path": "保存路径",
            "burnin_data": "烤机数据",
            "burnin_output": "输出路径"
        }

        self.create_widgets()

        # 独立窗口模式：居中显示
        if parent is None:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = (sw - 1275) // 2
            y = (sh - 510) // 2
            self.root.geometry(f"1275x510+{x}+{y}")

        self.pm: Optional[PowerMeterController] = None
        self.collector: Optional[PowerCollector] = None

    def create_widgets(self):
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        left_container = tk.Frame(main_container)
        left_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        param_frame = tk.LabelFrame(left_container, text="功率读取", padx=4, pady=8)
        param_frame.pack(fill=tk.X, expand=False)

        self.entries: Dict[str, tk.Entry] = {}

        connect_frame = tk.Frame(param_frame, padx=4, pady=8)
        connect_frame.pack(fill=tk.X, padx=3, pady=4)

        self._add_param_entry(connect_frame, "usb_resource", "USB资源:", self.params.get("usb_resource", ""), row=0)
        self._add_param_entry(connect_frame, "save_path", "保存路径:", self.params.get("save_path", "./data"), row=1)

        # 创建按钮框架，放在参数设置边框内部
        buttons_frame = tk.Frame(param_frame)
        buttons_frame.pack(fill=tk.X, padx=3, pady=4)

        # 创建一个内部框架来容纳按钮并使其居中
        buttons_container = tk.Frame(buttons_frame)
        buttons_container.pack(side=tk.TOP, anchor=tk.CENTER)

        def list_visa_resources():
            # 优先 @py，回退默认后端。注意：@py 可能成功创建 RM 却
            # 找不到 USB 设备（比如仪器用的是 NI-VISA 驱动而非 WinUSB），
            # 此时继续尝试默认后端，避免空手而归。
            usb_res = []
            for backend_tag in ("@py", ""):
                label = backend_tag or "默认"
                try:
                    rm = (pyvisa.ResourceManager(backend_tag)
                          if backend_tag else pyvisa.ResourceManager())
                    res = rm.list_resources()
                    usb_res = [item for item in res if 'USB' in item]
                    rm.close()
                    if usb_res:
                        self.log(f"[Diag] 通过 {label} 后端找到 USB 资源")
                        break
                    else:
                        self.log(f"[Diag] {label} 后端未发现 USB 资源，继续尝试下一个...")
                except Exception as e:
                    self.log(f"[Diag] {label} 后端枚举失败: {e}")
                    continue
            if not usb_res:
                self.log("[Diag] 未发现任何可用 USB 资源。")
            else:
                self.log("[Diag] 可用USB资源：")
                for r in usb_res:
                    self.log(f"- {r}")

        self.btn_list_resources = tk.Button(buttons_container, text="地址",
                                            command=list_visa_resources,
                                            bg="#1D74C0", fg="#FFFFFF", width=8, cursor="hand2")
        self.btn_list_resources.pack(side=tk.LEFT, padx=8, expand=False)

        self.btn_connect = tk.Button(buttons_container, text="连接", command=self.connect_power_meter, bg="#1D74C0", fg="#FFFFFF", width=8, cursor="hand2")
        self.btn_connect.pack(side=tk.LEFT, padx=8, expand=False)

        self.btn_collect = tk.Button(buttons_container, text="开始", command=self.start_collect, bg="#4CAF50", fg="#FFFFFF", width=8, cursor="hand2")
        self.btn_collect.pack(side=tk.LEFT, padx=8, expand=False)

        # 烤机数据计算功能框
        burnin_frame = tk.LabelFrame(left_container, text="烤机数据计算", padx=4, pady=8)
        burnin_frame.pack(fill=tk.X, expand=False, pady=(10, 0))

        burnin_input_frame = tk.Frame(burnin_frame, padx=4, pady=8)
        burnin_input_frame.pack(fill=tk.X, padx=3, pady=4)

        self._add_param_entry(burnin_input_frame, "burnin_data", "烤机数据:", self.params.get("burnin_data", ""), row=0, browse="file", entry_width=25)
        self._add_param_entry(burnin_input_frame, "burnin_output", "输出路径:", self.params.get("burnin_output", ""), row=1, browse="dir", entry_width=25)

        # 烤机计算按钮框架
        burnin_buttons_frame = tk.Frame(burnin_frame)
        burnin_buttons_frame.pack(fill=tk.X, padx=3, pady=4)

        burnin_buttons_container = tk.Frame(burnin_buttons_frame)
        burnin_buttons_container.pack(side=tk.TOP, anchor=tk.CENTER)

        self.btn_burnin_start = tk.Button(burnin_buttons_container, text="开始计算", command=self.start_burnin_calculation, bg="#4CAF50", fg="#FFFFFF", width=8, cursor="hand2")
        self.btn_burnin_start.pack(side=tk.LEFT, padx=8, expand=False)

        log_frame = tk.LabelFrame(main_container, text="运行日志", padx=6, pady=6)
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _add_param_entry(self, parent, key, label, default="", row=0, browse=None, entry_width=30):
        tk.Label(parent, text=label, anchor="w", width=8).grid(row=row, column=0, sticky="w", padx=4, pady=4)
        ent = tk.Entry(parent, width=entry_width)
        ent.insert(0, str(self.params.get(key, default)))
        ent.grid(row=row, column=1, padx=4, pady=4)
        self.entries[key] = ent
        if browse == "file":
            tk.Button(parent, text="浏览", command=lambda k=key: self.browse_file(k),
                     bg="#F0F0F0", fg="#000000", width=4, height=1, relief="flat", cursor="hand2").grid(row=row, column=2, padx=0, pady=2)
        if browse == "dir":
            tk.Button(parent, text="浏览", command=lambda k=key: self.browse_dir(k),
                     bg="#F0F0F0", fg="#000000", width=4, height=0, relief="flat", cursor="hand2").grid(row=row, column=2, padx=0, pady=2)
        return ent

    # log() 复用基类 BaseTestGUI.log（线程安全，root.after 异步写入 log_box）

    def browse_file(self, param_key: str):
        filename = filedialog.askopenfilename(title="选择文件", filetypes=[("所有文件", "*.*")])
        if filename:
            self.entries[param_key].delete(0, tk.END)
            self.entries[param_key].insert(0, filename)

    def browse_dir(self, param_key: str):
        dirname = filedialog.askdirectory(title="选择保存目录")
        if dirname:
            self.entries[param_key].delete(0, tk.END)
            self.entries[param_key].insert(0, dirname)

    def get_params(self) -> Dict[str, Any]:
        p = {}
        for k in self.params.keys():
            try:
                if k in self.entries:
                    val = self.entries[k].get()
                    if k in ("usb_resource", "save_path", "burnin_data", "burnin_output"):
                        p[k] = val
                    else:
                        p[k] = float(val)
                else:
                    p[k] = self.params[k]
            except Exception:
                p[k] = self.params[k] if k in ("usb_resource", "save_path", "burnin_data", "burnin_output") else float(self.params[k])
        return p

    def connect_power_meter(self):
        usb_res = self.entries["usb_resource"].get().strip()
        if not usb_res:
            self.log("请填写 USB 资源 (VISA地址)")
            return
        try:
            self.pm = PowerMeterController(resource=usb_res, log_func=self.log)
            self.pm.connect()
            idn = self.pm.query_idn()
            self.log(f"[Diag] 连接成功: IDN='{idn}'")
        except Exception as e:
            self.log(f"[Diag] 连接失败: {e}")

    def start_collect(self):
        p = self.get_params()

        # 测试开始前清空输出文件夹
        save_dir = p.get("save_path", "")
        if save_dir:
            clear_directory(save_dir, log_func=self.log)

        if not self.pm:
            usb_res = p["usb_resource"]
            if not usb_res:
                self.log("未填写 USB 资源地址")
                self.auto_close(CFG.timing.auto_close_short_ms)
                return
            try:
                self.pm = PowerMeterController(resource=usb_res, log_func=self.log)
                self.pm.connect()
            except Exception as e:
                self.log(f"[错误] 连接功率计失败: {e}")
                self.log(f"连接功率计失败: {e}")
                self.auto_close(CFG.timing.auto_close_short_ms)
                return

        self.collector = PowerCollector(self.pm, log_func=self.log)

        def target():
            try:
                success = self.collector.collect_and_save(p["save_path"])
                if success:
                    self.log("功率数据采集完成!")
                else:
                    self.log("功率数据采集失败")
            except Exception as e:
                self.log(f"[线程异常] {e}")
                self.log(f"采集失败: {e}")
            finally:
                # 一键测试模式：自动关闭窗口，触发进程退出
                self.auto_close(CFG.timing.auto_close_short_ms)

        thread = threading.Thread(target=target, daemon=True)
        thread.start()
        self.log("[主] 采集线程已启动")

    def start_burnin_calculation(self):
        p = self.get_params()

        burnin_data = p.get("burnin_data", "").strip()
        burnin_output = p.get("burnin_output", "").strip()

        if not burnin_data:
            self.log("请选择烤机数据文件")
            return

        # 确定输出目录
        if burnin_output:
            output_dir = burnin_output
        else:
            # 如果没有指定输出路径，使用输入文件的目录
            output_dir = os.path.dirname(burnin_data) or CFG.power.fallback_save_dir

        self.log(f"[烤机计算] 输入文件: {burnin_data}")
        self.log(f"[烤机计算] 输出目录: {output_dir}")

        try:
            # 读取CSV文件中的Power列数据
            power_values = []
            with open(burnin_data, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)

                # 跳过前14行，第15行是表头
                for _ in range(CFG.power.burnin_header_lines):
                    next(reader, None)

                headers = next(reader, None)  # 读取第15行的表头

                # 查找Power列的索引
                power_col_index = None
                if headers:
                    for i, header in enumerate(headers):
                        if header.strip() == CFG.power.burnin_power_column:
                            power_col_index = i
                            break

                if power_col_index is None:
                    self.log("[烤机计算] 错误: 未找到'Power (W)'列")
                    return

                self.log(f"[烤机计算] 找到Power列，索引: {power_col_index}")

                # 读取Power列数据
                for row in reader:
                    if len(row) > power_col_index:
                        try:
                            value = float(row[power_col_index])
                            power_values.append(value)
                        except ValueError:
                            continue

            if not power_values:
                self.log("[烤机计算] 错误: 未读取到任何有效的Power数据")
                return

            self.log(f"[烤机计算] 读取到 {len(power_values)} 条Power数据")

            # 计算统计数据
            avg_value = statistics.mean(power_values)
            std_dev = statistics.stdev(power_values) if len(power_values) > 1 else 0
            max_value = max(power_values)
            min_value = min(power_values)

            self.log(f"[烤机计算] 平均值: {avg_value}")
            self.log(f"[烤机计算] 标准差: {std_dev}")
            self.log(f"[烤机计算] 最大值: {max_value}")
            self.log(f"[烤机计算] 最小值: {min_value}")

            # 计算RMS和P2P
            rms = (std_dev / avg_value) * 100 if avg_value != 0 else 0
            p2p = (max_value - min_value) / avg_value if avg_value != 0 else 0

            self.log(f"[烤机计算] RMS = (标准差/平均值)×100% = {rms:.2f}%")
            self.log(f"[烤机计算] P2P = (最大值-最小值)/平均值 = {p2p:.2f}")

            # 输出结果到CSV文件
            output_file = os.path.join(output_dir, "burnin_result.csv")
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["指标", "值"])
                writer.writerow(["RMS (%)", f"{rms:.2f}"])
                writer.writerow(["P2P", f"{p2p:.2f}"])
                writer.writerow(["平均值", f"{avg_value:.6f}"])
                writer.writerow(["标准差", f"{std_dev:.6f}"])
                writer.writerow(["最大值", f"{max_value:.6f}"])
                writer.writerow(["最小值", f"{min_value:.6f}"])

            self.log(f"[烤机计算] 结果已保存到: {output_file}")
            self.log("[烤机计算] 计算完成!")

        except FileNotFoundError:
            self.log(f"[烤机计算] 错误: 文件未找到: {burnin_data}")
        except Exception as e:
            self.log(f"[烤机计算] 处理失败: {e}")

    # run() 复用基类 BaseTestGUI.run（启动 mainloop）


if __name__ == "__main__":
    gui = PowerGUI()
    gui.run()
