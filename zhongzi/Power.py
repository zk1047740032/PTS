from __future__ import annotations

import os
import time
import threading
import csv
from typing import Optional, Dict, Any
import ctypes
import pyvisa
import tkinter as tk
from tkinter import filedialog

# 启用DPI感知，解决高DPI屏幕下界面模糊问题
if os.name == 'nt':
    try:
        # 设置进程DPI感知
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        # 获取系统DPI
        dpi = ctypes.windll.user32.GetDpiForSystem()
        # 设置缩放因子
        scaling_factor = dpi / 96.0
    except Exception:
        scaling_factor = 1.0
else:
    scaling_factor = 1.0

class PowerMeterController:
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
            token = resp.split()[0].replace(",", "")
            return float(token)
        except Exception as e:
            self.log(f"[PM] 命令 '{cmd}' 读取失败: {e}")
            return None

    def read_power(self) -> float:
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
                out_dir = os.path.dirname(save_path) or "."
            
            os.makedirs(out_dir, exist_ok=True)
            
            csv_filename = os.path.join(out_dir, f"power_{time.strftime('%Y%m%d_%H%M%S')}.csv")
            
            with open(csv_filename, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Power_W", "Power_mW"])
                writer.writerow([timestamp, f"{power_w:.9f}", f"{power_mw:.3f}"])
            
            self.log(f"[Collector] 数据已保存到: {csv_filename}")
            return True
        except Exception as e:
            self.log(f"[Collector] 采集失败: {e}")
            return False


class PowerGUI:
    def __init__(self, parent=None):
        self.parent = parent
        
        if parent is None:
            self.root = tk.Tk()
            self.root.title("功率计数据采集")
            self.root.geometry("1220x250")
            self.root.resizable(True, True)
            try:
                self.root.iconbitmap(r'PreciLasers.ico')
            except:
                pass
        else:
            self.root = parent

        self.params = {
            "usb_resource": "",
            "save_path": r"C:\PTS\zhongzi\Power"
        }

        self.param_labels = {
            "usb_resource": "USB 资源 (VISA)",
            "save_path": "保存路径"
        }

        self.create_widgets()
        self.pm: Optional[PowerMeterController] = None
        self.collector: Optional[PowerCollector] = None

    def create_widgets(self):
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        left_container = tk.Frame(main_container)
        left_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        param_frame = tk.LabelFrame(left_container, text="参数设置", padx=4, pady=8)
        param_frame.pack(fill=tk.X, expand=False)

        self.entries: Dict[str, tk.Entry] = {}

        connect_frame = tk.Frame(param_frame, padx=4, pady=8)
        connect_frame.pack(fill=tk.X, padx=3, pady=4)

        self._add_param_entry(connect_frame, "usb_resource", "USB资源:", self.params.get("usb_resource", ""), row=0)
        self._add_param_entry(connect_frame, "save_path", "保存路径:", self.params.get("save_path", "./data"), row=1)

        # 创建按钮框架，放在参数设置边框的正下方，并使其拉伸
        buttons_frame = tk.Frame(left_container)
        buttons_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        # 创建一个内部框架来容纳按钮并使其居中
        buttons_container = tk.Frame(buttons_frame)
        buttons_container.pack(side=tk.TOP, anchor=tk.CENTER)

        def list_visa_resources():
            try:
                rm = pyvisa.ResourceManager()
                res = rm.list_resources()
                usb_res = [item for item in res if 'USB' in item]
                if not usb_res:
                    self.log("[Diag] 未发现任何可用 USB 资源。")
                else:
                    self.log("[Diag] 可用USB资源：")
                    for r in usb_res:
                        self.log(f"- {r}")
            except Exception as e:
                self.log(f"[Diag] USB 资源枚举失败: {e}")

        self.btn_list_resources = tk.Button(buttons_container, text="地址",
                                            command=list_visa_resources,
                                            bg="#1D74C0", fg="#FFFFFF", width=8)
        self.btn_list_resources.pack(side=tk.LEFT, padx=8, expand=False)

        self.btn_connect = tk.Button(buttons_container, text="连接", command=self.connect_power_meter, bg="#1D74C0", fg="#FFFFFF", width=8)
        self.btn_connect.pack(side=tk.LEFT, padx=8, expand=False)

        self.btn_collect = tk.Button(buttons_container, text="开始", command=self.start_collect, bg="#4CAF50", fg="#FFFFFF", width=8)
        self.btn_collect.pack(side=tk.LEFT, padx=8, expand=False)

        log_frame = tk.LabelFrame(main_container, text="运行日志", padx=6, pady=6)
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _add_param_entry(self, parent, key, label, default="", row=0, browse=None):
        tk.Label(parent, text=label, anchor="w", width=8).grid(row=row, column=0, sticky="w", padx=4, pady=4)
        ent = tk.Entry(parent, width=30)
        ent.insert(0, str(self.params.get(key, default)))
        ent.grid(row=row, column=1, padx=4, pady=4)
        self.entries[key] = ent
        if browse == "file":
            tk.Button(parent, text="浏览", command=lambda k=key: self.browse_file(k)).grid(row=row, column=2, padx=4, pady=4)
        if browse == "dir":
            tk.Button(parent, text="选择目录", command=lambda k=key: self.browse_dir(k)).grid(row=row, column=2, padx=4, pady=4)
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
                    if k in ("usb_resource", "save_path"):
                        p[k] = val
                    else:
                        p[k] = float(val)
                else:
                    p[k] = self.params[k]
            except Exception:
                p[k] = self.params[k] if k in ("usb_resource", "save_path") else float(self.params[k])
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
        
        if not self.pm:
            usb_res = p["usb_resource"]
            if not usb_res:
                self.log("未填写 USB 资源地址")
                return
            try:
                self.pm = PowerMeterController(resource=usb_res, log_func=self.log)
                self.pm.connect()
            except Exception as e:
                self.log(f"[错误] 连接功率计失败: {e}")
                self.log(f"连接功率计失败: {e}")
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

        thread = threading.Thread(target=target, daemon=True)
        thread.start()
        self.log("[主] 采集线程已启动")

    def run(self):
        if self.root.winfo_exists():
            self.root.mainloop()


if __name__ == "__main__":
    gui = PowerGUI()
    gui.run()
