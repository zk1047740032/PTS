import pyvisa
import time
import os
import csv
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk, ImageDraw, ImageFont
import ctypes

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

# ============ SpectrumSNR 类 ============
class SpectrumSNR:
    def __init__(self, params, log_func):
        self.params = params
        self.log = log_func
        self.rm = None
        self.osa = None

    # --- 小工具：带重试的查询 ---
    def _query(self, cmd, retries=3, delay=0.4):
        last_err = None
        for i in range(1, retries + 1):
            try:
                resp = self.osa.query(cmd).strip()
                if resp != "" and resp.upper() != "NAN":
                    return resp
                self.log(f"[警告] 第 {i} 次查询返回异常：{cmd} -> {resp}")
            except Exception as e:
                last_err = e
                self.log(f"[警告] 第 {i} 次查询异常：{cmd} -> {e}")
            time.sleep(delay)
        if last_err:
            raise last_err
        raise RuntimeError(f"查询失败：{cmd}")

    # --- 小工具：等待操作完成 ---
    def _opc_wait(self, label="操作"):
        self._query("*OPC?")
        self.log(f"[光谱仪] {label} 已完成")

    # 连接仪器
    def connect_instrument(self):
        self.rm = pyvisa.ResourceManager()
        self.log("[光谱仪] 正在连接...")
        OSA_ADDR = f"TCPIP::{self.params['OSA_IP']}::INSTR"
        self.osa = self.rm.open_resource(OSA_ADDR)
        timeout_s = float(self.params.get("VISA_TIMEOUT_S", 20))  # 默认 20 秒
        self.osa.timeout = int(timeout_s * 1000)  # 转换为毫秒
        self.osa.write_termination = "\n"
        self.osa.read_termination = "\n"
        idn = self._query("*IDN?")
        self.log(f"[光谱仪] 已连接：{idn}")

        # ✅ 做一次零点校准
        self.log("[光谱仪] 开始零点校准...")
        self.osa.write(":SYSTem:ZERO:STARt")
        self._opc_wait("零点校准")

    # 配置光谱仪
    def configure_osa(self):
        self.log("[光谱仪] 配置扫描参数...")

        # 停止当前扫频
        self.osa.write(":ABORt")
        # 设置单位和中心/跨度
        self.osa.write(":UNIT:X WAVelength")
        self.osa.write(f":SENSe:WAVelength:CENTer {self.params['CENTER']}NM")
        self.osa.write(f":SENSe:WAVelength:SPAN {self.params['SPAN']}NM")
        # 设置灵敏度为HIGH1（根据手册中的正确格式，注意大小写）
        self.osa.write(":SENSe:SENSe HIGH1")
        # 添加读取设置确认
        try:
            sense_value = self.osa.query(":SENSe:SENSe?")
            self.log(f"[光谱仪] 灵敏度已设置为HIGH1，确认值: {sense_value.strip()} (HIGH1对应值应为3)")
        except Exception as e:
            self.log(f"[光谱仪] 灵敏度已设置为HIGH1，但确认查询失败: {e}")

        # 设置参考电平（REF_LEVEL，单位 dBm）
        try:
            ref_level = float(self.params.get("REF_LEVEL", -4.0))
            # 先发设置命令（根据手册，使用完整的SCPI命令格式，包含Y1轨迹）
            self.osa.write(f":DISPlay:WINDow:TRACe:Y1:SCALe:RLEVel {ref_level}DBM")

            # 尝试读回确认，尝试几个常见的查询命令
            query_cmds = [
                ":DISPlay:WINDow:TRACe:Y1:SCALe:RLEVel?",
                ":DISPlay:WINDow:TRACe:Y:SCALe:RLEVel?",
                ":DISPlay:RLEVel?",
                ":DISP:RLEVel?",
                ":SENSe:POWer:REF?",
            ]
            readback = None
            for qc in query_cmds:
                try:
                    resp = self.osa.query(qc).strip()
                    if resp != "":
                        readback = (qc, resp)
                        break
                except Exception:
                    continue

            if readback:
                self.log(f"[光谱仪] 参考电平设置为 {ref_level} dBm，读回: {readback[1]} (via {readback[0]})")
            else:
                self.log(f"[光谱仪] 参考电平设置命令已发送: {ref_level} dBm（未能读回确认，可能该指令在此型号上不可查询）")
        except Exception as e:
            self.log(f"[光谱仪] 设置参考电平失败: {e}")

        # 读回确认
        cen_m = float(self._query(":SENSe:WAVelength:CENTer?"))
        span_m = float(self._query(":SENSe:WAVelength:SPAN?"))
        self.log(f"[光谱仪] 已设置 CENTER={cen_m*1e9:.3f} nm, SPAN={span_m*1e9:.3f} nm")

    # 测量光谱信噪比（曲线分析）
    def measure_snr(self):
        self.log("[光谱仪] 读取光谱曲线，计算主峰和次峰...")

        # 1. 扫描一次
        self.osa.write(":INITiate:SMODe SINGle")
        self.osa.write(":INITiate")
        self._opc_wait("光谱扫描")

        # 2. 获取波长和功率数据
        wl = np.array(self.osa.query_ascii_values(":TRACe:X? TRA")) * 1e9  # m -> nm
        power = np.array(self.osa.query_ascii_values(":TRACe:Y? TRA"))     # dBm
        self.log(f"[光谱仪] 获取到 {len(wl)} 个点")

        if len(wl) == 0 or len(power) == 0:
            raise RuntimeError("未获取到曲线数据")

        # 3. 找主峰
        idx_max = np.argmax(power)
        p1 = power[idx_max]
        wl1 = wl[idx_max]
        self.log(f"[主峰] {wl1:.3f} nm, {p1:.2f} dBm")

        # 4. 在 ±3 nm 以外找次峰
        mask = (wl < wl1 - 3) | (wl > wl1 + 3)
        if not np.any(mask):
            raise RuntimeError("没有找到 ±3 nm 以外的数据点")

        masked_power = power[mask]
        masked_wl = wl[mask]
        idx_second = np.argmax(masked_power)
        p2 = masked_power[idx_second]
        wl2 = masked_wl[idx_second]
        self.log(f"[次峰] {wl2:.3f} nm, {p2:.2f} dBm")

        # 5. 计算 SNR
        snr = p1 - p2
        self.log(f"[结果] 光谱信噪比 = {snr:.2f} dB")

        return snr, wl, power

    # 保存数据
    def save_data(self, snr, filename_base="spectrum_snr"):
        os.makedirs(self.params["OUTPUT_DIR"], exist_ok=True)
        csv_path = os.path.join(self.params["OUTPUT_DIR"], f"{filename_base}.csv")
        with open(csv_path, mode="w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["SNR(dB)"])
            writer.writerow([snr])
        self.log(f"[保存] 结果已保存到：{csv_path}")
        return csv_path

    # 保存截图
    def save_screenshot(self, save_path=None, snr_value=None):
        try:
            os.makedirs(self.params["OUTPUT_DIR"], exist_ok=True)
            if save_path is None:
                save_path = os.path.join(self.params["OUTPUT_DIR"], "spectrum.bmp")

            # 设置长一点的超时，比如 180 秒
            self.osa.timeout = 180000  

            # 1. 在仪器里保存截图到内部存储（BMP 格式）
            self.osa.write(':MMEMory:STORe:GRAPhics COLor,BMP,"spectrum",INT')
            self._opc_wait("保存截图到内部存储")

            # 2. 从内部存储读取文件
            raw = self.osa.query_binary_values(':MMEMory:DATA? "spectrum.bmp",INT',
                                               datatype='B', container=bytes)

            # 3. 写入到 PC
            temp_path = os.path.join(self.params["OUTPUT_DIR"], "temp_spectrum.bmp")
            with open(temp_path, "wb") as f:
                f.write(raw)

            # 4. 如果提供了SNR值，添加到图片上
            if snr_value is not None:
                from PIL import Image, ImageDraw, ImageFont
                img = Image.open(temp_path)
                draw = ImageDraw.Draw(img)
                try:
                    font = ImageFont.truetype("arial.ttf", 32)  # 如果系统有 Arial
                except:
                    font = ImageFont.load_default()
                text = f"SNR = {snr_value:.2f} dB"
                draw.text((80, 400), text, font=font, fill="white")  # 左上角写字
                img.save(save_path)
            else:
                # 没有SNR值，直接重命名临时文件
                os.replace(temp_path, save_path)

            self.log(f"[光谱仪] 截图保存成功: {save_path}")
            return save_path   # ✅ 返回文件路径

        except Exception as e:
            self.log(f"[错误] 截图保存失败: {e}")
            return None
        
    # 保存完整曲线
    def save_curve(self, wl, power, filename_base="spectrum_curve"):
        os.makedirs(self.params["OUTPUT_DIR"], exist_ok=True)
        csv_path = os.path.join(self.params["OUTPUT_DIR"], f"{filename_base}.csv")
        with open(csv_path, mode="w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Wavelength (nm)", "Power (dBm)"])
            for x, y in zip(wl, power):
                writer.writerow([x, y])
        self.log(f"[保存] 光谱曲线已保存到：{csv_path}")
        return csv_path



    def close(self):
        try:
            if self.osa:
                self.osa.close()
        finally:
            if self.rm:
                self.rm.close()

# ============ GUI 类 ============
class SpectrumSNRGUI:
    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
        else:
            self.root = tk.Toplevel(parent)
        self.root.title("信噪比")
        self.root.resizable(True, True)

        def SetCenter(window, width, height):
            sw, sh = window.winfo_screenwidth(), window.winfo_screenheight()
            window.geometry(f"{width}x{height}+{(sw-width)//2}+{(sh-height)//2}")

        SetCenter(self.root, 1280, 370)

        # 默认参数
        self.params = {
            "OSA_IP": "192.168.7.14",
            "OUTPUT_DIR": r"C:\PTS\zhongzi\SpectrumSNR",
            "CENTER": 1064,
            "SPAN": 150,      # 默认 150 nm
            "REF_LEVEL": -4.0, # 参考电平 (dBm)
            "VISA_TIMEOUT_S": 120,  # 20s
        }

        # 参数标签（去掉 CENTER 和 SPAN 的输入框）
        self.param_labels = {
            "OSA_IP": "光谱仪IP地址",
            "OUTPUT_DIR": "输出目录",
            "CENTER": "中心波长(nm)",
            "SPAN": "扫描范围(nm)",
            "REF_LEVEL": "参考电平(dBm)",
            "VISA_TIMEOUT_S": "VISA超时(s)",
        }

        self.create_widgets()

    def log(self, msg):
        t = time.strftime("[%H:%M:%S]")
        self.log_box.insert(tk.END, f"{t} {msg}\n")
        self.log_box.see(tk.END)
        self.root.update()

    def create_widgets(self):
        # 创建主框架，使用grid布局
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 左侧容器 - 用于容纳参数设置框和按钮，形成一个整体
        left_frame = tk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky="n", padx=(0, 5))
        
        # --- 参数设置 --- (左侧容器内) - 固定大小，不随窗口拉伸
        param_frame = tk.LabelFrame(left_frame, text="参数设置", padx=10, pady=10)
        param_frame.pack(fill=tk.X, padx=0, pady=0)

        self.entries = {}
        row = 0
        for k, v in self.params.items():
            tk.Label(param_frame, text=self.param_labels[k]).grid(row=row, column=0, sticky="e")
            entry = tk.Entry(param_frame, width=25)
            entry.insert(0, str(v))
            entry.grid(row=row, column=1, padx=5, pady=2)
            self.entries[k] = entry
            row += 1

        # --- 按钮区域 --- (左侧容器内，参数设置框下方，居中显示)
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=0, pady=8)
        
        # 创建一个内部框架来容纳按钮，实现居中
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')
        
        # 添加按钮
        tk.Button(inner_btn_frame, text="保存参数", command=self.update_params, bg="#f4a236", fg="#FFFFFF", width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(inner_btn_frame, text="开始测试", command=self.start_test, bg="#4CAF50", fg="#FFFFFF", width=12).pack(side=tk.LEFT, padx=6)

        # --- 日志窗口 --- (右侧) - 占据整个右侧区域
        log_frame = tk.LabelFrame(main_frame, text="运行日志", padx=5, pady=5)
        log_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)
        
        # 设置grid权重，确保参数设置列固定，日志框列可以扩展
        main_frame.grid_columnconfigure(0, weight=0)  # 参数设置列固定大小
        main_frame.grid_columnconfigure(1, weight=1)  # 日志框列可以扩展
        main_frame.grid_rowconfigure(0, weight=1)     # 第一行可以扩展

    def update_params(self):
        # 普通输入框参数
        for k, e in self.entries.items():
            v = e.get()
            v = v.strip()
            try:
                # 尝试将可解析为数字的字符串转换为 float/int（支持负数）
                num = float(v)
                if num.is_integer():
                    self.params[k] = int(num)
                else:
                    self.params[k] = num
            except Exception:
                # 不能解析为数字则保持字符串原样
                self.params[k] = v

        self.log(f"[参数] 中心波长：{self.params['CENTER']}nm | 扫描范围：{self.params['SPAN']}nm")

    def start_test(self):
        osa = SpectrumSNR(self.params, self.log)
        try:
            osa.connect_instrument()
            osa.configure_osa()
            snr, wl, power = osa.measure_snr()
            osa.save_data(snr)
            osa.save_curve(wl, power) 
            screenshot = osa.save_screenshot(snr_value=snr)
            if screenshot:   # 成功才展示
                self.show_image_popup(screenshot, snr)
        except Exception as e:
            self.log(f"[错误] 测试失败：{e}")
        finally:
            osa.close()

    def show_image_popup(self, img_path, snr_value):
        win = tk.Toplevel(self.root)
        win.title("测试完成 - 截图预览")

        # 读取原始图片
        img = Image.open(img_path)

        # ✅ 在图像上写 SNR 值
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("arial.ttf", 32)  # 如果系统有 Arial
        except:
            font = ImageFont.load_default()
        text = f"SNR = {snr_value:.2f} dB"
        draw.text((80, 400), text, font=font, fill="white")  # 左上角写字

        # 获取屏幕尺寸
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        max_w, max_h = int(sw * 0.8), int(sh * 0.8)

        disp_img = img
        if img.width > max_w or img.height > max_h:
            scale = min(max_w / img.width, max_h / img.height)
            new_size = (int(img.width * scale), int(img.height * scale))
            disp_img = img.resize(new_size, Image.LANCZOS)

        img_tk = ImageTk.PhotoImage(disp_img)

        # 挂载引用，避免被回收
        win.img = img
        win.disp_img = disp_img
        win.img_tk = img_tk

        # 顶部按钮区
        btn_frame = tk.Frame(win)
        btn_frame.pack(side=tk.TOP, pady=8)

        def save_img():
            save_path = filedialog.asksaveasfilename(
                defaultextension=".bmp",
                filetypes=[("BMP 文件", "*.bmp"), ("PNG 文件", "*.png"), ("所有文件", "*.*")],
                title="保存图片"
            )
            if save_path:
                # ✅ 保存的就是带文字的图
                win.img.save(save_path)
                messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")

        save_btn = tk.Button(btn_frame, text="保存图片", command=save_img)
        save_btn.pack()

        # 显示图片
        lbl = tk.Label(win, image=win.img_tk)
        lbl.pack(padx=6, pady=6)

    def run(self):
        # 保持原有的run方法
        if self.root.winfo_exists():
            self.root.mainloop()


# ============ 程序入口 ============
if __name__ == "__main__":
    gui = SpectrumSNRGUI()
    gui.run()
