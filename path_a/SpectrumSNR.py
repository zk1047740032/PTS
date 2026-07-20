#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SpectrumSNR — 光谱信噪比测量程序

两层架构（仪器 / GUI）：
    SpectrumSNR(VisaInstrument)  —— 光谱仪控制 + SNR 测量 + 数据保存
    SpectrumSNRGUI(BaseTestGUI)   —— 参数面板 + 按钮 + 日志 + 线程调度
"""
from __future__ import annotations

import os
import time
import csv
import threading
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk, ImageDraw, ImageFont

import sys, pathlib
# 确保能 import core（脚本独立运行时不以包形式组织）
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from core import (
    VisaInstrument,
    BaseTestGUI,
    visa_address,
    clear_directory,
    append_row_csv,
)
from core.config import CFG

# ============ SpectrumSNR 类 ============
class SpectrumSNR(VisaInstrument):
    def __init__(self, params, log_func):
        timeout_s = float(params.get("VISA_TIMEOUT_S", CFG.spectrum_snr.visa_timeout_s))  # 默认 20 秒
        super().__init__(log_func=log_func, timeout_ms=int(timeout_s * 1000))
        self.params = params
        self.osa = None  # self.inst 的业务别名

    # --- 小工具：带重试的查询 ---
    def _query(self, cmd, retries=CFG.spectrum_snr.query_retries, delay=CFG.spectrum_snr.query_retry_delay_s):
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
    def connect(self):
        self.log("[光谱仪] 正在连接...")
        ip = self.params['OSA_IP']
        ok = super().connect(visa_address(ip, kind="tcpip_instr"), idn=True)
        if not ok:
            return False
        self.osa = self.inst  # 业务别名指向同一资源对象
        idn = self._idn
        self.log(f"[光谱仪] 已连接：{idn}")

        # ✅ 做一次零点校准
        self.log("[光谱仪] 开始零点校准...")
        self.inst.write(":SYSTem:ZERO:STARt")
        self._opc_wait("零点校准")
        return True

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
        self.osa.write(f":SENSe:SENSe {CFG.spectrum_snr.sensitivity}")
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
        mask = (wl < wl1 - CFG.spectrum_snr.main_peak_exclusion_nm) | (wl > wl1 + CFG.spectrum_snr.main_peak_exclusion_nm)
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
        append_row_csv(csv_path, [f"{snr:.2f}"])
        self.log(f"[保存] 结果已保存到：{csv_path}")
        return csv_path

    # 保存截图
    def save_screenshot(self, save_path=None, snr_value=None):
        try:
            os.makedirs(self.params["OUTPUT_DIR"], exist_ok=True)
            if save_path is None:
                save_path = os.path.join(self.params["OUTPUT_DIR"], "spectrum.bmp")

            # 设置长一点的超时，比如 180 秒
            self.osa.timeout = CFG.timing.visa_extra_long_ms

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
        """关闭仪器连接，清理业务别名"""
        super().close()
        self.osa = None


# ============ GUI 类 ============
class SpectrumSNRGUI(BaseTestGUI):
    def __init__(self, parent=None):
        super().__init__(parent, title="信噪比 - 独立模式",
                         geometry="1250x370")

        self.params = {
            "OSA_IP": CFG.network.osa,
            "OUTPUT_DIR": str(CFG.dirs.spectrum_snr),
            "CENTER": CFG.spectrum_snr.center_nm,
            "SPAN": CFG.spectrum_snr.span_nm,
            "REF_LEVEL": CFG.spectrum_snr.ref_level_dbm,
            "VISA_TIMEOUT_S": CFG.spectrum_snr.visa_timeout_s,
        }

        self.param_labels = {
            "OSA_IP": "光谱仪IP地址",
            "OUTPUT_DIR": "输出目录",
            "CENTER": "中心波长(nm)",
            "SPAN": "扫描范围(nm)",
            "REF_LEVEL": "参考电平(dBm)",
            "VISA_TIMEOUT_S": "VISA超时(s)",
        }

        self.create_widgets()

    # log() 复用基类 BaseTestGUI.log（线程安全，root.after 异步写入 log_box）

    def create_widgets(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        left_frame = tk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky="n", padx=(0, 5))

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

        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=0, pady=8)
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')

        # 保存按钮引用，以便禁用/启用
        self.btn_save = tk.Button(inner_btn_frame, text="保存参数", command=self.update_params, bg="#f4a236", fg="#FFFFFF", width=12, cursor="hand2")
        self.btn_save.pack(side=tk.LEFT, padx=6)

        self.btn_start = tk.Button(inner_btn_frame, text="开始测试", command=self.start_test, bg="#4CAF50", fg="#FFFFFF", width=12, cursor="hand2")
        self.btn_start.pack(side=tk.LEFT, padx=6)

        log_frame = tk.LabelFrame(main_frame, text="运行日志", padx=5, pady=5)
        log_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.log_box = tk.Text(log_frame)
        self.log_box.pack(fill=tk.BOTH, expand=True)

        main_frame.grid_columnconfigure(0, weight=0)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)

    def update_params(self):
        for k, e in self.entries.items():
            v = e.get().strip()
            try:
                num = float(v)
                if num.is_integer():
                    self.params[k] = int(num)
                else:
                    self.params[k] = num
            except Exception:
                self.params[k] = v
        self.log(f"[参数] 中心波长：{self.params['CENTER']}nm | 扫描范围：{self.params['SPAN']}nm")

    # --- 点击按钮只启动线程 ---
    def start_test(self):
        # 1. 禁用按钮，防止重复点击
        self.btn_start.config(state="disabled", text="测试中...")
        self.btn_save.config(state="disabled")

        # 2. 启动后台线程
        thread = threading.Thread(target=self.run_measurement_thread)
        thread.daemon = True  # 设置为守护线程，主程序关闭时它也会关闭
        thread.start()

    # --- 实际的耗时逻辑放在这里 ---
    def run_measurement_thread(self):
        osa = SpectrumSNR(self.params, self.log) # 注意：这里的 log 已经是线程安全的了
        try:
            if not osa.connect():
                self.log("[错误] 无法连接光谱仪")
                return
            clear_directory(self.params["OUTPUT_DIR"], log_func=self.log)
            self.log("[主机] 已清空文件夹")
            osa.configure_osa()
            snr, wl, power = osa.measure_snr()
            osa.save_data(snr)
            osa.save_curve(wl, power)
            screenshot = osa.save_screenshot(snr_value=snr)

            if screenshot:
                # 弹窗必须在主线程显示，否则报错
                self.root.after(0, lambda: self.show_image_popup(screenshot, snr))

        except Exception as e:
            self.log(f"[错误] 测试失败：{e}")
        finally:
            osa.close()
            # 恢复按钮状态（也要在主线程做）
            self.root.after(0, self.restore_ui_state)

    # --- 恢复界面状态 ---
    def restore_ui_state(self):
        self.btn_start.config(state="normal", text="开始测试")
        self.btn_save.config(state="normal")
        self.log("[系统] 测试流程结束，仪器已释放。")
        # 一键测试模式：自动关闭窗口，触发进程退出
        self.auto_close(CFG.timing.auto_close_short_ms)

    def show_image_popup(self, img_path, snr_value):
        # 注意：此函数由 root.after 调用，已经运行在主线程，可以安全操作 UI
        win = tk.Toplevel(self.root)
        win.title("测试完成 - 截图预览")

        try:
            img = Image.open(img_path)
            # 在图像上写 SNR 值
            draw = ImageDraw.Draw(img)
            try:
                font = ImageFont.truetype("arial.ttf", 32)
            except:
                font = ImageFont.load_default()
            text = f"SNR = {snr_value:.2f} dB"
            draw.text((80, 400), text, font=font, fill="white")

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
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".bmp",
                    filetypes=[("BMP 文件", "*.bmp"), ("PNG 文件", "*.png"), ("所有文件", "*.*")],
                    title="保存图片"
                )
                if save_path:
                    win.img.save(save_path)
                    messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")

            save_btn = tk.Button(btn_frame, text="保存图片", command=save_img, cursor="hand2")
            save_btn.pack()

            lbl = tk.Label(win, image=win.img_tk)
            lbl.pack(padx=6, pady=6)
        except Exception as e:
            self.log(f"[UI错误] 无法显示图片: {e}")

    # run() 复用基类 BaseTestGUI.run（启动 mainloop）


# ============ 程序入口 ============
if __name__ == "__main__":
    gui = SpectrumSNRGUI()
    gui.run()
