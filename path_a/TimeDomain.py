#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TimeDomain —— 时域测量程序（示波器 + 信号源）

两层架构（仪器 / GUI）：
    TimeDomain(VisaInstrument)  —— 示波器控制 + 信号源控制 + 时域测量
    TimeDomainGUI(BaseTestGUI)   —— 参数面板 + 按钮 + 日志
"""

import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from core import (
    VisaInstrument,
    BaseTestGUI,
    visa_address,
)
from core.config import CFG
from core.utils import append_row_csv, clear_directory

import time
import os
import csv
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk


# ============ TimeDomain 类 ============
class TimeDomain(VisaInstrument):
    """示波器 + 信号源双仪器控制器。

    示波器（scope）作为主仪器走 VisaInstrument 基类，
    信号源（gen）作为辅助仪器自行管理。
    """
    def __init__(self, params, log_func):
        super().__init__(log_func=log_func, timeout_ms=CFG.timing.visa_default_ms)
        self.params = params
        self.scope = None  # self.inst 的业务别名
        self.gen = None

    def connect_instruments(self):
        self.log("[示波器] 正在连接...")
        scope_address = visa_address(self.params['SCOPE_IP'], kind="inst0")
        if not super().connect(scope_address, idn=True):
            raise RuntimeError("示波器连接失败")
        self.scope = self.inst  # 业务别名，保持原方法中 self.scope 引用不变
        self.log(f"[示波器] 已连接：{self.idn}")

        self.log("[信号源] 正在连接...")
        gen_address = visa_address(self.params['GEN_IP'], kind="inst0")
        self.gen = self.rm.open_resource(gen_address)
        self.log(f"[信号源] 已连接：{self.gen.query('*IDN?').strip()}")

    def calculate_optimal_scale_factor(self, vpp):
        """根据峰峰值计算最佳放大倍数"""
        # 异常值处理：如果峰峰值过大或过小，使用默认值
        if vpp < CFG.time_domain.vpp_min_valid or vpp > CFG.time_domain.vpp_max_valid:
            self.log(f"[警告] 峰峰值 {vpp:.4f} V 异常，使用默认放大倍数16倍")
            return 16

        # 根据用户提供的经验值，建立峰峰值与最佳放大倍数的对应关系
        vpp_mv = vpp * 1000  # 转换为mV

        self.log(f"[调试] 峰峰值: {vpp_mv:.2f} mV")

        # 基于用户反馈优化的放大倍数选择逻辑
        # 不再局限于2的n次方，根据峰峰值动态调整
        scale_factors = dict(CFG.time_domain.scale_map)

        # 查找对应的放大倍数
        for (min_mv, max_mv), factor in scale_factors.items():
            if min_mv <= vpp_mv < max_mv:
                self.log(f"[调试] 峰峰值 {vpp_mv:.2f} mV 位于 {min_mv}-{max_mv} mV 范围，选择放大倍数: {factor} 倍")
                return factor

        # 默认返回16倍
        return 16

    def read_stable_vpp(self, channel, num_measurements=CFG.time_domain.vpp_stable_count, delay=CFG.time_domain.vpp_stable_delay_s):
        """读取稳定的峰峰值，去除异常值"""
        measurements = []

        for i in range(num_measurements):
            vpp = self.read_measurement(":MEAS:VPP?", channel)
            # 只保留合理范围内的测量值（0.01V到10V）
            if CFG.time_domain.vpp_min_valid <= vpp <= CFG.time_domain.vpp_max_valid:
                measurements.append(vpp)
            time.sleep(delay)

        if not measurements:
            # 如果所有测量值都异常，返回默认值
            return CFG.time_domain.vpp_default

        # 计算平均值
        return sum(measurements) / len(measurements)

    def configure_scope(self, freq):
        ch = self.params["SCOPE_CH"]
        self.log(f"[示波器] 配置 {ch} ...")
        self.scope.write(":STOP")
        self.scope.write(f":{ch}:COUP AC")
        self.log(f"[示波器] 设置时基模式为 YT")
        self.scope.write(":TIMebase:MODE MAIN")

        # 根据频率设置时基
        if freq == 100:
            timebase = 0.01  # 10ms
            self.log(f"[示波器] 设置时基为 10ms (100Hz)")
        elif freq == 300:
            timebase = 0.002  # 2ms
            self.log(f"[示波器] 设置时基为 2ms (300Hz)")
        else:
            timebase = 0.01  # 默认10ms
            self.log(f"[示波器] 设置默认时基为 10ms")
        # 发送时基设置指令
        self.scope.write(f":TIMebase:MAIN:SCALe {timebase}")

        # 第一步：设置一个适中的初始刻度，确保能测量到完整信号
        initial_scale = 0.5  # 先设置为0.5V/div
        self.scope.write(f":{ch}:SCAL {initial_scale}")
        self.scope.write(":RUN")
        time.sleep(CFG.time_domain.signal_stabilize_s)  # 增加等待时间，确保信号稳定
        self.scope.write(":MEAS:CLE")
        self.scope.write(f":MEAS:VAVG {ch}")
        self.scope.write(f":MEAS:VPP {ch}")

        # 测量稳定的峰峰值
        stable_vpp = self.read_stable_vpp(ch, num_measurements=CFG.time_domain.vpp_stable_count, delay=CFG.time_domain.vpp_stable_delay_s)
        self.log(f"[示波器] 稳定峰峰值测量: {stable_vpp:.4f} V (基于多次测量，去除异常值)")

        # 根据稳定的峰峰值计算最佳放大倍数
        optimal_scale_factor = self.calculate_optimal_scale_factor(stable_vpp)
        self.log(f"[示波器] 根据稳定峰峰值 {stable_vpp:.4f} V，自动选择放大倍数: {optimal_scale_factor} 倍")

        # 设置最终的垂直刻度
        final_scale = initial_scale / optimal_scale_factor
        self.scope.write(f":{ch}:SCAL {final_scale}")
        time.sleep(CFG.time_domain.signal_stabilize_s)  # 增加等待时间，确保信号稳定

        # 最终测量，确认效果
        final_vpp = self.read_stable_vpp(ch, num_measurements=CFG.time_domain.final_vpp_count, delay=CFG.time_domain.final_vpp_delay_s)
        self.log(f"[示波器] 最终峰峰值测量: {final_vpp:.4f} V")
        self.log(f"[示波器] 波形垂直刻度从 {initial_scale:.3f} V/div 调整到 {final_scale:.3f} V/div")

    def configure_gen(self):
        ch = self.params['GEN_CH']
        self.log(f"[信号源] 设置通道 {ch} TRI 波...")
        self.gen.write(f":SOUR{ch}:FUNC TRI")
        self.gen.write(f":SOUR{ch}:FREQ {self.params['GEN_FREQ']}")
        self.gen.write(f":SOUR{ch}:VOLT {self.params['GEN_VOLT']}")
        self.gen.write(f":SOUR{ch}:VOLT:OFFS {self.params['GEN_OFFSET']}")
        self.gen.write(f":OUTP{ch} ON")
        time.sleep(CFG.time_domain.gen_settle_s)

    def read_measurement(self, cmd, channel=None, retries=CFG.time_domain.meas_retries, delay=CFG.time_domain.meas_retry_delay_s):
        if channel is None:
            channel = self.params["SCOPE_CH"]
        for attempt in range(1, retries+1):
            try:
                val_str = self.scope.query(f"{cmd} {channel}").strip()
                if val_str and val_str not in CFG.time_domain.invalid_sentinels:
                    return float(val_str)
                self.log(f"[警告] 第 {attempt} 次读取失败，返回：{val_str}")
                time.sleep(delay)
            except Exception as e:
                self.log(f"[错误] 第 {attempt} 次读取异常：{e}")
                time.sleep(delay)
        raise RuntimeError(f"无法读取有效测量值：{cmd}")

    def save_data(self, data_dict, filename_base):
        os.makedirs(self.params["OUTPUT_DIR"], exist_ok=True)
        csv_path = os.path.join(self.params["OUTPUT_DIR"], f"{filename_base}.csv")
        dat_path = os.path.join(self.params["OUTPUT_DIR"], f"{filename_base}.dat")
        with open(csv_path, mode="w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(data_dict.keys())
            writer.writerow(data_dict.values())
        np.savetxt(dat_path, [list(data_dict.values())], header=" ".join(data_dict.keys()))
        self.log(f"[保存] 数据已保存到：{csv_path} 和 {dat_path}")

    def save_screenshot(self, filename="scope_screenshot.png"):
        self.scope.write(f":DISP:DATA? ON,OFF,PNG")
        raw_img = self.scope.read_raw()[11:]
        screenshot_path = os.path.join(self.params["OUTPUT_DIR"], filename)
        with open(screenshot_path, "wb") as f:
            f.write(raw_img)
        self.log(f"[保存] 截图已保存到 {screenshot_path}")
        return screenshot_path

    def close(self):
        # 示波器停止指令（保留原日志文案）
        if self.scope:
            try:
                self.scope.write(":STOP")
                self.log("[示波器] 已发送停止指令")
            except Exception as e:
                self.log(f"[错误] 发送停止指令失败：{e}")
        # 信号源关闭输出
        if self.gen:
            try:
                ch = self.params['GEN_CH']
                self.gen.write(f":OUTP{ch} OFF")
                self.log("[信号源] 已发送停止输出指令")
            except Exception as e:
                self.log(f"[错误] 发送停止输出指令失败：{e}")
            finally:
                self.gen.close()
        # 基类关闭：scope (self.inst) + rm
        super().close()

    def setCurAc(self):
        """设置电流耦合方式为直流，测试结束后调用"""
        ch = self.params["SCOPE_CH"]
        self.scope.write(f":{ch}:COUP DC")
        self.log(f"[示波器] 电流耦合方式改为直流")

# ============ GUI 类 ============
class TimeDomainGUI(BaseTestGUI):
    def __init__(self, parent=None):
        super().__init__(parent, title="时域 - 独立模式", geometry="1320x420",
                         resizable=(True, True))

        # 内部参数仍使用英文键
        self.params = {
            "SCOPE_IP": CFG.network.scope,
            "GEN_IP": CFG.network.sig_gen_timedomain,
            "OUTPUT_DIR": str(CFG.dirs.time_domain),
            "GEN_FREQ": CFG.time_domain.test_frequencies[0],  # 内部使用，不显示在UI
            "GEN_VOLT": CFG.time_domain.gen_volt,
            "GEN_OFFSET": CFG.time_domain.gen_offset,
            "GEN_CH": CFG.time_domain.gen_channel,
            "SCOPE_CH": CFG.time_domain.scope_channel,
        }

        # 中文显示对应表
        self.param_labels = {
            "SCOPE_IP": "示波器IP",
            "GEN_IP": "信号源IP",
            "OUTPUT_DIR": "输出目录",
            "GEN_CH": "信号通道",
            "SCOPE_CH": "示波器通道",
            "GEN_VOLT": "信号幅度(V)",
            "GEN_OFFSET": "信号偏置(V)",
        }

        self.create_widgets()

    def create_widgets(self):
        # 创建主容器，使用grid布局
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
        for key, val in self.params.items():
            # 跳过GEN_FREQ和SCOPE_CH，不显示在UI中
            if key in ["GEN_FREQ"]:
                continue
            tk.Label(param_frame, text=self.param_labels[key]).grid(row=row, column=0, sticky="e", padx=5, pady=2)
            entry = tk.Entry(param_frame, width=30)
            entry.insert(0, str(val))
            entry.grid(row=row, column=1, padx=5, pady=2)
            self.entries[key] = entry
            row += 1

        # --- 按钮区域 --- (左侧容器内，参数设置框下方，居中显示)
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=0, pady=8)

        # 创建一个内部框架来容纳按钮，实现居中
        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')

        # 添加按钮
        tk.Button(inner_btn_frame, text="保存参数", command=self.update_params, bg="#f4a236", fg="#FFFFFF", width=12, cursor="hand2").pack(side=tk.LEFT, padx=6)
        tk.Button(inner_btn_frame, text="开始测试", command=self.start_test, bg="#4CAF50", fg="#FFFFFF", width=12, cursor="hand2").pack(side=tk.LEFT, padx=6)

        # --- 日志显示区域 - 右侧 --- 占据整个右侧区域
        log_frame = tk.LabelFrame(main_frame, text="运行日志", padx=5, pady=5)
        log_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.log_box = tk.Text(log_frame)  # 基类 log() 需要的实例属性
        self.log_box.pack(fill=tk.BOTH, expand=True)

        # 设置grid权重，确保参数设置列固定，日志框列可以扩展
        main_frame.grid_columnconfigure(0, weight=0)  # 参数设置列固定大小
        main_frame.grid_columnconfigure(1, weight=1)  # 日志框列可以扩展
        main_frame.grid_rowconfigure(0, weight=1)     # 第一行可以扩展

    def update_params(self):
        for k, e in self.entries.items():
            v = e.get()
            try:
                if v.replace('.', '', 1).isdigit():
                    if '.' in v:
                        self.params[k] = float(v)
                    else:
                        self.params[k] = int(v)
                else:
                    self.params[k] = v
            except:
                self.params[k] = v
        # GEN_FREQ由程序内部控制，不通过UI更新
        self.log("[设置] 参数已更新")

    def start_test(self):
        """启动后台测试线程，避免阻塞 UI 日志更新。"""
        self.update_params()
        self.start_worker(self._run_test)

    def _run_test(self):
        """后台线程执行的测试逻辑。"""
        td = TimeDomain(self.params, self.log)
        try:
            # 清空输出文件夹
            out_dir = str(self.params["OUTPUT_DIR"])
            self.log(f"[测试] 正在清空输出文件夹: {out_dir}")
            clear_directory(out_dir, log_func=self.log)
            self.log("[测试] 输出文件夹清空完成")

            td.connect_instruments()
            # 保存原始频率参数
            original_freq = self.params["GEN_FREQ"]
            # 测试频率列表
            test_freqs = list(CFG.time_domain.test_frequencies)

            for freq in test_freqs:
                self.log(f"\n[测试] 开始 {freq}Hz 测试")
                # 设置当前测试频率
                self.params["GEN_FREQ"] = freq
                # 先配置信号源（设置频率）
                td.configure_gen()
                # 再配置示波器（根据频率设置时基）
                td.configure_scope(freq)
                # 暂停波形刷新，确保 Vpp 读数与截图一致
                td.scope.write(":STOP")
                time.sleep(1.5)
                # 读取测量结果
                vavg = td.read_measurement(":MEAS:VAVG?")
                vpp = td.read_measurement(":MEAS:VPP?")
                self.log(f"[结果] {freq}Hz - Vavg = {vavg:.4f} V")
                self.log(f"[结果] {freq}Hz - Vpp  = {vpp:.4f} V")
                # 截图前记录 Vpp 到 CSV
                vpp_csv_path = os.path.join(self.params["OUTPUT_DIR"], CFG.time_domain.vpp_csv_filename)
                append_row_csv(vpp_csv_path, [freq, vpp * 1000], header=["频率(Hz)", "Vpp(mV)"])
                self.log(f"[记录] Vpp 已写入 {vpp_csv_path}")
                # 保存数据，文件名包含频率信息
                #td.save_data({"Vavg(V)": vavg, "Vpp(V)": vpp}, filename_base=f"scope_measurement_{freq}Hz")
                # 保存截图，文件名包含频率信息
                screenshot = td.save_screenshot(filename=f"scope_screenshot_{freq}Hz.png")
                # 显示截图
                #self.show_image_popup(screenshot)
                #td.scope.write(":STOP")

            # 恢复原始频率参数
            self.params["GEN_FREQ"] = original_freq
            self.log("\n[测试] 所有频率测试完成")
            td.setCurAc()
        except Exception as e:
            self.log(f"[错误] 测试失败：{e}")
        finally:
            td.close()
            # 一键测试模式：自动关闭窗口，触发进程退出
            self.auto_close(CFG.timing.auto_close_short_ms)

    def show_image_popup(self, img_path):
        win = tk.Toplevel(self.root)
        # 从图片路径中提取频率信息
        if "100Hz" in img_path:
            freq = "100Hz"
        elif "300Hz" in img_path:
            freq = "300Hz"
        else:
            freq = "截图预览"
        win.title(f"{freq}")

        # 打开原始图片
        original_img = Image.open(img_path)

        # 创建画布
        canvas = tk.Canvas(win, bg="gray")
        canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 保存图片路径和原始图片的引用
        canvas.original_img = original_img

        # 绑定窗口大小变化事件
        def resize_image(event):
            # 获取画布的当前尺寸
            canvas_width = event.width
            canvas_height = event.height

            # 计算缩放比例，保持图片的宽高比
            img_ratio = original_img.width / original_img.height
            canvas_ratio = canvas_width / canvas_height

            if img_ratio > canvas_ratio:
                # 图片更宽，按宽度缩放
                new_width = canvas_width
                new_height = int(canvas_width / img_ratio)
            else:
                # 图片更高，按高度缩放
                new_height = canvas_height
                new_width = int(canvas_height * img_ratio)

            # 缩放图片
            resized_img = original_img.resize((new_width, new_height), Image.LANCZOS)
            img_tk = ImageTk.PhotoImage(resized_img)

            # 清除画布并显示新图片
            canvas.delete("all")
            canvas.create_image(canvas_width//2, canvas_height//2, anchor=tk.CENTER, image=img_tk)

            # 保存图片引用，防止被垃圾回收
            canvas.img_tk = img_tk

        # 初始显示图片
        canvas.bind("<Configure>", resize_image)

        # 保存图片功能
        def save_img():
            save_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG 文件", "*.png")],
                title="保存图片"
            )
            if save_path:
                original_img.save(save_path)
                messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")

        # 创建保存按钮
        button_frame = tk.Frame(win)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Button(button_frame, text="保存图片", command=save_img, cursor="hand2").pack(side=tk.BOTTOM, padx=5)


# ============ 程序入口 ============
if __name__ == "__main__":
    gui = TimeDomainGUI()
    gui.run()
