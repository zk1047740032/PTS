import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows DPI 自适应

import os
import re
import csv
import time
import math
import queue
import functools
import threading
import pyvisa
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import serial
import matplotlib
matplotlib.use('Agg')  # 后端绘图，不阻塞 GUI
import matplotlib.pyplot as plt

from PIL import Image, ImageTk

# ===============  DFB 种子激光器 RS-485 串口控制  ===============
class DFBLaserController:
    """通过 RS-485 串口直连 DFB 种子激光器，替代 pywinauto GUI 操控

    协议帧格式: 0x50 | 0x00 | ADDR | 命令 | 数据长度 | [数据] | 和校验 | 异或校验 | 0x0D | 0x0A
    数值转换: 温度保留三位小数 ×1000，电流单位直接为 mA（无需乘除）
    """

    CMD_SYS_INFO  = 0xAA   # 系统信息查询 → 响应 0xB0
    CMD_REALTIME  = 0xA9   # 实时查询     → 响应 0xB7
    # ---- 以下为设置指令（来自 docs/DFB协议.md 第 119-204 行）----
    CMD_SET_CURRENT1 = 0xA3   # 设置电流1 ( 2B 两位小数 + 1B 保存标志)
    CMD_SET_CURRENT2 = 0xA4   # 设置电流2
    CMD_SET_TEMP_GRATING = 0xA5  # 设置光栅温度 (2B 三位小数 + 1B 保存标志)
                                # 温度*1000 = (波长*10000 - Wave_T_offset) * 10000 / Wave_T_index
    CMD_SET_POWER_MODE = 0x5A   # 功率模式设置 (2B Power_Set + 1B 开关)
    CMD_SET_CURRENT_SW = 0xA8   # 电流开关 (1B 开关 + 1B 保存标志)

    def __init__(self, port, addr=100, log_func=print):
        self.port = port
        self.addr = addr
        self.log = log_func
        self.serial = None
        self._lock = threading.Lock()
        # 系统信息缓存（用于波长->温度换算），_read_system_info() 填充
        self._wave_t_index = None   # Wave_T_index（2B，符号未定）
        self._wave_t_offset = None  # Wave_T_offset（4B，符号未定）

    def open(self, timeout_s=1.5):
        self.serial = serial.Serial(
            port=self.port,
            baudrate=115200,
            timeout=timeout_s,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
        )
        self.log(f"[DFB激光器] 串口 {self.port} 已连接")

    def close(self):
        if self.serial and self.serial.is_open:
            self.serial.close()

    # ---------- 协议底层 ----------
    @staticmethod
    def _checksum(data: bytes):
        """返回 (和校验, 异或校验)"""
        return sum(data) & 0xFF, functools.reduce(lambda a, b: a ^ b, data, 0)

    def _send_frame(self, cmd_byte: int, payload: bytes = b'', timeout: float | None = None) -> bytes | None:
        """timeout=None 使用串口默认超时；传入 float 则临时覆盖（在锁内切换，线程安全）"""
        frame = bytes([0x50, 0x00, self.addr, cmd_byte, len(payload)])
        frame += payload
        s, x = self._checksum(frame[1:])
        frame += bytes([s, x, 0x0D, 0x0A])

        with self._lock:
            if timeout is not None:
                orig_timeout = self.serial.timeout
                self.serial.timeout = timeout
            try:
                self.serial.reset_input_buffer()
                self.serial.write(frame)

                header = self.serial.read(5)
                if len(header) < 5 or header[0] != 0x50:
                    return None

                data_len = header[4]
                rest_len = data_len + 4
                rest_data = self.serial.read(rest_len)

                if len(rest_data) < rest_len:
                    return None

                full_frame = header + rest_data
                recv_s, recv_x = self._checksum(full_frame[1:-4])
                if recv_s != full_frame[-4] or recv_x != full_frame[-3]:
                    return None

                return full_frame
            finally:
                if timeout is not None:
                    self.serial.timeout = orig_timeout

    def _query_realtime(self, timeout: float | None = None) -> dict[str, float] | None:
        """发送实时查询命令，返回解析后的字典，或 None"""
        raw = self._send_frame(self.CMD_REALTIME, timeout=timeout)
        if raw is None or len(raw) < 49:  # 实时响应总长 49 字节
            return None
        # 跳过 5 字节帧头(0x50,00,ADDR,cmd,data_len)，去掉 4 字节尾部(sum,xor,0x0D,0x0A)
        data = raw[5:-4]
        # 实时查询响应数据段共 0x28 = 40 字节
        if len(data) < 40:
            return None

        def u2(offset, signed=False):
            return int.from_bytes(data[offset:offset+2], 'big', signed=signed)
        def u4(offset, signed=False):
            return int.from_bytes(data[offset:offset+4], 'big', signed=signed)

        return {
            'diode_set_temp':    u2(0, signed=True)  / 1000.0,   # 二极管设置温度 °C
            'grating_set_temp':  u2(2)                / 1000.0,   # 光栅设置温度 °C（无符号）
            'set_current':       u2(4),                              # 设置电流 mA（硬件单位直接为 mA）
            'power_set':         u2(8)                / 100.0,    # 设置功率
            'power_mode':        data[11],                         # 功率模式开关
            'current_switch':    data[13],                         # 电流开关
            'diode_temp':        u2(15, signed=True) / 1000.0,    # 当前二极管温度 °C
            'grating_temp':      u2(18)               / 1000.0,   # 当前光栅温度 °C（无符号）
            'case_temp':         u2(21, signed=True) / 1000.0,    # 箱体温度 °C
            'actual_current':    u2(23),                            # 当前电流 mA（硬件单位直接为 mA）
            'wavelength':        u4(30)               / 10000.0,   # 波长 nm（四位小数，硬件单位 0.0001nm）
            'power_pd':          u2(36)               / 100.0,    # PD 功率
            'diode_temp2':       u2(38, signed=True) / 1000.0,    # 当前二极管温度2 °C
        }

    def _read_system_info(self) -> bool:
        """读取系统信息（0xAA → 0xB0），缓存 Wave_T_index / Wave_T_offset。
        返回 True 表示读取成功。
        """
        try:
            raw = self._send_frame(self.CMD_SYS_INFO)
            if raw is None:
                self.log("[DFB激光器] 读取系统信息失败：无响应")
                return False
            # 响应总长 133 字节；data 区为 5 字节帧头 + ... + 4 字节帧尾
            data = raw[5:-4]
            if len(data) < 25 + 64:  # 至少要到参数段末尾（设备版本16 + SN8 + ADDR1 + 参数64）
                self.log(f"[DFB激光器] 读取系统信息失败：响应长度不足 (got {len(data)})")
                return False
            # data 布局: 设备版本(16) + SN(8) + ADDR(1) + 参数(64) + ...
            # 参数段起始 offset = 16+8+1 = 25
            params_start = 25
            # Wave_T_index: 参数段 offset 0,1 (2 字节)
            # Wave_T_offset: 参数段 offset 2,3,4,5 (4 字节)
            self._wave_t_index = int.from_bytes(data[params_start:params_start+2], 'big', signed=True)
            self._wave_t_offset = int.from_bytes(data[params_start+2:params_start+6], 'big', signed=True)
            self.log(f"[DFB激光器] 系统信息：Wave_T_index={self._wave_t_index}, Wave_T_offset={self._wave_t_offset}")
            return True
        except Exception as e:
            self.log(f"[DFB激光器] 读取系统信息失败: {e}")
            return False

    # ---------- 对外接口 ----------
    def get_wavelength_nm(self) -> float | None:
        try:
            d = self._query_realtime()
            return d['wavelength'] if d else None
        except Exception as e:
            self.log(f"[DFB激光器] 读取波长失败: {e}")
            return None

    def set_wavelength_nm(self, val_nm: float):
        """设置中心波长 —— 协议 0xA5（光栅温度）
        公式（docs/DFB协议.md 第 127 行）：
            温度*1000 = (波长*10000 - Wave_T_offset) * 10000 / Wave_T_index
        其中 Wave_T_index / Wave_T_offset 来自 0xAA 系统信息，首次调用时自动获取并缓存。
        """
        try:
            if self._wave_t_index is None or self._wave_t_offset is None:
                if not self._read_system_info():
                    self.log(f"[DFB激光器] 设置波长失败：无法读取系统信息 (Wave_T_index/offset)")
                    return
            if self._wave_t_index == 0:
                self.log(f"[DFB激光器] 设置波长失败：Wave_T_index=0，公式除零")
                return
            # 按协议公式计算温度原始值（已带三位小数）
            target_raw = int(round(val_nm * 10000))
            temp_raw = int(round((target_raw - self._wave_t_offset) * 10000 / self._wave_t_index))
            # 2 字节有符号温度区间 [-32768, 32767]，超出范围则截断并告警
            if temp_raw < -32768 or temp_raw > 32767:
                self.log(f"[DFB激光器] 警告：计算温度 {temp_raw} 超 2 字节范围，已截断")
                temp_raw = max(-32768, min(32767, temp_raw))
            # 0xA5 帧 payload: 温度(2B, big, signed) + 保存标志(0/1)
            payload = temp_raw.to_bytes(2, 'big', signed=True) + bytes([0])
            self._send_frame(self.CMD_SET_TEMP_GRATING, payload)
            self.log(f"[DFB激光器] 已设置波长: {val_nm:.6f} nm → 光栅温度 {temp_raw/1000.0:.3f} °C")
            time.sleep(0.5)
        except Exception as e:
            self.log(f"[DFB激光器] 设置波长失败: {e}")

    def get_current_mA(self) -> float | None:
        try:
            d = self._query_realtime()
            return d['actual_current'] if d else None
        except Exception as e:
            self.log(f"[DFB激光器] 读取电流失败: {e}")
            return None

    def set_current_mA(self, val_mA: float):
        """设置电流 —— 协议 0xA3（电流1）
        协议帧: 50 00 ADDR A3 03 电流(2B, 单位 mA) 保存标志(0/1) sum xor 0D 0A
        注意: 硬件直接接收 mA 整数值，不需要 ×100（协议文档中"两位小数"描述有误）。
        协议提示: 电流不建议修改，会影响其他数据参数；
                 注意不能在功率模式开启的时候直接关闭电流（写入非零电流即视为开启）
        """
        try:
            val_raw = int(round(val_mA))  # 硬件单位直接为 mA
            if val_raw < 0 or val_raw > 65535:
                self.log(f"[DFB激光器] 电流 {val_mA:.2f} mA 超出 2 字节范围，已截断")
                val_raw = max(0, min(65535, val_raw))
            # 0xA3 帧 payload: 电流(2B, big, unsigned) + 保存标志(1B, 0=不保存)
            payload = val_raw.to_bytes(2, 'big', signed=False) + bytes([0])
            self._send_frame(self.CMD_SET_CURRENT1, payload)
            self.log(f"[DFB激光器] 已设置电流1: {val_mA:.2f} mA")
            time.sleep(0.5)
        except Exception as e:
            self.log(f"[DFB激光器] 设置电流失败: {e}")

    def get_temperature_c(self) -> float | None:
        """读取光栅当前温度 —— 通过实时查询 grating_temp 字段"""
        try:
            d = self._query_realtime()
            return d['grating_temp'] if d else None
        except Exception as e:
            self.log(f"[DFB激光器] 读取温度失败: {e}")
            return None

    def set_temperature_c(self, val_c: float):
        """设置光栅温度 —— 协议 0xA5
        协议帧: 50 00 ADDR A5 03 温度(2B, signed, ×1000) 保存标志(1B, 0=不保存) sum xor 0D 0A
        """
        try:
            val_raw = int(round(val_c * 1000))  # 温度 ×1000，保留三位小数
            if val_raw < -32768 or val_raw > 32767:
                self.log(f"[DFB激光器] 温度 {val_c:.3f} °C 超出 2 字节有符号范围，已截断")
                val_raw = max(-32768, min(32767, val_raw))
            # 0xA5 帧 payload: 温度(2B, big, signed) + 保存标志(1B, 0=不保存)
            payload = val_raw.to_bytes(2, 'big', signed=True) + bytes([0])
            self._send_frame(self.CMD_SET_TEMP_GRATING, payload)
            self.log(f"[DFB激光器] 已设置光栅温度: {val_c:.3f} °C")
            time.sleep(0.5)
        except Exception as e:
            self.log(f"[DFB激光器] 设置温度失败: {e}")


# ===============  频谱仪控制 & 峰值检测  ===============
class SingleFrequency:
    def __init__(self, ip, timeout_s=60.0, log=print, cmd_map=None):
        self.ip = ip
        self.timeout_s = timeout_s
        self.log = log
        self.rm = None
        self.sa = None
        self.last_rbw_hz = None
        self.last_vbw_hz = None
        self.CMD = {
            'idn': '*IDN?\n',
            'abort': ':ABORt\n',
            'opc': '*OPC?\n',
            'f_center': ':SENSe:FREQuency:CENTer {hz}\n',
            'f_span':   ':SENSe:FREQuency:SPAN {hz}\n',
            'f_start':  ':SENSe:FREQuency:STARt {hz}\n',
            'f_stop':   ':SENSe:FREQuency:STOP {hz}\n',
            'q_start':  ':SENSe:FREQuency:STARt?\n',
            'q_stop':   ':SENSe:FREQuency:STOP?\n',
            'rbw': ':SENSe:BANDwidth:RESolution {hz}\n',
            'vbw': ':SENSe:BANDwidth:VIDeo {hz}\n',
            'rbw?': ':SENSe:BANDwidth:RESolution?\n',
            'vbw?': ':SENSe:BANDwidth:VIDeo?\n',
            'sweep_points?': ':SWEep:POINts?\n',
            'trace_mode_write': ':TRACe:MODE WRITe\n',
            'trace_mode_max':   ':TRACe:MODE MAXHold\n',
            'trace_clear':      ':TRACe:CLEAr\n',
            'trace_data':       ':TRACe:DATA? TRACE1\n',
            'init_once': ':INITiate:IMMediate\n',
            'avg_on':  ':AVERage:STATe ON\n',
            'avg_off': ':AVERage:STATe OFF\n',
            'avg_count': ':AVERage:COUNt {n}\n',
        }
        if cmd_map:
            self.CMD.update(cmd_map)

    def open(self):
        self.rm = pyvisa.ResourceManager()
        #self.sa = self.rm.open_resource(f"TCPIP::{self.ip}::INSTR")
        self.sa = self.rm.open_resource(f"TCPIP::{self.ip}::5025::SOCKET")
        self.sa.timeout = int(self.timeout_s * 1000)
        self.sa.write_termination = '\n'
        self.sa.read_termination = '\n'
        self.sa.encoding = 'latin-1'  # 避免仪器返回非ASCII字节时解码失败

        idn = self.query(self.CMD['idn']).strip()
        self.log(f"[频谱仪] 已连接：{idn}")

        self.write(":CALC:MARK1:MODE NORM")        # 普通标记模式（必须）
        self.write(":CALC:MARK1 ON")                # 打开标记1显示
        self.write(":CALC:MARK1:FUNC NOIS")         # 开启噪声标记（手册第111页精确命令）
        self.log("[频谱仪] 噪声标记已开启 → Nrs dBm/Hz")
        self.write(":CALC:MARK1:MAX")              # 立即跳到最高峰（最实用）
        self.write(":SWE:TYPE:AUTO:RUL DRAN")        # 打开动态范围优先
        self.log("[频谱仪] 动态范围优先已开启")
        self.write(":UNIT:POW DBM")          # 纵轴刻度单位设置为DBM
        self.log("[频谱仪] 纵轴刻度单位已设置为DBM")

        return idn

    def close(self):
        try:
            if self.sa:
                self.sa.close()
        finally:
            if self.rm:
                self.rm.close()

    def write(self, scpi):
        self.sa.write(scpi)

    def query(self, scpi):
        return self.sa.query(scpi)

    def opc(self, label='操作'):
        self.query(self.CMD['opc'])
        self.log(f"[频谱仪] {label} 完成")

    def set_freq_span(self, center=None, span=None, start=None, stop=None):
        if center is not None:
            self.write(self.CMD['f_center'].format(hz=float(center)))
        if span is not None:
            self.write(self.CMD['f_span'].format(hz=float(span)))
        if start is not None:
            self.write(self.CMD['f_start'].format(hz=float(start)))
        if stop is not None:
            self.write(self.CMD['f_stop'].format(hz=float(stop)))

    def set_bw(self, rbw_hz, vbw_hz=None):
        self.write(self.CMD['rbw'].format(hz=float(rbw_hz)))
        time.sleep(0.5)
        self.last_rbw_hz = float(rbw_hz)
        if vbw_hz is not None:
            self.write(self.CMD['vbw'].format(hz=float(vbw_hz)))
            time.sleep(0.5)
            self.last_vbw_hz = float(vbw_hz)
        # 查询实际 RBW（容错：去单位）
        try:
            q = self.CMD.get('rbw?')
            if q:
                resp = self.query(q)
                num = re.findall(r"[-+]?\d*\.?\d+", str(resp))
                if num:
                    self.last_rbw_hz = float(num[0])
        except Exception:
            pass

    def set_avg(self, on=True, count=4):
        self.write(self.CMD['avg_on' if on else 'avg_off'])
        try:
            self.write(self.CMD['avg_count'].format(n=int(count)))
        except Exception:
            pass

    def set_trace_mode(self, max_hold=False):
        self.write(self.CMD['trace_mode_max' if max_hold else 'trace_mode_write'])
        time.sleep(0.5)

    def set_sweep_type(self, sweep_type: str):
        """
        设置扫描优先级
        sweep_type: 'SPD'（速度优先）或 'DYN'（动态范围优先）
        """
        try:
            self.write(f":SWE:TYPE {sweep_type}")
            time.sleep(0.2)
            self.log(f"[频谱仪] 设置扫描优先级为: {sweep_type}")
        except Exception as e:
            self.log(f"[错误] 设置扫描优先级失败: {e}")

    def set_sweep_time(self, sweep_time_s: float):
        """
        设置扫描时间（秒）
        """
        try:
            self.write(f":SWE:TIME {sweep_time_s}")
            time.sleep(0.2)
        except Exception as e:
            self.log(f"[错误] 设置扫描时间失败: {e}")

    def sweep_once(self, label='扫频'):
        """
        改进后的同步扫频逻辑：
        1. 强制关闭连续扫（防止仪器多扫或乱跳）
        2. 发送触发指令
        3. 使用 *OPC? 阻塞等待，直到仪器真正完成
        """
        try:
            # 1. 确保处于单次扫描模式（关键！否则 OPC 查询可能失效或立即返回）
            self.write(':INITiate:CONTinuous OFF')

            # 2. 清除 Trace 数据，防止读到旧的
            self.write(self.CMD['trace_clear'])

            # 3. 触发扫描 (在平均开启时，这会启动完整的一组平均)
            self.write(self.CMD['init_once'])

            # 4. 阻塞等待完成
            current_timeout = self.sa.timeout
            self.sa.timeout = 15000  # 临时设置为 15 秒 (单位毫秒)

            # 发送查询，程序会在这里"卡住"，直到仪器扫完
            self.query('*OPC?')

            # 还原超时时间
            self.sa.timeout = current_timeout

        except Exception as e:
            self.log(f"[错误] 扫频同步超时或失败: {e}")
            # 如果超时，发送中止指令让仪器停下来
            self.write(self.CMD['abort'])

    def set_detector(self, mode: str = "RMS", trace: int = 1):
        """
        设置检波器模式
        常见模式: POSitive, NEGative, SAMPle, RMS
        """
        try:
            self.write(f":DETector:FUNCtion{trace} {mode}")
            time.sleep(0.5)
            self.log(f"[频谱仪] 已设置检波器模式: {mode}")
        except Exception as e:
            self.log(f"[错误] 设置检波器失败: {e}")

    def get_trace_xy(self):
        try:
            raw_data = self.sa.query_ascii_values(self.CMD['trace_data'])

            y_dbm = np.array(raw_data) # 直接转为numpy数组

            f_start = float(self.query(self.CMD['q_start']))
            f_stop = float(self.query(self.CMD['q_stop']))
            n = int(float(self.query(self.CMD['sweep_points?'])))

            x = np.linspace(f_start, f_stop, num=n)
            return x, y_dbm
        except Exception as e:
            try:
                f_start = float(self.query(self.CMD['q_start']))
                f_stop = float(self.query(self.CMD['q_stop']))
                n = int(float(self.query(self.CMD['sweep_points?'])))
                x = np.linspace(f_start, f_stop, num=n)
                return x, np.full(n, -100.0)
            except:
                return np.array([]), np.array([])

    def sweep_continuous_on(self, label="连续粗扫"):
        """开启连续扫描（:INITiate:CONTinuous ON）。尽量先清空 trace，以避免历史峰污染。"""
        try:
            self.write(self.CMD.get('trace_clear', ':TRACe:CLEAr\n'))
            time.sleep(0.5)
        except Exception:
            pass
        try:
            self.write(':INITiate:CONTinuous ON\n')
            time.sleep(0.5)
            self.query_opc(timeout=5.0)
            self.log(f"[频谱仪] 开启连续扫: {label}")
        except Exception as e:
            self.log(f"[频谱仪] 开启连续扫失败: {e}")
            raise

    def sweep_continuous_off(self, label="停止连续粗扫"):
        """关闭连续扫描（:INITiate:CONTinuous OFF）。"""
        try:
            self.write(':INITiate:CONTinuous OFF\n')
            time.sleep(0.5)
            self.log(f"[频谱仪] 关闭连续扫: {label}")
        except Exception as e:
            self.log(f"[频谱仪] 关闭连续扫失败: {e}")

    def query_opc(self, timeout: float | None = None) -> bool:
        """
        等待操作完成 (*OPC?)。
        timeout: 秒 (可选)，如果不传就用 self.timeout_s
        """
        try:
            if self.sa is None:
                raise RuntimeError("未连接频谱仪 (self.sa is None)")

            if timeout is not None:
                self.sa.timeout = int(timeout * 1000)
            else:
                self.sa.timeout = int(self.timeout_s * 1000)

            resp = self.sa.query(self.CMD.get("opc", "*OPC?\n"))
            return resp.strip() == "1"

        except Exception as e:
            self.log(f"[频谱仪] query_opc 失败: {e}")
            return False


class PeakDetector:
    def __init__(self, thresh_db=1.0, prom_db=1.0, guard=10, log_func=print):
        self.thresh_db = float(thresh_db)
        self.prom_db = float(prom_db)
        self.guard = int(guard)
        self.log = log_func

    def find(self, x, y_dbm):
        if len(y_dbm) < 2 * self.guard + 1:
            return []

        edge_points = int(len(y_dbm) * 0.1)
        if edge_points < 10:
            edge_points = 10

        edge_data = np.concatenate([y_dbm[:edge_points], y_dbm[-edge_points:]])
        noise = float(np.mean(edge_data))

        peaks = []
        g = self.guard

        narrow_guard = max(1, int(g / 2))

        for i in range(g, len(y_dbm) - g):
            y = float(y_dbm[i])

            is_local_max = True
            for j in range(1, narrow_guard + 1):
                if y <= float(y_dbm[i-j]) or y <= float(y_dbm[i+j]):
                    is_local_max = False
                    break

            left_nb = y_dbm[max(0, i-g-2):i]
            right_nb = y_dbm[i+1:min(len(y_dbm), i+g+3)]
            left_mean = float(np.mean(left_nb)) if len(left_nb) else noise
            right_mean = float(np.mean(right_nb)) if len(right_nb) else noise

            local_noise = min(left_mean, right_mean, noise)

            if (
                is_local_max and
                (y - local_noise >= self.thresh_db) and
                (y - max(left_mean, right_mean) >= self.prom_db * 0.8)
            ):
                peaks.append((float(x[i]), y, local_noise))
                self.log(f"[峰值检测] 检测到峰值: {x[i]/1e9:.3f} GHz, 功率: {y:.2f} dBm")

        return peaks

    def save_csv_png(self, x, y, peaks, out_dir, name, rbw_hz=1e3):
        os.makedirs(out_dir, exist_ok=True)
        csv_path = os.path.join(out_dir, f'{name}.csv')
        with open(csv_path, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['Frequency(Hz)', 'Power(dBm)'])
            for xi, yi in zip(x, y):
                w.writerow([xi, yi])
        peak_csv = os.path.join(out_dir, f'{name}_peaks.csv')
        with open(peak_csv, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['PeakFreq(Hz)', 'PeakPower(dBm)', 'NoiseFloor(dBm)'])
            for (fx, py, nb) in peaks:
                w.writerow([fx, py, nb])

        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.rcParams['axes.unicode_minus'] = False
        png_path = os.path.join(out_dir, f'{name}.png')
        x_mhz = np.array(x) / 1e6

        fig, ax = plt.subplots(figsize=(12, 6))
        ax.set_facecolor('black')
        ax.plot(x_mhz, y, linewidth=1.2, color='yellow')
        ax.set_xlabel('Frequency (MHz)', fontsize=18)
        ax.set_ylabel('Power (dBm)', fontsize=18)
        ax.margins(x=0)
        ax.grid(linestyle=':', linewidth=0.8, alpha=0.6, color='white')
        ax.tick_params(axis='x', colors='black', size=7, labelsize=15)
        ax.tick_params(axis='y', colors='black', size=7, labelsize=15)

        if peaks:
            main_peak = max(peaks, key=lambda p: p[1])
            for (fx, py, nb) in peaks:
                fx_mhz = fx / 1e6
                ax.axvline(fx_mhz, linestyle='--', linewidth=0.8, color='gray', alpha=0.8)

            fx, py, nb = main_peak
            fx_mhz = fx / 1e6

            lines = [
                f"单频: {fx_mhz:.2f} MHz",
                f"Y: {py:.2f} dBm",
            ]
            txt = "\n".join(lines)

            ax.annotate(
                txt,
                xy=(fx_mhz, py),
                xytext=(8, -6),
                textcoords='offset points',
                ha='left',
                va='top',
                fontsize=12,
                color='black',
                bbox=dict(boxstyle='round,pad=0.3', fc='white', alpha=0.6),
                arrowprops=dict(arrowstyle='->', color='white', lw=0.6)
            )

        plt.tight_layout()
        fig.savefig(png_path, dpi=600)
        plt.close(fig)
        return csv_path, png_path, peak_csv


# ===============  GUI & 流程编排  ===============
class SingleFrequencyGUI:
    def __init__(self, parent=None):
        self.parent = parent

        if parent is None:
            self.root = tk.Tk()
            self.root.title("单频")
            self.root.geometry("1270x950")
            self.root.resizable(True, True)
        else:
            self.root = parent

        self.params_1um = {
            # 仪器 & 输出
            'IP地址': '192.168.7.15',
            '输出目录': r'C:\PTS\zhongzi\SingleFrequency\1.0μm',

            # 串口
            'DFB串口': 'COM1',

            # 测试时长
            '测试时长(分钟)': 30.0,

            # 温度参数
            '温度上限(°C)': 60.0,
            '温度下限(°C)': 18.0,
            '温度步长(°C)': 0.1,
            '温度变化频率(s)': 3.0,

            # 电流参数
            '电流上限(mA)': 600.0,
            '电流下限(mA)': 100.0,
            '电流步长(mA)': 5.0,
            '电流变化频率(s)': 5.0,

            # 扫描参数
            '邻域点数': 10,
            '峰值阈值(dB)': 3.0,
            '邻域显著性(dB)': 5.0,
        }

        self.params_1_5um = {
            # 仪器 & 输出
            'IP地址': '192.168.7.15',
            '输出目录': r'C:\PTS\zhongzi\SingleFrequency\1.5μm',

            # 串口
            'DFB串口': 'COM1',

            # 测试时长
            '测试时长(分钟)': 30.0,

            # 温度参数
            '温度上限(°C)': 56.0,
            '温度下限(°C)': 20.0,
            '温度步长(°C)': 0.1,
            '温度变化频率(s)': 3.0,

            # 电流参数
            '电流上限(mA)': 1400.0,
            '电流下限(mA)': 1400.0,
            '电流步长(mA)': 0.0,
            '电流变化频率(s)': 0.0,

            # 扫描参数
            '邻域点数': 10,
            '峰值阈值(dB)': 3.0,
            '邻域显著性(dB)': 5.0,
        }

        self.test_type_var = tk.StringVar(value="1μm")

        self._build_ui()
        self.stop_flag = threading.Event()
        self.pause_flag = threading.Event()
        self.worker = None

        # 统计计数器
        self.current_cycle_count = 0
        self.temperature_cycle_count = 0
        self.sweep_count = 0
        self.peak_count = 0

        # 实时参数更新定时器 ID
        self._realtime_update_id = None

        # 异步实时参数刷新：后台线程 + 队列
        self._realtime_q = queue.Queue()
        self._realtime_stop = threading.Event()
        self._realtime_thread = None

    # —— UI ——
    def _build_ui(self):
        # 创建主框架，分为左右两部分
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 左侧框架 - 参数设置
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))

        # 右侧框架 - 运行日志
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 右侧顶部横栏：统计数据 + 实时参数 并排
        top_bar = tk.Frame(right_frame)
        top_bar.pack(fill=tk.X, pady=4)

        # 统计数据显示面板（收窄，靠左）
        stats_frame = tk.LabelFrame(top_bar, text='统计数据', padx=6, pady=6)
        stats_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        # 创建统计标签
        self.stats_labels = {}
        stats_items = [
            ('电流循环次数', 'current_cycle_count'),
            ('温度循环次数', 'temperature_cycle_count'),
            ('频率循环次数', 'sweep_count'),
            ('出峰次数', 'peak_count')
        ]

        for i, (label_text, var_name) in enumerate(stats_items):
            label = tk.Label(stats_frame, text=f"{label_text}: ")
            label.grid(row=i, column=0, sticky='w', padx=4, pady=2)

            value_label = tk.Label(stats_frame, text="0", font=("Arial", 10, "bold"))
            value_label.grid(row=i, column=1, sticky='w', padx=4, pady=2)

            self.stats_labels[var_name] = value_label

        # 实时参数显示面板（靠右，占据剩余空间）
        realtime_frame = tk.LabelFrame(top_bar, text='实时参数', padx=6, pady=6)
        realtime_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.realtime_labels = {}
        realtime_items = [
            ('当前温度 (°C)', 'actual_temp'),
            ('设置温度 (°C)', 'set_temp'),
            ('当前电流 (mA)', 'actual_current'),
            ('设置电流 (mA)', 'set_current'),
        ]

        for i, (label_text, key) in enumerate(realtime_items):
            r = i // 2
            c = (i % 2) * 2
            label = tk.Label(realtime_frame, text=f"{label_text}: ")
            label.grid(row=r, column=c, sticky='e', padx=2, pady=2)
            value_label = tk.Label(realtime_frame, text="--", font=("Arial", 10, "bold"))
            value_label.grid(row=r, column=c + 1, sticky='w', padx=2, pady=2)
            self.realtime_labels[key] = value_label

        # 测试类型选择区（居中+外框）- 放在左侧
        type_labelframe = tk.LabelFrame(left_frame, text='测试项选择', padx=10, pady=10)
        type_labelframe.pack(fill=tk.X, pady=8)
        type_frame = tk.Frame(type_labelframe)
        type_frame.pack(expand=True)
        tk.Label(type_frame, text="测试类型:").pack(side=tk.LEFT, padx=6)
        type_combo = tk.OptionMenu(type_frame, self.test_type_var, "1μm", "1.5μm", command=self._on_test_type_change)
        type_combo.pack(side=tk.LEFT, padx=6)
        type_frame.pack(anchor='center')

        # 连接与地址设置区 - 放在左侧上方
        conn_frame = tk.LabelFrame(left_frame, text='连接与地址', padx=8, pady=8)
        conn_frame.pack(fill=tk.X, pady=6)

        # 参数设置区 - 放在左侧下方
        param_frame = tk.LabelFrame(left_frame, text='参数设置', padx=8, pady=8)
        param_frame.pack(fill=tk.X, pady=6)

        self.entries = {}
        self.param_labels = {}
        self.conn_frame = conn_frame
        self.param_frame = param_frame

        # 首次构建UI（使用1μm的参数）
        self._build_param_ui()

        # 按钮区域放在参数设置框下方，居中显示
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=8)

        inner_btn_frame = tk.Frame(btn_frame)
        inner_btn_frame.pack(anchor='center')

        tk.Button(inner_btn_frame, text='开始测试', bg="#4CAF50",fg= "#FFFFFF", command=self.start).pack(side='left', padx=6)
        self.pause_btn = tk.Button(inner_btn_frame, text='暂停', bg="#FFA000", fg="#FFFFFF", command=self._toggle_pause)
        self.pause_btn.pack(side='left', padx=6)
        tk.Button(inner_btn_frame, text='停止测试', bg="#f44336", fg= "#FFFFFF", command=self.stop).pack(side='left', padx=6)

        # 运行日志区 - 放在右侧
        logf = tk.LabelFrame(right_frame, text='运行日志', padx=6, pady=6)
        logf.pack(fill=tk.BOTH, expand=True)
        self.log_box = tk.Text(logf)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _safe_log_append(self, text):
        self.log_box.insert(tk.END, text)
        self.log_box.see(tk.END)

    def log(self, msg):
        t = time.strftime('[%H:%M:%S]')
        self.root.after(0, lambda: self._safe_log_append(f"{t} {msg}\n"))

    def update_stats(self):
        """更新统计数据显示"""
        for var_name, label in self.stats_labels.items():
            if hasattr(self, var_name):
                value = getattr(self, var_name)
                self.root.after(0, lambda l=label, v=value: l.config(text=str(v)))

    def _update_realtime_display(self):
        """GUI 线程回调：非阻塞地消费后台线程产出的实时参数，更新 label。
        串口 IO 已下沉到 _realtime_worker 后台线程，这里只做队列读取与 label 更新。"""
        try:
            while True:
                updates = self._realtime_q.get_nowait()
                for key, text in updates.items():
                    if key in self.realtime_labels:
                        self.realtime_labels[key].config(text=text)
        except queue.Empty:
            pass
        finally:
            if self._realtime_thread is not None and self._realtime_thread.is_alive():
                self._realtime_update_id = self.root.after(1500, self._update_realtime_display)

    def _realtime_worker(self):
        """后台线程：每 1.5s 拉取一次 DFB 实时数据，塞入队列。
        串口访问受控制器内部锁保护，不阻塞 GUI 线程。
        温度数据从 DFB 的 grating_temp / grating_set_temp 获取。"""
        while not self._realtime_stop.is_set():
            updates = {}

            def fmt(v, decimals=2):
                if v is None:
                    return "--"
                return f"{v:.{decimals}f}"

            # DFB（短超时 0.3s，避免设备无响应时长时间持锁阻塞主线程）
            try:
                if hasattr(self, 'lc') and self.lc is not None:
                    dfb_data = self.lc._query_realtime(timeout=0.3)
                    if dfb_data:
                        updates['actual_current'] = fmt(dfb_data.get('actual_current'), 1)
                        updates['set_current'] = fmt(dfb_data.get('set_current'), 1)
                        updates['actual_temp'] = fmt(dfb_data.get('grating_temp'), 2)
                        updates['set_temp'] = fmt(dfb_data.get('grating_set_temp'), 2)
            except Exception:
                pass

            if 'actual_current' not in updates:
                updates['actual_current'] = "--"
            if 'set_current' not in updates:
                updates['set_current'] = "--"
            if 'actual_temp' not in updates:
                updates['actual_temp'] = "--"
            if 'set_temp' not in updates:
                updates['set_temp'] = "--"

            try:
                self._realtime_q.put_nowait(updates)
            except queue.Full:
                pass

            # 可中断的 sleep
            for _ in range(15):
                if self._realtime_stop.is_set():
                    return
                time.sleep(0.1)

    def _save_params(self):
        target = self.params_1um if self.test_type_var.get() == "1μm" else self.params_1_5um
        for k, e in self.entries.items():
            v = e.get()
            try:
                val = float(v)
                target[k] = int(val) if val.is_integer() else val
            except Exception:
                target[k] = v
        self.log('[参数] 已更新')

    def _toggle_pause(self):
        """切换暂停/继续状态。"""
        if not hasattr(self, 'pause_flag'):
            self.pause_flag = threading.Event()
        if not self.pause_flag.is_set():
            # pause
            self.pause_flag.set()
            try:
                self.pause_btn.config(text='继续', bg='#4CAF50')
            except Exception:
                pass
            self.log('[用户] 已暂停，点击继续以恢复')
        else:
            # resume
            self.pause_flag.clear()
            try:
                self.pause_btn.config(text='暂停', bg='#FFA000')
            except Exception:
                pass
            self.log('[用户] 已继续，恢复运行')

    def _pause_point(self):
        """在长循环中调用来实现暂停：当 pause_flag 被设置时阻塞，直到清除或 stop_flag 被设置。"""
        if not hasattr(self, 'pause_flag'):
            return
        while self.pause_flag.is_set():
            time.sleep(0.2)
            if self.stop_flag.is_set():
                raise KeyboardInterrupt

    def start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo('提示', '测试已在进行中')
            return
        self._save_params()
        self.stop_flag.clear()
        self.worker = threading.Thread(target=self._run, daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_flag.set()
        self.log('[用户] 请求停止…')

    # —— 图片弹窗 ——
    def show_image_popup(self, image_path, title="结果预览"):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.resizable(False, False)
        pil_img = Image.open(image_path)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        max_w, max_h = int(sw * 0.8), int(sh * 0.8)
        disp_img = pil_img
        if pil_img.width > max_w or pil_img.height > max_h:
            scale = min(max_w / pil_img.width, max_h / pil_img.height)
            new_size = (int(pil_img.width * scale), int(pil_img.height * scale))
            disp_img = pil_img.resize(new_size, Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(disp_img)
        win.orig_img = pil_img
        win.img_tk = img_tk
        btn_frame = tk.Frame(win)
        btn_frame.pack(side=tk.TOP, fill='x', pady=8)

        def _save_img():
            save_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                     filetypes=[("PNG 文件", "*.png"), ("BMP 文件", "*.bmp"), ("所有文件", "*.*")],
                                                     title="保存图片")
            if save_path:
                try:
                    win.orig_img.save(save_path)
                    messagebox.showinfo("保存成功", f"图片已保存到：{save_path}")
                except Exception as ex:
                    messagebox.showerror("保存失败", str(ex))
        tk.Button(btn_frame, text="保存图片", command=_save_img).pack()
        tk.Label(win, image=win.img_tk).pack(padx=6, pady=6)
        win.update_idletasks()
        w = win.winfo_width(); h = win.winfo_height()
        x = (sw - w) // 2; y = (sh - h) // 2
        win.geometry(f"+{x}+{y}")

    # —— 频谱扫描工具函数 ——
    def _wait_wavelength_stable(self, p: dict, context: str = "") -> bool:
        """等待波长稳定
        返回: True 表示达到稳定，False 表示超时未稳定
        """
        same = 0
        last_val = None
        tol = 0.001
        consec_ok = 3
        max_wait = 300.0
        interval = 0.2

        t0 = time.time()
        while time.time() - t0 < max_wait:
            if getattr(self, 'pause_flag', None) and self.pause_flag.is_set():
                self._pause_point()
            if self.stop_flag.is_set():
                raise KeyboardInterrupt

            wl = self.lc.get_wavelength_nm()
            if wl is None:
                time.sleep(interval)
                continue

            if last_val is not None and abs(wl - last_val) < tol:
                same += 1
                if same >= consec_ok:
                    self.log(f"[等待稳定{context}] 波长已稳定 (连续{consec_ok}次Δ<{tol}nm)")
                    return True
            else:
                same = 0

            last_val = wl
            if same > 0:
                self.log(f"[等待稳定{context}] 当前波长 {wl:.4f} nm, 已连续稳定 {same}/{consec_ok}")
            time.sleep(interval)

        self.log(f"[等待稳定{context}] 超时未稳定")
        return False

    # —— 主流程 ——
    def _run(self):
        # 根据当前选择的测试类型选择参数集
        p = self.params_1um if self.test_type_var.get() == "1μm" else self.params_1_5um
        out_dir = os.path.abspath(str(p['输出目录']))
        os.makedirs(out_dir, exist_ok=True)

        # 清空输出文件夹
        if os.path.exists(out_dir):
            self.log(f"[测试] 正在清空输出文件夹: {out_dir}")
            for item in os.listdir(out_dir):
                item_path = os.path.join(out_dir, item)
                if os.path.isfile(item_path):
                    os.remove(item_path)
                elif os.path.isdir(item_path):
                    import shutil
                    shutil.rmtree(item_path)
            self.log(f"[测试] 输出文件夹清空完成")

        # 获取测试时长参数
        test_duration_min = float(p.get('测试时长(分钟)', 30.0))
        test_duration_sec = test_duration_min * 60
        start_time = time.time()
        end_time = start_time + test_duration_sec
        self.log(f"[测试] 开始时间: {time.strftime('%H:%M:%S', time.localtime(start_time))}")
        self.log(f"[测试] 预计结束时间: {time.strftime('%H:%M:%S', time.localtime(end_time))}")
        self.log(f"[测试] 测试时长: {test_duration_min:.1f} 分钟")

        sa = SingleFrequency(ip=str(p['IP地址']), timeout_s=60.0, log=self.log)

        # DFB 激光器初始化（串口 RS-485，温度通过 0xA5 命令控制光栅温度）
        lc = DFBLaserController(port=str(p['DFB串口']), log_func=self.log)

        try:
            # 连接设备
            lc.open()
            self.lc = lc
            sa.open()
            sa.set_avg(on=True, count=2)

            # 启动异步实时参数刷新：后台线程做串口 IO，GUI 线程只消费队列
            self._realtime_stop.clear()
            self._realtime_thread = threading.Thread(target=self._realtime_worker, daemon=True)
            self._realtime_thread.start()
            self._realtime_update_id = self.root.after(0, self._update_realtime_display)

            # 读取初始状态（仅用于记录）
            wl0 = self.lc.get_wavelength_nm()
            cur0 = self.lc.get_current_mA()
            temp0 = self.lc.get_temperature_c()
            self.log(f"[DFB激光器] 初始状态：波长 {wl0} nm，电流 {cur0} mA，温度 {temp0} °C")

            # 获取温度相关参数
            temp_max = float(p.get("温度上限(°C)", 30.0))
            temp_min = float(p.get("温度下限(°C)", 20.0))
            temp_step = float(p.get("温度步长(°C)", 1.0))
            temp_freq = float(p.get("温度变化频率(s)", 5.0))

            # 获取电流相关参数
            cur_max = float(p.get("电流上限(mA)", 600.0))
            cur_min = float(p.get("电流下限(mA)", 100.0))
            cur_step = float(p.get("电流步长(mA)", 50.0))
            cur_freq = float(p.get("电流变化频率(s)", 5.0))

            # 初始化温度和电流
            temp = temp_min
            cur = cur_min

            # 设置初始温度和电流
            self.log(f"[DFB激光器] 设置初始温度: {temp:.2f} °C")
            lc.set_temperature_c(temp)

            self.log(f"[DFB激光器] 设置初始电流: {cur:.2f} mA")
            lc.set_current_mA(cur)

            # 初始化细扫参数
            sa.write(":TRACe1:MODE WRITe")
            sa.write(":TRACe:CLEar TRACE1")
            sa.set_trace_mode(max_hold=False)
            sa.write(":INITiate:CONTinuous OFF")
            sa.set_avg(on=True, count=2)
            sa.set_detector("RMS", trace=1)
            sa.write(":BANDwidth:VIDeo:RATIO 1")
            sa.set_sweep_type('SPD')
            sa.set_sweep_time(1)

            span = 500.0 * 1e6
            step = 500.0 * 1e6
            f_start = 0.0 * 1e6
            f_stop = 18000.0 * 1e6
            center = f_start + span / 2.0

            # 先设置频宽为500MHz，再设置RBW为30kHz
            sa.set_freq_span(center=center, span=span)
            sa.set_bw(rbw_hz=30.0 * 1e3)

            # ---- 线程 1：温度控制（DFB 0xA5 光栅温度命令） ----
            def temperature_control_thread():
                if temp_step <= 0 or temp_min >= temp_max:
                    # 步长为 0 或上下限相同 → 只设一次，不循环
                    self.log(f"[DFB激光器] 固定温度: {temp_min:.2f} °C")
                    while not self.stop_flag.is_set():
                        try:
                            if self.pause_flag.is_set():
                                time.sleep(0.2)
                                continue
                            time.sleep(0.5)
                        except Exception:
                            time.sleep(1.0)
                    return

                local_temp = temp_min
                temp_increasing = True
                prev_temp_increasing = True
                last_update = time.time()

                while not self.stop_flag.is_set():
                    try:
                        if self.pause_flag.is_set():
                            time.sleep(0.2)
                            continue

                        current_time = time.time()
                        if current_time - last_update >= temp_freq:
                            if temp_increasing:
                                local_temp += temp_step
                                if local_temp >= temp_max:
                                    local_temp = temp_max
                                    temp_increasing = False
                            else:
                                local_temp -= temp_step
                                if local_temp <= temp_min:
                                    local_temp = temp_min
                                    temp_increasing = True

                            if temp_increasing != prev_temp_increasing:
                                if temp_increasing:
                                    self.temperature_cycle_count += 1
                                    self.log(f"[统计] 温度循环次数: {self.temperature_cycle_count}")
                                    self.update_stats()
                                prev_temp_increasing = temp_increasing

                            self.log(f"[DFB激光器] 设置温度: {local_temp:.2f} °C")
                            lc.set_temperature_c(local_temp)
                            last_update = current_time

                        time.sleep(0.1)
                    except Exception as e:
                        self.log(f"[错误] 温度控制线程出错：{e}")
                        time.sleep(1.0)

            # ---- 线程 2：电流控制（DFB 0xA3 命令） ----
            def current_control_thread():
                if cur_step <= 0 or cur_min >= cur_max:
                    # 步长为 0 或上下限相同 → 只设一次，不循环
                    self.log(f"[DFB激光器] 固定电流: {cur_min:.2f} mA")
                    while not self.stop_flag.is_set():
                        try:
                            if self.pause_flag.is_set():
                                time.sleep(0.2)
                                continue
                            time.sleep(0.5)
                        except Exception:
                            time.sleep(1.0)
                    return

                local_cur = cur_min
                cur_increasing = True
                prev_cur_increasing = True
                last_update = time.time()

                while not self.stop_flag.is_set():
                    try:
                        if self.pause_flag.is_set():
                            time.sleep(0.2)
                            continue

                        current_time = time.time()
                        if current_time - last_update >= cur_freq:
                            if cur_increasing:
                                local_cur += cur_step
                                if local_cur >= cur_max:
                                    local_cur = cur_max
                                    cur_increasing = False
                            else:
                                local_cur -= cur_step
                                if local_cur <= cur_min:
                                    local_cur = cur_min
                                    cur_increasing = True

                            if cur_increasing != prev_cur_increasing:
                                if cur_increasing:
                                    self.current_cycle_count += 1
                                    self.log(f"[统计] 电流循环次数: {self.current_cycle_count}")
                                    self.update_stats()
                                prev_cur_increasing = cur_increasing

                            self.log(f"[DFB激光器] 设置电流: {local_cur:.2f} mA")
                            lc.set_current_mA(local_cur)
                            last_update = current_time

                        time.sleep(0.1)
                    except Exception as e:
                        self.log(f"[错误] 电流控制线程出错：{e}")
                        time.sleep(1.0)

            # 启动温度 + 电流两个独立线程
            threading.Thread(target=temperature_control_thread, daemon=True).start()
            threading.Thread(target=current_control_thread, daemon=True).start()

            # 主循环：持续进行细扫
            while not self.stop_flag.is_set():
                # 检查测试时长是否已到
                current_time = time.time()
                if current_time >= end_time:
                    self.log(f"[测试] 测试时长已到，结束测试")
                    self.stop_flag.set()
                    break

                # 检查暂停状态
                if self.pause_flag.is_set():
                    self._pause_point()

                # 执行细扫
                sa.set_freq_span(center=center, span=span)

                # 每次新跨度开始前彻底清屏 + 重新开平均
                sa.write(":TRACe:CLEar TRACE1")
                sa.write(":AVERage:COUNt 2")
                sa.write(":AVERage:STATe ON")

                for repeat in range(2):
                    sa.set_sweep_time(1)
                    sa.sweep_once(f'细扫@{center/1e9:.3f}GHz')
                    self.log(f"细扫@{center/1e9:.3f}GHz")
                    x, y = sa.get_trace_xy()

                    # 细扫峰值检测
                    fine_peak = PeakDetector(thresh_db=float(p['峰值阈值(dB)']), prom_db=float(p['邻域显著性(dB)']), guard=int(p['邻域点数']), log_func=self.log)
                    peaks = fine_peak.find(x, y)
                    if peaks:
                        # 获取实际温度和电流值
                        actual_temp = self.lc.get_temperature_c()
                        actual_cur = self.lc.get_current_mA()

                        temp_str = f"{actual_temp:.3f}" if actual_temp is not None else "unknown"
                        cur_str = f"{actual_cur:.1f}" if actual_cur is not None else "unknown"

                        # 保存数据，包含温度和电流信息
                        tag = f"T{temp_str}C_I{cur_str}mA"
                        tag2 = f"fine_{tag}_{int(center/1e6)}MHz"
                        rbw_used = sa.last_rbw_hz if getattr(sa, 'last_rbw_hz', None) else 30.0 * 1e3
                        csvp, pngp, peakcsv = fine_peak.save_csv_png(x, y, peaks, out_dir, tag2, rbw_hz=rbw_used)
                        self.log(f"[细扫] 命中异常峰，保存：{tag2}.csv/.png/_peaks.csv")

                        # 出峰次数统计
                        self.peak_count += 1
                        self.log(f"[统计] 出峰次数: {self.peak_count}")
                        self.update_stats()
                        break

                # 更新细扫中心频率
                center += step
                if center - span / 2.0 >= f_stop:
                    # 细扫完成一轮，重新开始
                    center = f_start + span / 2.0
                    self.sweep_count += 1
                    self.log(f"[统计] 频率循环次数: {self.sweep_count}")
                    self.update_stats()

            self.log("— 全部流程结束 —")

        except StopIteration as stop_reason:
            self.log(f'[终止] {stop_reason}')
            self.root.after(0, lambda: messagebox.showinfo('终止', str(stop_reason)))
        except KeyboardInterrupt:
            self.log('[停止] 用户终止。')
            self.root.after(0, lambda: messagebox.showinfo('已停止', '已根据请求停止扫描。'))
        except Exception as e:
            self.log(f'[错误] 测试失败：{e}')
            self.root.after(0, lambda err=str(e): messagebox.showerror('错误', err))
        finally:
            # 先停止实时参数刷新后台线程，避免与恢复/关闭操作抢串口
            try:
                self._realtime_stop.set()
                if self._realtime_thread is not None:
                    self._realtime_thread.join(timeout=2.0)
                    self._realtime_thread = None
                if self._realtime_update_id is not None:
                    self.root.after_cancel(self._realtime_update_id)
                    self._realtime_update_id = None
                try:
                    while True:
                        self._realtime_q.get_nowait()
                except queue.Empty:
                    pass
            except Exception:
                pass

            try:
                # 恢复初始状态
                if wl0 is not None:
                    self.log("[恢复] 正在将设备恢复到初始状态...")
                    self.lc.set_wavelength_nm(wl0)
                    if self._wait_wavelength_stable(p, " - 恢复初始波长"):
                        self.log(f"[恢复] 波长已恢复到初始值: {wl0:.6f} nm")
                    else:
                        self.log("[警告] 恢复初始波长超时")

                if cur0 is not None:
                    self.lc.set_current_mA(cur0)
                    time.sleep(2.0)
                    self.log(f"[恢复] 电流已恢复到初始值: {cur0:.2f} mA")

                if temp0 is not None:
                    self.lc.set_temperature_c(temp0)
                    time.sleep(2.0)
                    self.log(f"[恢复] 温度已恢复到初始值: {temp0:.2f} °C")

                self.log("[恢复] 设备已恢复到初始状态")
            except Exception as e:
                self.log(f"[警告] 恢复初始状态时出错: {e}")
            finally:
                try:
                    sa.close()
                except Exception:
                    pass
                try:
                    lc.close()
                except Exception:
                    pass
                # 重置实时参数标签为 "--"
                for key in self.realtime_labels:
                    self.realtime_labels[key].config(text="--")

    def run(self):
        self.root.mainloop()

    def _on_test_type_change(self, event=None):
        if self.test_type_var.get() == "1μm":
            self.params = self.params_1um.copy()
        else:
            self.params = self.params_1_5um.copy()
        self._refresh_entries()

    def _refresh_entries(self):
        # 清除旧的UI元素
        for widget in self.conn_frame.winfo_children():
            widget.destroy()
        for widget in self.param_frame.winfo_children():
            widget.destroy()

        self.entries.clear()
        self.param_labels.clear()

        # 重新构建UI
        self._build_param_ui()

    def _get_serial_ports(self):
        """获取系统可用串口列表"""
        try:
            import serial.tools.list_ports
            ports = [p.device for p in serial.tools.list_ports.comports()]
            if not ports:
                ports = ['COM1']
            return ports
        except Exception:
            return ['COM1', 'COM2', 'COM3', 'COM4']

    def _refresh_serial_ports(self):
        """刷新串口下拉框的可用端口列表"""
        ports = self._get_serial_ports()
        for key in ['DFB串口']:
            widget = self.entries.get(key)
            if widget is not None:
                current = widget.get()
                widget['values'] = ports
                if current and current not in ports:
                    widget.set(ports[0])

    def _build_param_ui(self):
        """构建参数UI界面"""
        current_params = self.params_1um if self.test_type_var.get() == "1μm" else self.params_1_5um

        # 串口列表
        serial_ports = self._get_serial_ports()
        serial_field_keys = {'DFB串口'}

        # 连接与地址参数
        conn_params = ['IP地址', '输出目录', 'DFB串口']
        for i, k in enumerate(conn_params):
            if k not in current_params:
                continue
            label = tk.Label(self.conn_frame, text=k)
            label.grid(row=i, column=0, sticky='e')

            if k in serial_field_keys:
                # 串口下拉框 + 刷新按钮放在一个子 Frame 中
                serial_row = tk.Frame(self.conn_frame)
                serial_row.grid(row=i, column=1, padx=4, pady=1, sticky='w')
                e = ttk.Combobox(serial_row, values=serial_ports, state='normal', width=12)
                e.set(str(current_params[k]))
                e.pack(side='left')
                refresh_btn = tk.Button(serial_row, text='刷新', width=4,
                                        command=self._refresh_serial_ports)
                refresh_btn.pack(side='left', padx=(4, 0))
            else:
                # 普通输入框
                e = tk.Entry(self.conn_frame, width=28)
                e.insert(0, str(current_params[k]))
                e.grid(row=i, column=1, padx=4, pady=2)

            self.entries[k] = e
            self.param_labels[k] = label

        # 其他参数
        other_params = [k for k in current_params.keys() if k not in conn_params]
        for i, k in enumerate(other_params):
            label = tk.Label(self.param_frame, text=k)
            label.grid(row=i, column=0, sticky='e')
            e = tk.Entry(self.param_frame, width=23)
            e.insert(0, str(current_params[k]))
            e.grid(row=i, column=1, padx=4, pady=2)
            self.entries[k] = e
            self.param_labels[k] = label


if __name__ == '__main__':
    gui = SingleFrequencyGUI()
    gui.run()
