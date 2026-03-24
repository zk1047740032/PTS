import os
import sys
from drivers import wlmData
from drivers import wlmConst

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

class SignalGenerator:
    def __init__(self, log_callback=None) -> None:
        """
        初始化信号发生器控制类

        参数:
            log_callback (callable, optional): 日志回调函数，用于输出日志信息
        """
        self.rm = None # 存储Visa资源管理器
        self.inst = None # 存储仪器实例
        self.log = log_callback or (lambda msg: None)

    def connect(self, ip_address, max_retries=3, retry_interval=2):
        """
        连接信号发生器（添加重试机制）

        参数:
            ip_address (str): 信号发生器的 IP 地址
            max_retries (int, optional): 最大重试次数，默认3次
            retry_interval (int, optional): 重试间隔时间（秒），默认2秒
        返回:
            bool: 连接成功返回 True，失败返回 False
        """
        for attempt in range(max_retries+1):
            try:
                self.rm = pyvisa.ResourceManager()
                self.inst = self.rm.open_resource(f'TCPIP0::{ip_address}::INSTR')
                self.inst.timeout = 10000 # 设置通信超时时间，防止仪器无响应时程序无限等待
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'
                # 验证连接是否可用
                self.inst.query('*IDN?')
                self.log(f"[信号源] 已连接到信号发生器（第{attempt+1}次尝试）")
                return True
            except Exception as e:
                if attempt < max_retries:
                    self.log(f"[信号源] 第{attempt+1}次连接失败：{e}，{retry_interval}秒后重试...")
                    time.sleep(retry_interval)
                else:
                    self.log(f"[信号源] 连接失败，共尝试{max_retries}次，请尝试手动断开重连")
                    return False
        return False

    def configure(self, waveform="RAMP", freq=0.1, volt=10, offset=5):
        """
        配置信号发生器输出参数

        参数:
            waveform (str, optional): 波形类型，如"SIN"表示正弦波。默认"SIN"。
            freq (float, optional): 输出频率，单位Hz。默认0.1 Hz。
            volt (float, optional): 输出幅值，单位Vpp。默认10 Vpp。
            offset (float, optional): 直流偏移，单位Vdc。默认5 Vdc。

        返回:
            bool: 配置成功返回True，失败返回False。
        """
        if not self.inst:
            self.log("[信号源] 未连接到信号发生器")
            return False
        try:
            # 设置波形类型（RAMP=三角波）
            self.inst.write(f":SOUR1:FUNC {waveform}")
            self.log(f"[信号源] 设置波形: {waveform}")
            
            # 设置频率（单位：Hz）
            self.inst.write(f":SOUR1:FREQ {freq}Hz")
            self.log(f"[信号源] 设置频率: {freq} Hz")
            
            # 设置幅值（单位：Vpp）
            self.inst.write(f":SOUR1:VOLT {volt}")
            self.log(f"[信号源] 设置幅值: {volt} Vpp")
            
            # 设置偏移（单位：Vdc）
            self.inst.write(f":SOUR1:VOLT:OFFS {offset}")
            self.log(f"[信号源] 设置偏移: {offset} Vdc")
            
            return True
        except Exception as e:
            self.log(f"[信号源] 配置失败: {e}")
            return False

    def set_output(self, on=True):
        '''
        设置信号发生器输出开关状态

        参数:
            on (bool, optional): 是否打开输出。True 表示打开，False 表示关闭。默认 True。

        返回:
            bool: 设置成功返回 True，失败返回 False。
        '''
        if not self.inst:
            self.log("[信号源] 未连接到信号发生器")
            return False
        try:
            if on:
                self.inst.write(":OUTP1 ON")
                self.log("[信号源] 打开输出")
            else:
                self.inst.write(":OUTP1 OFF")
                self.log("[信号源] 关闭输出")
            return True
        except Exception as e:
            self.log(f"[信号源] 设置输出状态失败: {e}")
            return False

    def close(self):
        """
        关闭信号发生器连接与资源管理器，释放VISA资源。
        """
        if self.inst:
            try:
                self.inst.close()
                self.log("[信号源] 已关闭信号发生器连接")
                self.inst = None
            except Exception:
                pass
        if self.rm:
            try:
                self.rm.close()
                self.inst = None
            except Exception:
                pass

class HighFinesseWLM:
    def __init__(self, log_func=print):
        self.dll = wlmData.LoadDLL()
        self.log = log_func

if __name__ == "__main__":
    WLM = HighFinesseWLM()
    version_type = WLM.dll.GetWLMVersion(0)
    version_ver = WLM.dll.GetWLMVersion(1)
    version_rev = WLM.dll.GetWLMVersion(2)
    version_build = WLM.dll.GetWLMVersion(3)
    print(f'WLM version: [{version_type}.{version_ver}.{version_rev}.{version_build}]')