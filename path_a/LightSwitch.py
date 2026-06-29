#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
光开关控制模块

通过 USB VISA 协议控制光开关，支持通道查询与切换。
SCPI 指令:
    - 查询当前通道: :OSW1:CHAN?
    - 设置通道:      :OSW1:CHAN n
"""

import pyvisa
from typing import Optional, Callable


class OpticalSwitch:
    """
    光开关控制类

    通过 VISA (USB) 连接光开关设备，支持通道查询与切换。

    使用示例:
        sw = OpticalSwitch("USB0::0x0005::0x0012::87104000113::INSTR")
        if sw.connect():
            ch = sw.get_channel()
            sw.set_channel(2)
            sw.close()
    """

    def __init__(self,
                 visa_resource: str,
                 log_func: Callable[[str], None] = print):
        """
        初始化光开关控制器

        Args:
            visa_resource: VISA 资源字符串，如 "USB0::0x0005::0x0012::87104000113::INSTR"
            log_func: 日志回调函数
        """
        self.visa_resource = visa_resource
        self.log = log_func
        self.rm: Optional[pyvisa.ResourceManager] = None
        self.inst: Optional[pyvisa.ResourceManager] = None

    def connect(self) -> bool:
        """
        连接光开关设备

        Returns:
            True 表示连接成功，False 表示失败
        """
        try:
            self.rm = pyvisa.ResourceManager()
            self.inst = self.rm.open_resource(self.visa_resource)
            self.inst.timeout = 5000  # 5 秒超时
            self.inst.write_termination = '\n'
            self.inst.read_termination = '\n'

            # 验证连接：查询当前通道
            ch = self.get_channel()
            self.log(f"[光开关] 连接成功，当前通道: {ch}")
            return True
        except Exception as e:
            self.log(f"[光开关] 连接失败: {e}")
            self.inst = None
            return False

    def get_channel(self) -> int:
        """
        查询当前激活的通道

        Returns:
            当前通道编号 (1, 2, 3, ...)

        Raises:
            Exception: 设备未连接或通信失败
        """
        if not self.inst:
            raise Exception("光开关未连接")
        resp = self.inst.query(":OSW1:CHAN?")
        return int(resp.strip())

    def set_channel(self, channel: int) -> None:
        """
        切换光开关到指定通道，并验证切换结果

        Args:
            channel: 目标通道编号 (1, 2, 3, ...)

        Raises:
            Exception: 设备未连接、通信失败或切换后验证不匹配
        """
        if not self.inst:
            raise Exception("光开关未连接")

        self.inst.write(f":OSW1:CHAN {channel}")
        self.log(f"[光开关] 发送切换指令 -> 通道 {channel}")

        # 验证切换结果
        actual = self.get_channel()
        if actual != channel:
            raise Exception(
                f"光开关通道切换失败: 期望通道 {channel}，实际通道 {actual}"
            )
        self.log(f"[光开关] 切换确认: 通道 {actual}")

    def close(self) -> None:
        """关闭光开关连接"""
        try:
            if self.inst:
                self.inst.close()
                self.log("[光开关] 连接已关闭")
        except Exception:
            pass
        finally:
            self.inst = None
            self.rm = None
