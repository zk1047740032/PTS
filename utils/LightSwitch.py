#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
光开关控制模块

通过 USB VISA 协议控制光开关，支持通道查询与切换。
SCPI 指令:
    - 查询当前通道: :OSW1:CHAN?
    - 设置通道:      :OSW1:CHAN n
"""

import sys, pathlib
# 确保能 import core（脚本独立运行时不以包形式组织）
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from core import VisaInstrument
from core.config import CFG
from typing import Optional, Callable


class OpticalSwitch(VisaInstrument):
    """
    光开关控制类（继承 VisaInstrument）

    通过 VISA (USB) 连接光开关设备，支持通道查询与切换。
    self.visa_resource 为业务字段，对应基类 self.address。

    注意：连接验证用 :OSW1:CHAN? 而非 *IDN?（原版行为），故 connect 传 idn=False
    关闭基类 *IDN? 验证后，由本方法自行 get_channel() 验证，保持零行为变更。

    使用示例:
        sw = OpticalSwitch(CFG.usb.optical_switch)  # USB0::0x0005::0x0012::87104000113::INSTR
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
        super().__init__(log_func=log_func, address=visa_resource, timeout_ms=CFG.timing.visa_short_ms)
        self.visa_resource = visa_resource  # 业务字段，对应基类 self.address

    def connect(self, visa_resource: Optional[str] = None) -> bool:
        """
        连接光开关设备

        Args:
            visa_resource: VISA 资源字符串，未传时用构造时传入的 self.visa_resource。

        Returns:
            True 表示连接成功，False 表示失败
        """
        addr = visa_resource or self.visa_resource
        # idn=False 关闭基类 *IDN? 验证（光开关用 :OSW1:CHAN? 验证），max_retries=0 走单次
        ok = super().connect(addr, idn=False)
        if not ok:
            return False
        try:
            # 验证连接：查询当前通道
            ch = self.get_channel()
            self.log(f"[光开关] 连接成功，当前通道: {ch}")
            return True
        except Exception as e:
            self.log(f"[光开关] 连接失败: {e}")
            super().close()
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
        """关闭光开关连接（复用 VisaInstrument.close，释放 inst 与 rm）"""
        super().close()
        self.log("[光开关] 连接已关闭")

if __name__ == "__main__":
    sw = OpticalSwitch("USB0::0x0005::0x0012::87104000113::INSTR", print)
    sw.connect()