#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core.base_instrument —— VISA 仪器控制器基类

吸收 path_a / path_b 中各仪器控制类的共性：
SignalGenerator / LinewidthTester / RinAnalyzer / BackgroundNoiseAnalyzer /
SpectrumSNR / SingleFrequency / OpticalSwitch / PowerMeterController。

子类只需实现自己的业务方法（configure/measure/set_channel 等），
连接、释放、SCPI 读写、停止标志等通用行为由本基类承担。
"""

from __future__ import annotations

import threading
import time
from typing import Any, Callable, Optional

import pyvisa

__all__ = ["VisaInstrument"]


class VisaInstrument:
    """
    VISA 仪器控制器基类。

    统一封装 ResourceManager 管理、仪器连接（可选重试、可选 *IDN? 验证）、
    SCPI 读写、VISA 缓存清理、停止标志、资源释放。

    设计取舍：
        用 ``connect`` 的 max_retries 参数吸收"有重试/无重试"差异。
        现有 8 个控制器中仅 2 个用到重试（path_b/WaveLength.py 与
        path_a/LineWidth_FSV3004.py 的 SignalGenerator），不值得为此分叉
        出两个子类层级。无重试需求者 max_retries=0 走单次即可。

    使用示例::

        class PowerMeterController(VisaInstrument):
            def __init__(self, resource, log_func=print, timeout_ms=5000):
                super().__init__(log_func=log_func, address=resource,
                                 timeout_ms=timeout_ms)
            def query_idn(self):
                return self.query("*IDN?").strip()

    属性:
        log: 日志回调。
        rm: pyvisa.ResourceManager，连接成功后赋值。
        inst: 已打开的仪器资源对象，未连接时为 None。
        address: 当前/待连接的 VISA 地址。
        timeout_ms: 通信超时（毫秒）。
        stop_flag: threading.Event，用于线程安全的中断通知。
        _idn: 成功连接且 idn=True 时缓存的仪器标识字符串。
    """

    def __init__(
        self,
        *,
        log_func: Optional[Callable[[str], None]] = print,
        address: Optional[str] = None,
        timeout_ms: int = 10000,
    ) -> None:
        """
        初始化仪器控制器基类。

            参数:
                log_func (callable, optional): 日志回调，默认 print。传 None 则静默。
                address (str, optional): VISA 资源地址。connect() 时若未显式传入则用此值。
                timeout_ms (int): 通信超时时间（毫秒），默认 10000。
        """
        self.log: Callable[[str], None] = log_func or (lambda msg: None)
        self.rm: Optional[pyvisa.ResourceManager] = None
        self.inst: Optional[Any] = None
        self.address: Optional[str] = address
        self.timeout_ms: int = timeout_ms
        self.stop_flag = threading.Event()
        self._idn: Optional[str] = None

    # ------------------------------------------------------------------
    # 连接 / 释放
    # ------------------------------------------------------------------
    def connect(
        self,
        address: Optional[str] = None,
        *,
        max_retries: int = 0,
        retry_interval: int = 2,
        idn: bool = True,
        idn_ok: Optional[Callable[[str], bool]] = None,
    ) -> bool:
        """
        连接仪器：打开 VISA 资源，设置超时与终止符，可选 *IDN? 验证。

            带逐次重试。无重试需求者 max_retries=0（默认）。

            参数:
                address (str, optional): VISA 资源地址。为 None 时用构造时传入的 self.address。
                max_retries (int): 失败后的最大重试次数，默认 0（单次）。
                retry_interval (int): 重试间隔（秒），默认 2。
                idn (bool): 是否在连接后查询 *IDN? 验证，默认 True。
                idn_ok (callable, optional): 对 *IDN? 返回值的校验回调，返回 False 视为连接失败。
                    用于捕获"返回错误信息但未抛异常"的半连接状态（如某些仪器返回 '0'）。

            返回:
                bool: True 表示连接成功（并通过可选的 idn_ok 校验），False 表示失败。
        """
        addr = address or self.address
        for attempt in range(max_retries + 1):
            try:
                try:
                    self.rm = pyvisa.ResourceManager("@py")
                except Exception:
                    self.rm = pyvisa.ResourceManager()
                self.inst = self.rm.open_resource(addr)
                self.inst.timeout = self.timeout_ms
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'

                self._idn = None
                if idn:
                    self._idn = self.inst.query("*IDN?").strip()
                    if idn_ok is not None and not idn_ok(self._idn):
                        raise RuntimeError(f"*IDN? 校验未通过: {self._idn!r}")

                self.address = addr
                return True
            except Exception as e:
                if attempt < max_retries:
                    self.log(
                        f"[仪器] 第{attempt + 1}次连接失败：{e}，{retry_interval}秒后重试..."
                    )
                    time.sleep(retry_interval)
                else:
                    self.log(f"[仪器] 连接失败，共尝试{max_retries + 1}次")
                    return False
        return False

    def close(self) -> None:
        """
        关闭仪器连接并释放 ResourceManager。

        对 inst.close 与 rm.close 各自 try/except，保证一个失败不影响另一个，
        最后将两者置 None。与现有所有 close() 写法行为一致。
        """
        if self.inst:
            try:
                self.inst.close()
            except Exception:
                pass
            self.inst = None
        if self.rm:
            try:
                self.rm.close()
            except Exception:
                pass
            self.rm = None

    # ------------------------------------------------------------------
    # SCPI 读写
    # ------------------------------------------------------------------
    def write(self, scpi: str) -> None:
        """发送 SCPI 写指令。仪器未连接时由 pyvisa 抛 RuntimeError。"""
        self.inst.write(scpi)

    def query(self, scpi: str) -> str:
        """发送 SCPI 查询指令并返回响应字符串。"""
        return self.inst.query(scpi)

    def query_binary(self, scpi: str, **kwargs: Any) -> list:
        """发送 SCPI 查询并以二进制方式读取返回值。

        kwargs 透传给 pyvisa 的 query_binary_values（如 datatype/container/is_big_endian）。
        用于读取 Trace 数据、截图等二进制响应。
        """
        return self.inst.query_binary_values(scpi, **kwargs)

    def clear(self) -> None:
        """清理 VISA 缓存，避免上一条失败查询污染后续操作。提取自 LinewidthTester。"""
        if self.inst:
            try:
                self.inst.clear()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # 中断控制
    # ------------------------------------------------------------------
    def stop(self) -> None:
        """置位停止标志，通知测量线程终止后续操作。"""
        self.stop_flag.set()
        self.log("[仪器] 停止信号已设置")

    def clear_stop(self) -> None:
        """清除停止标志，用于新一轮测量前复位。"""
        self.stop_flag.clear()

    # ------------------------------------------------------------------
    # 便利属性
    # ------------------------------------------------------------------
    @property
    def is_connected(self) -> bool:
        """是否已打开仪器资源。"""
        return self.inst is not None

    @property
    def idn(self) -> Optional[str]:
        """缓存的 *IDN? 标识字符串（连接成功且 idn=True 时填充）。"""
        return self._idn
