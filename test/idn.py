import pyvisa
import numpy as np
import os
import matplotlib.pyplot as plt

class Instrument():
    def __init__(self):
        self.ip = None
        self.rm = None
        self.inst = None

    def connect(self, ip):
        self.ip = ip
        self.rm = pyvisa.ResourceManager()
        self.inst = self.rm.open_resource(f"TCPIP0::{ip}::INSTR")
        self.inst.timeout = 1000
        self.inst.writetermination = '\n'
        self.inst.readtermination = '\n'

    def write(self, CMD):
        if self.inst:
            self.inst.write(CMD)
            return True
        else:
            print("仪器未连接")
            return False

    def query_setting(self, CMD):
        try:
            res = self.inst.query(CMD)
            return res
        except Exception as e:
            print(f"查询错误：{e}")

    def trac_plot(self):
        self.inst.write("FORM ASC")
        data_str = self.inst.query("TRAC:DATA? TRACE1")
        y = [float(v) for v in data_str.strip().split(',')]
        center = float(self.inst.query("FREQ:CENTER?"))
        span = float(self.inst.query("FREQ:SPAN?"))
        freq_start = center - span/2
        freq_end = center + span/2
        x = np.linspace(freq_start, freq_end, len(y))
        plt.plot(x, y)
        plt.show()