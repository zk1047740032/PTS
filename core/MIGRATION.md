# 迁移手册 / Migration Guide

将 `path_a/` 与 `path_b/` 下的存量测试程序逐步迁移到 `core/` 基类。
本文档是跨对话续作的唯一进度来源——**每迁移完一个文件，必须更新本文件的 checklist**。

---

## 设计理念 / Design Principles

> 详见 `CLAUDE.md` 与 memory `rin-migration-principles`。核心三条：

1. **零行为变更 (Zero behavior change)**：迁移后程序行为与原版逐字节一致。
2. **向后兼容优先于对称美 (Compatibility over symmetry)**：用参数和钩子吸收"脏"差异（方法名不一、地址格式不一、有无重试），不强迫现状就范。
3. **渐进迁移 (Incremental)**：一个文件一个文件迁，每个都可独立验证。

---

## 每文件迁移 Checklist / Per-file Checklist

迁移任意一个文件时，按此顺序执行，每步完成打勾。

### A. 分析阶段
- [ ] 通读该文件全部类，识别"控制器类"(连 VISA 的) 与 "GUI 类"
- [ ] 找出本文件特有差异：方法名(connect/instrument/open)、地址格式、有无重试、stop_flag 是 bool 还是 Event、自动关窗延时值、是否有多线程编排

### B. 改造阶段
- [ ] 文件头：删 `ctypes` DPI 块、删 `import pyvisa`、加 `from core import ...`
- [ ] 控制器类：继承 `VisaInstrument`，`__init__` 调 `super().__init__(log_func=, timeout_ms=, address=)`
- [ ] 控制器类：`connect()` 改为调 `super().connect(visa_address(...), idn=...)`，业务字段名(如 `self.instrument`)设为 `self.inst` 的别名
- [ ] 控制器类：`close()` 删手写样板，若保留业务别名则重写并调用 `super().close()` + 清别名
- [ ] 截图读回：用 `read_instrument_screenshot()`，**透传原临时路径**不要改路径
- [ ] trace 两列写：用 `write_xy_csv()`；单行带表头：用 `append_row_csv()`
- [ ] 清空目录：用 `clear_directory(dir, log_func=)`
- [ ] GUI 类：继承 `BaseTestGUI`，窗口构造交给 `super().__init__(parent, title=, geometry=, icon=)`
- [ ] GUI 类：删除自己的 `log()` 与 `run()`，复用基类
- [ ] GUI 类：`self.root.after(N, self.root.destroy)` 改为 `self.auto_close(N)`

### C. 不擅自套用的情况 (Do NOT force-fit)
- [ ] `TestRunner` 这类协调器：若对外暴露固定签名方法被 GUI 带参调用，**不套** `BaseTestRunner`，保留原结构（除非同时改调用方）
- [ ] 复杂多线程编排(如 `SingleFrequency` 双并行控制线程 + 实时轮询)：可不套 runner，只用 `VisaInstrument` + `BaseTestGUI`
- [ ] 子类方法签名/职责与基类不同时(如按 param_key 而非 entry 对象操作)，保留子类方法，不强用基类版本

### D. 验证阶段
- [ ] 语法检查：`python -c "import ast; ast.parse(open('path_a/XXX.py',encoding='utf-8').read())"`
- [ ] 导入 + 继承：`python` 内联脚本 import 该模块，断言 `issubclass(子类, 基类)`
- [ ] connect/close 行为：mock pyvisa（见 memory `rin-verification-pattern`），断言别名指向 inst、close 清别名
- [ ] GUI 冒烟：构造 GUI、确认按钮已创建、log 能写入、`root.after(100, destroy); run()` 干净退出
- [ ] 原行为对比（需仪器时跳过）：保持逻辑分支、数值、日志文案不变

### E. 收尾
- [ ] 更新本文件"进度"表格的状态
- [ ] git 暂不提交（除非用户要求）

---

## 进度 / Progress

状态：`done` / `in_progress` / `todo` / `skip`(不迁)

| 文件 | 控制器类 | GUI 类 | 状态 | 备注 |
|------|---------|--------|------|------|
| path_a/Rin_FSV3004.py | RinAnalyzer, BackgroundNoiseAnalyzer | RinGUI | done | TestRunner 保留未套模板；stop_flag 保持 bool |
| path_a/Power.py | PowerMeterController | PowerGUI | done | PowerCollector/burnin 纯计算保留；connect 保留子类实现不设终止符 |
| path_a/PhaseNoise.py | (无仪器,用 pywinauto) | PhaseNoiseGUI | done | click_button 模块级函数；无 VISA，只用 BaseTestGUI；start_test 保留自定义 |
| path_a/LightSwitch.py | OpticalSwitch | (无GUI) | done | USB VISA 完整地址透传；connect 传 idn=False 后用 get_channel 自行验证(零行为变更) |
| path_a/LineWidth_FSV3004.py | SignalGenerator, LinewidthTester | LineWidth_FSV3004_GUI | done | SignalGenerator/LinewidthTester 继承 VisaInstrument；GUI 继承 BaseTestGUI；SignalGenerator 与 path_b 重复，后续可消除 |
| path_a/SingleFrequency.py | DFBLaserController(serial), SingleFrequency(VISA) | SingleFrequencyGUI | done | DFB serial 不继承；SingleFrequency→VisaInstrument；GUI→BaseTestGUI；双线程编排+PeakDetector用write_xy_csv |
| path_a/SpectrumSNR.py | SpectrumSNR | SpectrumSNRGUI | done | connect 使用 visa_address(tcpip_instr)；截图保留 MMEM:STORe:GRAPhics BMP 原逻辑(非 HCOPy PNG)；GUI 线程简单直接编排未套 Runner |
| path_a/TimeDomain.py | TimeDomain | TimeDomainGUI | done | 双仪器：scope→VisaInstrument(主)，gen 手动管理；GUI→BaseTestGUI |
| path_b/WaveLength.py | SignalGenerator, WavemeterController, WlmController | WaveLengthTestGUI | todo | **SignalGenerator 与 path_a 重复**；WavemeterController 走 wlmData DLL 非 VISA |

---

## 关键设计决策记录 / Key Decisions Log

> 这些决策难以从代码直接看出，是新对话继续迁移时必须知道的。亦见 memory `rin-migration-principles`、`rin-migration-techniques`。

| 决策 | 原因 |
|------|------|
| 控制器用 `self.instrument` 等业务别名指向 `self.inst` | 存量业务方法大量引用 `self.instrument`，改全部引用成本高、易错。别名保零行为变更 |
| `RinAnalyzer.stop_flag` 保持 `bool` 而非基类的 `Event` | `TestRunner` 直接真值判断 `ra.stop_flag`，改类型会破坏停止逻辑 |
| `connect` 默认 `idn=False` (Rin/BNA) | 原 Rin/BNA 的 connect 不查 *IDN?，保持一致；需 IDN 的(如 SignalGenerator)传 `idn=True` |
| 截图临时路径透传，不用工具默认 | 仪器本地路径 `C:\PTS\Rin\_temp.png` 等是实测可用的，改路径有仪器侧风险 |
| `TestRunner` 暂不继承 `BaseTestRunner` | 它对外暴露 `run_rin(ra,ui_root,ip)`/`run_background(...)`/`stop()` 固定签名被 GUI 带参调用，套模板需同时改两边 |
| 日志从同步改为基类异步 `root.after` | 功能等价，后台线程更安全；文案/时间戳格式不变 |
| `PowerMeterController.connect` 保留子类实现不调 `super().connect` | 原实现不设读/写终止符（只设 timeout）且失败抛异常（调用方 try/except）。基类 connect 会强制设 `\n` 终止符并吞异常返回 bool，会改变 USB 功率计通信行为与失败流程。故保留原 connect 体（内部直接用 `pyvisa`，需保留 `import pyvisa`），仅 `__init__` 走基类获取 rm/inst/log/timeout_ms。GUI 的 `list_visa_resources` 也直接用 pyvisa |
| `OpticalSwitch.connect` 传 `idn=False` 后自行 `get_channel()` 验证 | 原版连接验证用 `:OSW1:CHAN?` 而非 `*IDN?`。基类 `idn` 参数只支持 `*IDN?`，故关闭基类 IDN 验证，由子类在 connect 成功后调 `get_channel()` 复刻原验证逻辑（含失败时 `super().close()` 与返回 False），保持零行为变更。USB 完整资源串直接透传 `super().__init__(address=...)`，不套 `visa_address()` |

---

## 验证脚本模板 / Verification Script Template

mock pyvisa 验证 connect/close 通用模板（替换模块名与方法名）：

```python
import importlib.util, sys, pathlib, types
sys.path.insert(0, str(pathlib.Path('.').resolve()))
import core.base_instrument as bi
class FakeRM:
    def open_resource(self, addr):
        r = types.SimpleNamespace()
        def q(c): return 'Fake,Model,1,1' if c=='*IDN?' else '1'
        r.query=q; r.write=lambda c:None; r.clear=lambda:None
        r.read_termination=None; r.write_termination=None; r.timeout=None; r.addr=addr
        r.close=lambda: None
        return r
    def close(self): pass
bi.pyvisa = types.SimpleNamespace(ResourceManager=FakeRM)
spec = importlib.util.spec_from_file_location('mod', 'path_a/XXX.py')
m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
# ... 实例化、connect、断言别名、close、断言清理
```

GUI 冒烟模板：

```python
import importlib.util, sys, pathlib
sys.path.insert(0, str(pathlib.Path('.').resolve()))
spec = importlib.util.spec_from_file_location('mod', 'path_a/XXX.py')
m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
g = m.XxxGUI()
for btn in ['btn1','btn2']: assert hasattr(g, btn)
g.root.after(100, g.root.destroy); g.run()
```
