# CLAUDE.md — 项目指引 / Project Guide for Claude

本文件会在 Claude Code 每次启动时自动加载。**修改 `path_a/`、`path_b/`、`core/` 相关代码前必读**。

## 项目简介 / Overview

PTS 种子激光器自动化测试系统。Windows 桌面应用，Python + tkinter。
- `path_a/`、`path_b/`：多个独立测试程序（光开关/功率计/相位噪声/线宽/RIN/单频/频谱SNR/时域/波长计），每个都是 GUI + VISA 仪器控制 + 后台线程 + CSV/PNG 落盘。
- `core/`：从上述程序抽象出的通用基类与工具（utils / VisaInstrument / BaseTestGUI / BaseTestRunner）。

## 当前主线任务 / Current Work

**正在将 `path_a/`、`path_b/` 存量程序渐进迁移到 `core/` 基类。**

- **进度与操作手册**：见 [core/MIGRATION.md](core/MIGRATION.md) ——跨对话续作的唯一进度来源。**每迁移完一个文件，必须更新那里的 checklist 与进度表。**
- 已迁移：`path_a/Rin_FSV3004.py`（其余 8 个 todo）。

## 迁移铁律 / Migration Invariants（违反即破坏设计）

1. **零行为变更**：迁移后程序行为与原版一致。不改测试逻辑、不改数值、不改日志文案。
2. **向后兼容优先于对称美**：用参数和钩子吸收存量差异（方法名、地址格式、有无重试、stop_flag 类型），**不**为整洁而重构现状。
3. **业务字段别名保留**：存量方法引用 `self.instrument`/`self.osa`/`self.scope` 等，迁移时把这些设为 `self.inst` 的别名，不要改名牵连全部引用。
4. **不擅自套模板**：
   - `TestRunner` 这类被 GUI 带参调用的协调器，不套 `BaseTestRunner`（除非同时改调用方）。
   - 复杂多线程编排（如 `SingleFrequency`）可只用 `VisaInstrument` + `BaseTestGUI`，不用 runner。
   - 子类方法签名与基类不同时，保留子类方法。
5. **一个文件迁完一个文件验证**：语法 + 导入 + 继承 + mock connect/close + GUI 冒烟，全过才算完。

## 续作指引 / How to Continue (new session)

当用户在新对话要求"继续迁移"时：
1. 先读 [core/MIGRATION.md](core/MIGRATION.md) 的进度表，确认下一个 `todo` 文件。
2. 该文件 controller 类参考已迁移的 `path_a/Rin_FSV3004.py`（`RinAnalyzer` / `BackgroundNoiseAnalyzer` 是范本）。
3. GUI 类参考 `path_a/Rin_FSV3004.py` 的 `RinGUI`。
4. 按 MIGRATION.md 的 checklist A→E 执行，用其中的验证脚本模板。
5. 迁完更新 MIGRATION.md 进度表。

## code/ 文档语言约定

- 结构标题用英文，说明用中文。
- 代码注释中文（与存量一致）。

## 依赖 / Dependencies

仅 `pyvisa` / `tkinter` / 标准库 + 既有 `numpy`/`PIL`/`pywinauto`/`serial`/`matplotlib`。**core 基类不引入新依赖**。
