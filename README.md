# PreciTestSystem (PTS)

**频准自动化测试系统**

> 上海频准激光科技股份有限公司

PreciTestSystem（PTS）是上海频准激光科技股份有限公司研发的自动化测试平台，主要应用于激光器（种子源）和电流温度（CT）控制下的性能测试。系统通过图形化界面实现测试流程控制，支持多种仪器通信协议，生成标准化的测试报告。

## 主要功能

### 种子源测试 (zhongzi)
- **RIN 测试** - 相对强度噪声分析
- **线宽测试** - 激光器线宽测量
- **时域分析** - 时域响应测试
- **信噪比测试** - 光谱信噪比分析
- **单频测试** - 单频激光器特性测试
- **功率测试** - 光功率测量

### 电流温度测试 (qijian)
- **CT-波长** - 电流温度下的波长测试
- **CT-线宽** - 电流温度下的线宽测试
- **CT-功率** - 电流温度下的功率测试

## 技术架构

- **GUI框架**: tkinter
- **仪器通信**: pyvisa
- **数据处理**: numpy, pandas
- **可视化**: matplotlib (TkAgg后端)
- **报告生成**: docxtpl
- **Windows自动化**: pywinauto
- **打包工具**: PyInstaller

## 项目结构

```
PTS/
├── main_platform.py       # 主程序入口
├── zhongzi/               # 种子源测试模块
├── qijian/                # 电流温度测试模块
├── test/                  # 测试工具脚本
├── report/                # 测试报告输出目录
├── build/                 # PyInstaller构建目录
├── dist/                  # 打包输出目录
├── package.spec           # PyInstaller配置文件
└── 开发文档/              # 开发文档
```

## 快速开始

### 环境配置

```bash
# 创建虚拟环境
python -m venv .venv

# 激活虚拟环境
.venv\Scripts\activate  # Windows

# 安装依赖
pip install pyvisa numpy matplotlib pillow pywinauto pandas docxtpl
```

### 运行程序

```bash
python main_platform.py
```

### 打包发布

```bash
# 使用spec文件打包
pyinstaller package.spec

# 或直接打包
pyinstaller --onefile --windowed --icon=PreciLasers.ico main_platform.py
```

打包后的可执行文件位于 `dist/` 目录。

## 仪器配置

系统支持以下仪器：
- FSV3004 频谱分析仪 (IP: 192.168.7.10, Port: 5025)
- 4051 激光功率计
- CY4052D 线宽测试仪

具体IP地址和参数可在各测试模块中配置。

## 测试工具

### 通信测试
```bash
python test/收发指令.py
```
用于测试与仪器的通信连接，发送SCPI指令并获取响应。

### 数据读取
```bash
python test/数据读取.py
```
用于读取和分析测试生成的CSV数据文件。

## 开发指南

### 代码规范
- 使用 `from __future__ import annotations` 启用延迟类型注解
- 类和函数使用中文docstring
- 导入顺序：标准库 → typing → 第三方库 → 本地模块
- 函数参数和返回值必须使用类型注解

### 命名约定
- 类名：`CamelCase`（如 `RinAnalyzer`）
- 函数/方法：`snake_case`（如 `connect()`）
- 常量：`UPPER_SNAKE_CASE`

### Windows DPI适配
项目包含Windows高DPI屏幕适配代码，确保界面清晰显示。

## 许可证

仅供内部使用

## 联系方式

技术支持请联系开发团队
