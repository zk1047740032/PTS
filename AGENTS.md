# AGENTS.md - PreciTestSystem (PTS)

## Overview
PreciTestSystem (频准自动化测试系统) is a Python-based test system for optical equipment (laser testing) developed by Shanghai Preci Laser Technology Co., Ltd. (上海频准激光科技股份有限公司). "Preci" is the English name for "频准". It uses tkinter for GUI, pyvisa for instrument communication, and PyInstaller for building executables.

## Project Structure
```
PTS/
├── main_platform.py       # Main entry point - Integrated GUI platform
├── zhongzi/               # Core test modules (种子源测试)
│   ├── Rin_FSV3004.py     # RIN (Relative Intensity Noise) test
│   ├── LineWidth_FSV3004.py
│   ├── TimeDomain.py
│   ├── SpectrumSNR.py
│   └── ...
├── qijian/                # CT (Current & Temperature) modules
│   ├── CT_W.py            # CT Wavelength test
│   ├── CT_L.py            # CT Linewidth test
│   └── CT_P.py            # CT Power test
├── test/                  # Manual test scripts (数据读取.py, 收发指令.py)
├── report/                # Generated test reports
├── build/                 # PyInstaller build output
├── dist/                  # Built executables
├── package.spec           # PyInstaller spec file
└── 开发文档/              # Documentation in Chinese
```

## Commands

### Running the Application
```bash
# Activate virtual environment
.venv\Scripts\activate  # Windows

# Run main platform
python main_platform.py
```

### Building Executable (PyInstaller)
```bash
# Build using the spec file
pyinstaller package.spec

# Or direct command
pyinstaller --onefile --windowed --icon=PreciLasers.ico main_platform.py
```

### Running Tests
Tests in this project are manual scripts, not pytest/unittest:
```bash
# Test instrument communication
python test/收发指令.py

# Test data reading
python test/数据读取.py

# Run a specific module for testing
python -c "from zhongzi.Rin_FSV3004 import RinAnalyzer"
```

### Virtual Environment
```bash
# Create venv (if needed)
python -m venv .venv

# Install dependencies
pip install pyvisa numpy matplotlib pillow pywinauto pandas docxtpl
```

## Code Style Guidelines

### File Headers
```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module description in Chinese
"""
from __future__ import annotations
```

### Import Order
1. Standard library (os, time, csv, threading, etc.)
2. `typing` module
3. Third-party libraries (pyvisa, numpy, matplotlib, tkinter, etc.)
4. Local/relative imports

### Type Hints
Always use type hints for function parameters and return values:
```python
def connect(self, ip_address: str = "192.168.7.10", port: int = 5025) -> bool:
    ...
```

### Docstrings
Use Chinese docstrings for all public methods and classes:
```python
def method_name(self, param: str) -> bool:
    """
    方法功能描述
    
    参数:
        param (str): 参数说明
        
    返回:
        bool: 返回值说明
    """
```

### Naming Conventions
- Classes: `CamelCase` (e.g., `RinAnalyzer`, `LaserController`)
- Functions/methods: `snake_case` (e.g., `connect()`, `run_module_process()`)
- Constants: `UPPER_SNAKE_CASE` (e.g., `DEFAULT_TIMEOUT`)
- Private methods: prefix with `_` (e.g., `_internal_method()`)

### Error Handling
Always use try/except with specific exception handling:
```python
try:
    # operation
    self.log("成功执行操作")
    return True
except Exception as e:
    self.log(f"操作失败: {e}")
    return False
```

### Windows-Specific Code
```python
# DPI awareness for Windows
if os.name == 'nt':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        dpi = ctypes.windll.user32.GetDpiForSystem()
        scaling_factor = dpi / 96.0
    except Exception:
        scaling_factor = 1.0
else:
    scaling_factor = 1.0
```

### Logging
Use a logger function passed in or default to `print`:
```python
def __init__(self, log_func=default_logger):
    self.log = log_func

def default_logger(msg: str):
    print(msg)
```

### GUI (tkinter) Conventions
- Use `matplotlib.use('TkAgg')` before importing pyplot
- Set Chinese fonts for matplotlib:
```python
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False
```

### Module Organization
1. Imports
2. Windows DPI setup
3. Helper functions
4. Main class(es)
5. `if __name__ == "__main__":` block (if needed)

### Constants and Configuration
Hardcode default values in classes, use meaningful variable names:
```python
self.dc_value = 1.20  # 默认DC值
self.amplification = 14
self.file_wait_timeout_s = 30.0
self.file_wait_poll_s = 0.5
self.ip_address = "192.168.7.10"
```

### Threading
Use threading for non-blocking operations:
```python
import threading
self.stop_flag = False
```

### Instrument Communication
Use pyvisa for instrument control:
```python
rm = pyvisa.ResourceManager()
inst = rm.open_resource(f"TCPIP0::{ip_address}::{port}::SOCKET")
inst.timeout = 60000
inst.read_termination = '\n'
inst.write_termination = '\n'
```

### File Paths
Use raw strings for Windows paths:
```python
file_path = r"C:\PTS\zhongzi\Rin\FSV3004\Rin_1.DAT"
```

### Git Workflow
- Create feature branches for new features
- Commit messages in Chinese or English
- Test changes before committing

## Key Dependencies
- pyvisa - Instrument communication
- numpy - Numerical computing
- matplotlib - Plotting (TkAgg backend)
- pillow - Image processing
- pywinauto - Windows UI automation
- pandas - Data handling
- docxtpl - Word report generation
- pyinstaller - Building executables
