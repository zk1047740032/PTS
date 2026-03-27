# AGENTS.md - Development Guidelines for PTS

This document provides guidance for agentic coding agents operating in this repository.

## Project Overview

PTS (PreciTestSystem) is an automated test platform for laser testing applications, built with Python/tkinter. The system supports multiple test modules for RIN, linewidth, time domain, SNR, single frequency, and power measurements.

## Build, Lint, and Test Commands

### Running the Application
```bash
python main_platform.py
```

### Building Executable (PyInstaller)
```bash
pyinstaller package.spec
```
Output is placed in `dist/` directory.

### Running Individual Test Scripts
```bash
python test/收发指令.py     # Instrument communication test
python test/数据读取.py     # Data reading/analysis test
python test/WaveLength.py  # Wavelength test utility
python test/找按钮.py       # Button locator utility
```

### Running a Single Test
There is no formal test framework (pytest/unittest). To run a specific test script:
```bash
python test/<test_name>.py
```

## Code Style Guidelines

### General Principles
- Use `from __future__ import annotations` for forward references
- Keep code modular with clear separation between GUI and business logic
- Use Chinese docstrings for classes and functions

### Import Order
1. Standard library (`os`, `sys`, `time`, `threading`, etc.)
2. `typing` module
3. Third-party libraries (`pyvisa`, `numpy`, `matplotlib`, etc.)
4. Local modules (`from path_a import ...`, `from report import ...`)

```python
# Correct order example
import os
import time
from typing import Optional, List

import numpy as np
import pyvisa
from PIL import Image

from path_a.Rin_FSV3004 import RinGUI
from report.template_generator import generate_report
```

### Type Annotations
- All function parameters and return values MUST have type annotations
- Use `Optional[X]` instead of `X | None`
- Use `List[X]`, `Dict[K, V]` from typing

```python
def connect(self, ip: str, timeout: float = 60.0) -> bool:
    ...

def log(self, module: str, msg: str, level: str = "info", file_path: Optional[str] = None) -> None:
    ...
```

### Naming Conventions
| Type | Convention | Example |
|------|------------|---------|
| Classes | CamelCase | `Rin_4051_GUI`, `IntegratedPlatform` |
| Functions/Methods | snake_case | `connect()`, `start_test()` |
| Constants | UPPER_SNAKE_CASE | `DEFAULT_IP`, `MODULE_MAP` |
| Instance variables | snake_case | `self.log_callback`, `self.worker_thread` |
| Modules | snake_case | `path_a/`, `template_generator.py` |

### Function and Class Structure
- Keep functions focused (single responsibility)
- Use clear, descriptive names
- Include docstrings explaining parameters and return values
- Group related functions into classes

### Error Handling
- Use try/except blocks with specific exception types
- Log errors before re-raising or returning
- Provide meaningful error messages

```python
try:
    self.inst = self.rm.open_resource(res_str)
except Exception as e:
    self.log(f"连接失败: {e}")
    self.close()
    return False
```

### GUI Development (tkinter)
- Use `ttk` widgets when possible for better styling
- Handle DPI scaling for Windows high-DPI displays:

```python
if os.name == 'nt':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        dpi = ctypes.windll.user32.GetDpiForSystem()
        scaling_factor = dpi / 96.0
    except Exception:
        scaling_factor = 1.0
```

- Use `threading.Thread` with `daemon=True` for background tasks
- Never block the main GUI thread with long operations

### File Paths
- Use raw strings for Windows paths: `r"C:\PTS\zhongzi\Rin\..."`
- Use `os.path.join()` for path concatenation
- Create directories with `os.makedirs(path, exist_ok=True)` before writing

### Testing Guidelines
- Test scripts go in `test/` directory
- Use descriptive names: `<feature>_test.py` or `<test_描述>.py`
- Keep tests independent and idempotent

### Documentation
- Include Chinese docstrings for all public classes and functions
- Document parameters, return values, and exceptions
- Add inline comments for complex logic only

### Multi-Process Architecture
The main platform uses multiprocessing for parallel test execution:
- Each test module runs in a separate process
- Communication via `multiprocessing.Queue`
- Use `daemon=True` for worker processes

```python
p = multiprocessing.Process(
    target=run_module_process,
    args=(name, start_method, self.msg_queue, cmd_q),
    daemon=True
)
p.start()
```

### PyInstaller Configuration
When adding new dependencies, update `package.spec`:
- Add to `hiddenimports` list
- Add any required data files to `datas` tuple

```python
a = Analysis(['main_platform.py'],
             ...
             hiddenimports=['pyvisa', 'new_package', ...],
             datas=[('PreciLasers.ico', '.')],
             ...)
```

### Key Directories
| Directory | Purpose |
|-----------|---------|
| `path_a/` | Seed source test modules (RIN, linewidth, power, etc.) |
| `path_b/` | Current temperature test modules |
| `test/` | Test utilities and scripts |
| `report/` | Report generation templates and output |
| `drivers/` | Instrument driver interfaces |
| `build/` | PyInstaller build artifacts |
| `dist/` | Packaged executable output |

### Common Third-Party Libraries
- `pyvisa` - Instrument communication (SCPI)
- `numpy` - Numerical computing
- `matplotlib` - Plotting (use `Agg` backend for non-GUI)
- `docxtpl` - Word report generation
- `PIL` - Image processing
- `pywinauto` - Windows automation
