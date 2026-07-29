"""PTS Web 启动器 — 双击运行，后台启动服务 + 自动打开浏览器"""

import subprocess
import sys
import time
import webbrowser
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# 构建前端静态文件（如果还没构建过）
dist_dir = Path(__file__).resolve().parent / "frontend" / "dist"
if not dist_dir.exists():
    print("首次运行，正在构建前端...")
    subprocess.run(
        ["npm", "run", "build"],
        cwd=str(Path(__file__).resolve().parent / "frontend"),
        shell=True,
    )

# 启动 uvicorn（隐藏窗口）
uvicorn_cmd = [
    sys.executable, "-m", "uvicorn",
    "web.backend.main:app",
    "--host", "0.0.0.0",
    "--port", "8088",
]

print("正在启动 PTS Web 服务...")
proc = subprocess.Popen(
    uvicorn_cmd,
    cwd=str(PROJECT_ROOT),
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
)

# 等服务器就绪
time.sleep(2)

# 打开浏览器
url = "http://127.0.0.1:8088"
print(f"服务已启动，正在打开浏览器: {url}")
webbrowser.open(url)

print("关闭此窗口不会停止服务，服务在后台运行。")
print("要停止服务，请到任务管理器结束 python.exe 进程。")

# pyw 扩展名自动无控制台窗口
input("按 Enter 退出（服务将继续在后台运行）...")
