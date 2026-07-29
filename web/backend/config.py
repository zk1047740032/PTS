"""Web 后端专用配置"""

import os
import pathlib

# 项目根目录
PROJECT_ROOT = pathlib.Path(__file__).resolve().parent.parent.parent

# 数据库
DATABASE_URL = os.getenv(
    "PTS_DATABASE_URL",
    f"sqlite+aiosqlite:///{PROJECT_ROOT / 'data' / 'pts_test_results.db'}",
)

# Web 服务
HOST = os.getenv("PTS_WEB_HOST", "127.0.0.1")
PORT = int(os.getenv("PTS_WEB_PORT", "8080"))
CORS_ORIGINS = os.getenv("PTS_CORS_ORIGINS", "*").split(",")

# 前端静态文件（生产模式）
STATIC_DIR = PROJECT_ROOT / "web" / "frontend" / "dist"
