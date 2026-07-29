"""PTS Web 后端 — FastAPI 入口"""

import sys
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from .config import CORS_ORIGINS, HOST, PORT, STATIC_DIR
from .database import init_db

# 确保项目根目录在 sys.path 中（供子进程等使用）
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))


@asynccontextmanager
async def lifespan(app: FastAPI):
    """应用生命周期：启动时初始化数据库"""
    await init_db()
    yield


app = FastAPI(
    title="PTS 种子激光器测试系统",
    description="PreciLasers Seed Laser Automated Test System — Web Frontend",
    version="1.0.0",
    lifespan=lifespan,
)

# CORS — 开发时允许所有来源
app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ========== REST API 路由 ==========
from .routers.platform import router as platform_router
from .routers.test_run import router as test_run_router
from .routers.instruments import router as instruments_router
from .routers.config import router as config_router
from .routers.rin import router as rin_router

app.include_router(platform_router)
app.include_router(test_run_router)
app.include_router(instruments_router)
app.include_router(config_router)
app.include_router(rin_router)

# ========== WebSocket 路由 ==========
from .ws.test_log import router as ws_router
from .ws.agent import router as agent_ws_router

app.include_router(ws_router)
app.include_router(agent_ws_router)


# ========== 健康检查 ==========
@app.get("/api/health")
async def health():
    return {
        "status": "ok",
        "service": "PTS Web Backend",
        "version": "1.0.0",
    }


# ========== 生产模式：挂载前端静态文件 ==========
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="frontend")


# ========== 启动入口 ==========
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "web.backend.main:app",
        host=HOST,
        port=PORT,
        reload=True,
        log_level="info",
    )
