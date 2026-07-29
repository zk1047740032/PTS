"""配置管理 API"""

import json
import sys
from pathlib import Path

from fastapi import APIRouter, HTTPException

router = APIRouter(prefix="/api/config", tags=["config"])

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
_USER_CONFIG_PATH = _PROJECT_ROOT / "config" / "user_config.json"


def _load_user_config() -> dict:
    """读取用户配置文件"""
    if _USER_CONFIG_PATH.exists():
        with open(_USER_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_user_config(data: dict):
    """保存用户配置文件"""
    _USER_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_USER_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _get_default_config() -> dict:
    """从 CFG 读取默认配置值"""
    if str(_PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(_PROJECT_ROOT))
    from core.config import CFG
    import dataclasses

    def dataclass_to_dict(obj) -> dict:
        result = {}
        for field in dataclasses.fields(obj):
            value = getattr(obj, field.name)
            if dataclasses.is_dataclass(value):
                result[field.name] = dataclass_to_dict(value)
            else:
                result[field.name] = value
        return result

    return dataclass_to_dict(CFG)


@router.get("")
async def get_config():
    """获取完整配置（默认值 + 用户覆盖）"""
    defaults = _get_default_config()
    overrides = _load_user_config()
    return {"defaults": defaults, "overrides": overrides}


@router.get("/user")
async def get_user_overrides():
    """获取用户配置覆盖"""
    return _load_user_config()


@router.put("/user")
async def save_user_overrides(data: dict):
    """保存用户配置覆盖（合并模式）"""
    current = _load_user_config()
    current.update(data)
    _save_user_config(current)
    return {"saved": True, "overrides": current}


@router.post("/user/reset")
async def reset_user_overrides(data: dict | None = None):
    """重置用户配置

    Body (可选): {"keys": ["network.fsv3004_rin", ...]} — 重置指定键
    不传 body 则全部重置
    """
    if data and data.get("keys"):
        current = _load_user_config()
        for k in data["keys"]:
            current.pop(k, None)
        _save_user_config(current)
        return {"reset": True, "overrides": current}
    else:
        _save_user_config({})
        return {"reset": True, "overrides": {}}
