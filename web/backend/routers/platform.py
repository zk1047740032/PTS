"""主平台 API 路由"""

from fastapi import APIRouter, HTTPException

from ..services.test_service import (
    MODULE_REGISTRY,
    MODULE_CHANNEL_MAP,
    MODULE_START_METHODS,
    test_service,
)

router = APIRouter(prefix="/api", tags=["platform"])

# 通道分组
CHANNEL_GROUPS = {
    "光路A": {
        "通道1": ["PZT调制"],
        "通道2": ["时域"],
        "通道3": ["Rin_FSV3004", "线宽_FSV3004", "信噪比", "单频"],
        "通道4": ["相噪"],
    },
    "光路B": {
        "功率": ["功率"],
    },
}


@router.get("/modules")
async def list_modules():
    """返回所有测试模块列表"""
    modules = []
    for key, (module_path, gui_class) in MODULE_REGISTRY.items():
        modules.append({
            "key": key,
            "label": key,
            "module_path": module_path,
            "gui_class": gui_class,
            "channel": MODULE_CHANNEL_MAP.get(key, ""),
            "start_method": MODULE_START_METHODS.get(key, "start_test"),
            "is_running": key in test_service.running_modules,
        })
    return {
        "modules": modules,
        "channel_groups": CHANNEL_GROUPS,
    }


@router.get("/tests/running")
async def list_running():
    """当前运行中的测试"""
    return {"running": test_service.running_modules}


@router.post("/tests/start")
async def start_test(request: dict):
    """启动测试模块

    Body: {"module": "Rin_FSV3004", "auto_start": true}
    """
    module_key = request.get("module", "")
    auto_start = request.get("auto_start", True)

    if module_key not in MODULE_REGISTRY:
        raise HTTPException(400, f"未知模块: {module_key}")

    from ..database import async_session

    async with async_session() as db:
        run_id = await test_service.start_test(module_key, auto_start=auto_start, db_session=db)
        if run_id is None:
            raise HTTPException(409, f"模块已在运行: {module_key}")
        await db.commit()

    return {"run_id": run_id, "module": module_key, "status": "started"}


@router.post("/tests/stop/{module_key}")
async def stop_test(module_key: str):
    """停止测试模块"""
    ok = await test_service.stop_test(module_key)
    if not ok:
        raise HTTPException(404, f"模块未在运行: {module_key}")
    return {"status": "stopped"}


@router.post("/tests/batch-start")
async def batch_start(request: dict):
    """批量启动测试

    Body: {"modules": ["Rin_FSV3004", "信噪比"], "auto_start": true}
    """
    modules = request.get("modules", [])
    auto_start = request.get("auto_start", True)
    results = []

    from ..database import async_session

    async with async_session() as db:
        for m in modules:
            if m not in MODULE_REGISTRY:
                results.append({"module": m, "error": "未知模块"})
                continue
            if m in test_service.running_modules:
                results.append({"module": m, "error": "已在运行"})
                continue
            run_id = await test_service.start_test(m, auto_start=auto_start, db_session=db)
            results.append({"module": m, "run_id": run_id, "status": "started"})
        await db.commit()

    return {"results": results}
