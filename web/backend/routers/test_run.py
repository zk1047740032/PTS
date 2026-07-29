"""测试运行 API 路由"""

import os
import pathlib

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from ..database import get_db
from ..models.test_result import TestRun, TestLog

router = APIRouter(prefix="/api/tests", tags=["test_runs"])


@router.get("/history")
async def list_history(
    module: str | None = None,
    limit: int = 20,
    offset: int = 0,
    db: AsyncSession = Depends(get_db),
):
    """历史测试列表（分页）"""
    q = select(TestRun).order_by(TestRun.created_at.desc())
    if module:
        q = q.where(TestRun.module_name == module)
    q = q.offset(offset).limit(limit)
    result = await db.execute(q)
    runs = result.scalars().all()
    return {
        "runs": [
            {
                "id": r.id,
                "module_name": r.module_name,
                "module_label": r.module_label,
                "channel": r.channel,
                "status": r.status,
                "start_time": r.start_time.isoformat() if r.start_time else None,
                "end_time": r.end_time.isoformat() if r.end_time else None,
                "duration_seconds": r.duration_seconds,
                "result_summary": r.result_summary,
                "output_dir": r.output_dir,
                "created_at": r.created_at.isoformat() if r.created_at else None,
            }
            for r in runs
        ],
        "total": len(runs),
    }


@router.get("/{run_id}")
async def get_test_detail(run_id: str, db: AsyncSession = Depends(get_db)):
    """测试详情"""
    result = await db.execute(select(TestRun).where(TestRun.id == run_id))
    run = result.scalar_one_or_none()
    if not run:
        raise HTTPException(404, "测试记录不存在")

    # 获取日志
    log_result = await db.execute(
        select(TestLog).where(TestLog.test_run_id == run_id).order_by(TestLog.sequence)
    )
    logs = log_result.scalars().all()

    return {
        "id": run.id,
        "module_name": run.module_name,
        "module_label": run.module_label,
        "channel": run.channel,
        "status": run.status,
        "start_time": run.start_time.isoformat() if run.start_time else None,
        "end_time": run.end_time.isoformat() if run.end_time else None,
        "duration_seconds": run.duration_seconds,
        "result_summary": run.result_summary,
        "output_dir": run.output_dir,
        "logs": [
            {
                "sequence": log.sequence,
                "timestamp": log.timestamp.isoformat() if log.timestamp else None,
                "level": log.level,
                "message": log.message,
                "data": log.data,
            }
            for log in logs
        ],
    }


@router.get("/{run_id}/files")
async def list_output_files(run_id: str, db: AsyncSession = Depends(get_db)):
    """列出测试输出文件"""
    result = await db.execute(select(TestRun).where(TestRun.id == run_id))
    run = result.scalar_one_or_none()
    if not run or not run.output_dir:
        return {"files": []}

    output_dir = pathlib.Path(run.output_dir)
    if not output_dir.exists():
        return {"files": []}

    files = []
    for f in output_dir.iterdir():
        if f.is_file():
            files.append({
                "name": f.name,
                "size": f.stat().st_size,
                "type": f.suffix.lower(),
                "url": f"/api/tests/{run_id}/files/{f.name}",
            })
    return {"files": files}


@router.get("/{run_id}/files/{filename}")
async def get_output_file(run_id: str, filename: str, db: AsyncSession = Depends(get_db)):
    """获取测试输出文件"""
    result = await db.execute(select(TestRun).where(TestRun.id == run_id))
    run = result.scalar_one_or_none()
    if not run or not run.output_dir:
        raise HTTPException(404, "测试记录不存在")

    file_path = pathlib.Path(run.output_dir) / filename
    if not file_path.exists():
        raise HTTPException(404, "文件不存在")

    # CSV 文件返回 JSON
    if file_path.suffix.lower() == ".csv":
        import csv
        rows = []
        with open(file_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append(row)
        return JSONResponse({"filename": filename, "rows": rows})

    return FileResponse(str(file_path))
