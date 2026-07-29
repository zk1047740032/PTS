"""仪器管理 API"""

import asyncio
import sys
from pathlib import Path

from fastapi import APIRouter

router = APIRouter(prefix="/api/instruments", tags=["instruments"])


def _get_instrument_ips() -> list[dict]:
    """从 CFG 读取仪器 IP 列表"""
    _PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
    if str(_PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(_PROJECT_ROOT))
    from core.config import CFG

    instruments = [
        {"name": "FSV3004 (RIN)", "ip": CFG.network.fsv3004_rin},
        {"name": "FSV3004 (线宽)", "ip": CFG.network.fsv3004_linewidth},
        {"name": "频谱仪 (单频)", "ip": CFG.network.single_freq_sa},
        {"name": "信号发生器 (线宽/PZT)", "ip": CFG.network.sig_gen_linewidth},
        {"name": "信号发生器 (时域)", "ip": CFG.network.sig_gen_timedomain},
        {"name": "示波器", "ip": CFG.network.scope},
        {"name": "OSA (信噪比)", "ip": CFG.network.osa},
    ]
    return instruments


async def _ping(ip: str, timeout: float = 1.0) -> bool:
    """异步 ping"""
    if not ip or ip == "N/A":
        return False
    try:
        proc = await asyncio.create_subprocess_exec(
            "ping", "-n", "1", "-w", str(int(timeout * 1000)), ip,
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.DEVNULL,
        )
        await proc.wait()
        return proc.returncode == 0
    except Exception:
        return False


@router.get("/status")
async def instrument_status():
    """获取所有仪器连接状态"""
    instruments = _get_instrument_ips()
    # 并发 ping
    tasks = [_ping(inst["ip"]) for inst in instruments]
    results = await asyncio.gather(*tasks)
    for inst, reachable in zip(instruments, results):
        inst["reachable"] = reachable
    return {"instruments": instruments}
