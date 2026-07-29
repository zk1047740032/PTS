"""测试结果数据库模型"""

import uuid
from datetime import datetime

from sqlalchemy import DateTime, Float, ForeignKey, Integer, String, Text, func
from sqlalchemy.dialects.sqlite import JSON
from sqlalchemy.orm import Mapped, mapped_column, relationship

from ..database import Base


def _new_uuid() -> str:
    return uuid.uuid4().hex[:12]


class TestRun(Base):
    """一次测试运行记录"""

    __tablename__ = "test_runs"

    id: Mapped[str] = mapped_column(String(12), primary_key=True, default=_new_uuid)
    module_name: Mapped[str] = mapped_column(String(64), index=True, comment="模块标识，如 Rin_FSV3004")
    module_label: Mapped[str] = mapped_column(String(64), comment="中文显示名，如 RIN")
    channel: Mapped[str] = mapped_column(String(32), nullable=True, comment="所属通道")
    status: Mapped[str] = mapped_column(
        String(16), default="pending", index=True,
        comment="pending | running | completed | error | stopped"
    )
    start_time: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    end_time: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    duration_seconds: Mapped[float] = mapped_column(Float, nullable=True)
    result_summary: Mapped[dict] = mapped_column(JSON, nullable=True, comment="结构化结果摘要")
    output_dir: Mapped[str] = mapped_column(String(512), nullable=True, comment="测试输出目录")
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())

    logs: Mapped[list["TestLog"]] = relationship(
        "TestLog", back_populates="test_run", cascade="all, delete-orphan",
        order_by="TestLog.sequence",
    )


class TestLog(Base):
    """测试日志条目"""

    __tablename__ = "test_logs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    test_run_id: Mapped[str] = mapped_column(String(12), ForeignKey("test_runs.id"), index=True)
    sequence: Mapped[int] = mapped_column(Integer, default=0)
    timestamp: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())
    level: Mapped[str] = mapped_column(String(16), default="info", comment="info | warning | error | data")
    message: Mapped[str] = mapped_column(Text)
    data: Mapped[dict] = mapped_column(JSON, nullable=True, comment="结构化数据载荷")

    test_run: Mapped["TestRun"] = relationship("TestRun", back_populates="logs")
