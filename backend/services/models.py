from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass
class RawTask:
    row_no: int
    group: str
    task_name: str
    work_days: Optional[int] = None
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    progress: Optional[float] = None
    predecessor: Optional[str] = None
    floor: Optional[str] = None
    zone: Optional[str] = None


@dataclass
class ScheduledTask(RawTask):
    calculated_start: Optional[date] = None
    calculated_end: Optional[date] = None
    status: str = "정상"
    message: str = ""
    source: str = "input"
    warnings: list[str] = field(default_factory=list)
