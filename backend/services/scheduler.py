from __future__ import annotations

from dataclasses import asdict
from datetime import date, timedelta
from typing import Optional

from .models import RawTask, ScheduledTask


def add_days(start: date, days: int) -> date:
    """
    작업일을 달력일 기준으로 계산합니다.
    예: 1월 1일 + 10일 작업 = 1월 10일 종료
    """
    return start + timedelta(days=max(days, 1) - 1)


def days_between(start: date, end: date) -> int:
    return (end - start).days + 1


def normalize_work_days(task: RawTask) -> Optional[int]:
    """
    작업일 0은 오류가 아니라 마일스톤으로 인정합니다.
    화면/엑셀에서는 시작일과 종료일이 같은 1칸짜리 작업처럼 표시됩니다.
    """
    if task.work_days is not None and task.work_days >= 0:
        return task.work_days
    if task.start_date and task.end_date:
        return max(days_between(task.start_date, task.end_date), 1)
    return None


def clone_task(task: RawTask) -> ScheduledTask:
    return ScheduledTask(
        row_no=task.row_no,
        group=task.group,
        task_name=task.task_name,
        work_days=task.work_days,
        start_date=task.start_date,
        end_date=task.end_date,
        progress=task.progress,
        predecessor=task.predecessor,
        floor=task.floor,
        zone=task.zone,
    )


def build_schedule(
    raw_tasks: list[RawTask],
    project_start_date: date,
    auto_schedule_mode: str = "smart",
) -> list[ScheduledTask]:
    """
    간트차트 생성 전 날짜를 확정하는 단계.

    규칙:
    1. 시작일+종료일 있음: 그대로 사용
    2. 시작일+작업일 있음: 종료일 계산
    3. 종료일+작업일 있음: 시작일 계산
    4. 작업일만 있음: 자동 배치
    5. 작업일도 없고 날짜도 부족함: 오류
    """
    scheduled: list[ScheduledTask] = []
    by_task_name: dict[str, ScheduledTask] = {}
    last_global_start: Optional[date] = None
    last_global_end: Optional[date] = None
    last_group_end: dict[str, date] = {}

    for index, raw in enumerate(raw_tasks):
        task = clone_task(raw)
        work_days = normalize_work_days(raw)

        if raw.start_date and raw.end_date:
            task.calculated_start = raw.start_date
            task.calculated_end = raw.end_date
            task.work_days = work_days
            task.status = "정상"
            task.source = "input_date"

        elif raw.start_date and work_days is not None:
            task.calculated_start = raw.start_date
            task.calculated_end = add_days(raw.start_date, work_days)
            task.work_days = work_days
            task.status = "자동계산"
            task.source = "start_plus_duration"
            task.message = "시작일과 작업일을 기준으로 종료일을 계산했습니다."

        elif raw.end_date and work_days is not None:
            task.calculated_end = raw.end_date
            task.calculated_start = raw.end_date - timedelta(days=work_days - 1)
            task.work_days = work_days
            task.status = "자동계산"
            task.source = "end_minus_duration"
            task.message = "종료일과 작업일을 기준으로 시작일을 계산했습니다."

        elif work_days is not None:
            # 자동 배치
            start = None

            # 선행작업이 있으면 해당 작업 종료 다음 날
            if raw.predecessor and raw.predecessor in by_task_name:
                pred = by_task_name[raw.predecessor]
                if pred.calculated_end:
                    start = pred.calculated_end + timedelta(days=1)

            if start is None:
                if auto_schedule_mode == "sequential":
                    # 완전 순차 배치
                    start = project_start_date if last_global_end is None else last_global_end + timedelta(days=1)
                else:
                    # smart 모드:
                    # 같은 공종 안에서는 순차 배치
                    # 공종이 바뀌면 이전 전체 작업 시작점에서 일부 겹치게 배치
                    if raw.group in last_group_end:
                        start = last_group_end[raw.group] + timedelta(days=1)
                    elif last_global_start and last_global_end:
                        previous_duration = max(days_between(last_global_start, last_global_end), 1)
                        overlap_offset = max(int(previous_duration * 0.55), 1)
                        start = last_global_start + timedelta(days=overlap_offset)
                    else:
                        start = project_start_date

            task.calculated_start = start
            task.calculated_end = add_days(start, work_days)
            task.work_days = work_days
            task.status = "자동계산"
            task.source = "auto_scheduled"
            task.message = "시작일/종료일이 없어 자동 배치했습니다."

        else:
            task.status = "오류"
            task.source = "missing_data"
            task.message = "작업일 또는 시작일/종료일이 필요합니다."

        # 날짜 검증
        if task.calculated_start and task.calculated_end:
            if task.calculated_end < task.calculated_start:
                task.status = "오류"
                task.message = "종료일이 시작일보다 빠릅니다."
            else:
                actual_days = days_between(task.calculated_start, task.calculated_end)

                if task.work_days == 0:
                    task.warnings.append("작업일 0은 마일스톤으로 처리했습니다.")
                elif task.work_days is not None and abs(actual_days - task.work_days) >= 2:
                    task.warnings.append(
                        f"작업일({task.work_days})과 날짜 차이({actual_days})가 다릅니다."
                    )

                last_global_start = task.calculated_start
                last_global_end = task.calculated_end
                last_group_end[task.group] = task.calculated_end

        scheduled.append(task)
        by_task_name[task.task_name] = task

    return scheduled


def serialize_tasks(tasks: list[ScheduledTask]) -> list[dict]:
    result = []
    for task in tasks:
        result.append(
            {
                "rowNo": task.row_no,
                "group": task.group,
                "taskName": task.task_name,
                "workDays": task.work_days,
                "startDate": task.calculated_start.isoformat() if task.calculated_start else None,
                "endDate": task.calculated_end.isoformat() if task.calculated_end else None,
                "progress": task.progress,
                "predecessor": task.predecessor,
                "floor": task.floor,
                "zone": task.zone,
                "status": task.status,
                "message": task.message,
                "source": task.source,
                "warnings": task.warnings,
            }
        )
    return result
