from __future__ import annotations

from datetime import date
from io import BytesIO
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse

from services.excel_reader import read_tasks_from_excel
from services.scheduler import build_schedule, serialize_tasks
from services.gantt_generator import create_gantt_workbook


app = FastAPI(title="Gantt Structured API", version="1.0.0")

# 개발용 CORS. 배포 시에는 실제 프론트엔드 주소만 남기는 것을 권장합니다.
app.add_middleware(
    CORSMiddleware,
allow_origins=["*"],
allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health")
def health_check():
    return {"ok": True}


@app.post("/api/preview")
async def preview_excel(
    file: UploadFile = File(...),
    project_start: str = Form("2026-01-01"),
    auto_schedule_mode: str = Form("smart"),
):
    """
    Excel을 읽고 백데이터 상태를 먼저 보여주는 API.
    간트차트를 만들기 전에 시작일/종료일/작업일 누락 여부를 확인합니다.
    """
    content = await file.read()

    try:
        raw_tasks = read_tasks_from_excel(BytesIO(content))
        scheduled = build_schedule(
            raw_tasks,
            project_start_date=date.fromisoformat(project_start),
            auto_schedule_mode=auto_schedule_mode,
        )
    except Exception as exc:
        return JSONResponse(status_code=400, content={"ok": False, "message": str(exc)})

    return {
        "ok": True,
        "fileName": file.filename,
        "count": len(scheduled),
        "tasks": serialize_tasks(scheduled),
        "summary": {
            "normal": sum(1 for t in scheduled if t.status == "정상"),
            "auto": sum(1 for t in scheduled if t.status == "자동계산"),
            "error": sum(1 for t in scheduled if t.status == "오류"),
        },
    }


@app.post("/api/generate")
async def generate_gantt(
    file: UploadFile = File(...),
    project_start: str = Form("2026-01-01"),
    auto_schedule_mode: str = Form("smart"),
):
    """
    Excel을 읽고 날짜 계산 후 간트차트 Excel을 생성합니다.
    """
    content = await file.read()

    try:
        raw_tasks = read_tasks_from_excel(BytesIO(content))
        scheduled = build_schedule(
            raw_tasks,
            project_start_date=date.fromisoformat(project_start),
            auto_schedule_mode=auto_schedule_mode,
        )
        errors = [t for t in scheduled if t.status == "오류"]
        if errors:
            return JSONResponse(
                status_code=400,
                content={
                    "ok": False,
                    "message": "작업일/시작일/종료일이 부족한 작업이 있습니다.",
                    "tasks": serialize_tasks(errors),
                },
            )

        wb = create_gantt_workbook(scheduled)
        output = BytesIO()
        wb.save(output)
        output.seek(0)

    except Exception as exc:
        return JSONResponse(status_code=400, content={"ok": False, "message": str(exc)})

    filename = "gantt_result.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"'
    }
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.post("/api/upload")
async def upload_compat(
    file: UploadFile = File(...),
    project_start: str = Form("2026-01-01"),
):
    """
    기존 프론트엔드가 /api/upload 하나만 호출하는 구조일 때를 위한 호환 API.
    업로드 즉시 간트차트 Excel을 반환합니다.
    """
    return await generate_gantt(
        file=file,
        project_start=project_start,
        auto_schedule_mode="smart",
    )
