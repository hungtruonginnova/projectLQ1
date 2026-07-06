"""FastAPI backend for the weekly Overdue Parts Report."""

import json
import os
import re
import tempfile
from datetime import date
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response

from build_report import generate_report

APP_DIR = Path(__file__).parent
TEMPLATE = APP_DIR / "template" / "OVERDUE PARTS REPORT template PQ.xlsx"
CONFIG = APP_DIR / "config.yaml"

app = FastAPI(title="Overdue Parts Report API", version="1.0.0")


def _allowed_origins() -> list[str]:
    raw = os.environ.get("ALLOWED_ORIGINS", "*")
    return [o.strip() for o in raw.split(",") if o.strip()]


allowed = _allowed_origins()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"] if "*" in allowed else allowed,
    allow_origin_regex=r"https://.*\.vercel\.app"
    if any("*.vercel.app" in o for o in allowed)
    else None,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Report-Stats", "X-Report-Filename"],
)


def output_name(report_date: date) -> str:
    week = report_date.isocalendar()[1]
    return f"OVERDUE PARTS REPORT WK {week} ({report_date:%d.%m.%Y}) PQ.xlsx"


def parse_report_date(value: str | None) -> date:
    if not value:
        return date.today()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise HTTPException(status_code=422, detail="report_date must be YYYY-MM-DD")
    return date.fromisoformat(value)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/v1/report")
async def create_report(
    opd: UploadFile = File(...),
    invmaster: UploadFile = File(...),
    workorders: UploadFile = File(...),
    schedule: UploadFile = File(...),
    report_date: str | None = Form(default=None),
):
    if not TEMPLATE.exists():
        raise HTTPException(status_code=500, detail=f"Template not found: {TEMPLATE}")

    parsed_date = parse_report_date(report_date)
    filename = output_name(parsed_date)

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            paths = {
                "opd": tmp_path / "opd.xlsx",
                "invmaster": tmp_path / "invmaster.xlsx",
                "workorders": tmp_path / "workorders.xlsx",
                "schedule": tmp_path / "schedule.xlsx",
                "out": tmp_path / filename,
            }
            uploads = {
                "opd": opd,
                "invmaster": invmaster,
                "workorders": workorders,
                "schedule": schedule,
            }
            for key, upload in uploads.items():
                paths[key].write_bytes(await upload.read())

            stats = generate_report(
                paths["opd"],
                paths["invmaster"],
                paths["workorders"],
                paths["schedule"],
                TEMPLATE,
                paths["out"],
                report_date=parsed_date.isoformat(),
                config_path=CONFIG,
            )
            output_bytes = paths["out"].read_bytes()
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    stats_json = json.dumps(stats, default=str)
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Report-Filename": filename,
            "X-Report-Stats": stats_json,
        },
    )
