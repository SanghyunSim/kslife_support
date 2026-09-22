#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
교육 평가서 자동화 API 서버.

- POST /api/webhook   : 프론트엔드(Cloudflare Worker 경유)가 호출. 참석자 메모를 받아
                         즉시 202로 응답하고, 백그라운드에서 Claude 호출 → PPT 생성 →
                         다운로드 링크를 Telegram으로 전송.
- POST /api/fill-ppt  : 이미 만들어진 {name, content} 배열을 받아 PPT 바이너리로 응답
                         (수동 테스트/디버깅용, Claude 호출 없음).
- GET  /downloads/{f} : 생성된 PPT 파일 다운로드 (추측 불가능한 UUID 파일명 + 일정 기간
                         후 자동 삭제로 보호. 클릭 한 번으로 받아야 하므로 API 키 불필요).
- GET  /latest.pptx   : 가장 최근에 생성된 PPT (고정 URL). PC에서 이 주소를 북마크해두면
                         매번 Telegram 링크를 열 필요 없이 항상 최신 결과를 받을 수 있음.
"""
import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, Header, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

from claude_client import generate_evaluation
from fill_template import build_pptx
from telegram_client import send_telegram_error, send_telegram_message

BASE_DIR = Path(__file__).parent
# 인원 수에 따라 템플릿을 고른다. Template2는 한 장에 표 3개(9명), Template은 표 4개(12명).
TEMPLATE_SMALL = BASE_DIR / "Template2.pptx"
TEMPLATE_LARGE = BASE_DIR / "Template.pptx"
SMALL_TEMPLATE_MAX = 9
SCRIPTS_DIR = BASE_DIR / "pptx_scripts"
DOWNLOADS_DIR = BASE_DIR / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)

# 한 요청에서 허용할 최대 인원 수 (비정상적으로 큰 요청으로 인한 과부하 방지)
MAX_RECORDS = 300
RETENTION_DAYS = 7

API_KEY = os.environ.get("PPTX_API_KEY")
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "https://kslife-pptx.duckdns.org")

app = FastAPI(title="kslife PPT fill API")
app.mount("/downloads", StaticFiles(directory=DOWNLOADS_DIR), name="downloads")


class Record(BaseModel):
    name: str = Field(min_length=1)
    content: str = Field(min_length=1)


class FillRequest(BaseModel):
    records: list[Record] = Field(min_length=1, max_length=MAX_RECORDS)


class Attendee(BaseModel):
    name: str = Field(min_length=1)
    notes: str = Field(min_length=1)


class WebhookRequest(BaseModel):
    date: str = Field(min_length=1)
    attendees: list[Attendee] = Field(min_length=1, max_length=MAX_RECORDS)


def _require_api_key(x_api_key: str | None) -> None:
    if not API_KEY:
        raise HTTPException(status_code=500, detail="서버에 PPTX_API_KEY가 설정되어 있지 않습니다.")
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="유효하지 않은 API 키입니다.")


def _pick_template(n_records: int) -> Path:
    """9명까지는 Template2.pptx, 10명부터는 Template.pptx."""
    return TEMPLATE_SMALL if n_records <= SMALL_TEMPLATE_MAX else TEMPLATE_LARGE


def _cleanup_dir(path: Path) -> None:
    shutil.rmtree(path, ignore_errors=True)


def _cleanup_old_downloads() -> None:
    cutoff = time.time() - RETENTION_DAYS * 86400
    for f in DOWNLOADS_DIR.glob("*.pptx"):
        if f.name != "latest.pptx" and f.stat().st_mtime < cutoff:
            f.unlink(missing_ok=True)


def _process_webhook(date: str, attendees: list[dict]) -> None:
    try:
        _cleanup_old_downloads()

        records = []
        for a in attendees:
            content = generate_evaluation(a["name"], a["notes"])
            records.append({"name": a["name"], "content": content})

        file_id = uuid.uuid4().hex
        output_path = DOWNLOADS_DIR / f"{file_id}.pptx"
        work_dir = Path(tempfile.mkdtemp(prefix="pptx_wh_"))
        try:
            build_pptx(
                template_path=str(_pick_template(len(records))),
                records=records,
                output_path=str(output_path),
                scripts_dir=str(SCRIPTS_DIR),
                work_dir=str(work_dir),
            )
        finally:
            _cleanup_dir(work_dir)

        shutil.copyfile(output_path, DOWNLOADS_DIR / "latest.pptx")

        send_telegram_message(
            f"[{date}] 평가서 생성 완료 ({len(records)}명)\n"
            f"PC에서 확인: {PUBLIC_BASE_URL}/latest.pptx"
        )
    except Exception as e:
        print(f"[webhook 처리 실패] {e}")
        send_telegram_error(f"[{date}] 평가서 생성 실패: {e}")


@app.post("/api/webhook", status_code=202)
def webhook(body: WebhookRequest, background_tasks: BackgroundTasks, x_api_key: str | None = Header(default=None)):
    _require_api_key(x_api_key)
    attendees = [a.model_dump() for a in body.attendees]
    background_tasks.add_task(_process_webhook, body.date, attendees)
    return {"status": "accepted"}


@app.post("/api/fill-ppt")
def fill_ppt(body: FillRequest, x_api_key: str | None = Header(default=None)):
    _require_api_key(x_api_key)

    template_path = _pick_template(len(body.records))
    if not template_path.exists():
        raise HTTPException(status_code=500, detail=f"{template_path.name}을 찾을 수 없습니다.")

    work_dir = Path(tempfile.mkdtemp(prefix="pptx_req_"))
    output_path = work_dir / f"{uuid.uuid4().hex}.pptx"

    try:
        records = [r.model_dump() for r in body.records]
        build_pptx(
            template_path=str(template_path),
            records=records,
            output_path=str(output_path),
            scripts_dir=str(SCRIPTS_DIR),
            work_dir=str(work_dir),
        )
    except Exception as e:
        _cleanup_dir(work_dir)
        raise HTTPException(status_code=500, detail=f"PPT 생성 실패: {e}")

    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="평가서.pptx",
        background=BackgroundTask(_cleanup_dir, work_dir),
    )


@app.get("/latest.pptx")
def latest():
    path = DOWNLOADS_DIR / "latest.pptx"
    if not path.exists():
        raise HTTPException(status_code=404, detail="아직 생성된 파일이 없습니다.")
    return FileResponse(
        path=path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="평가서_최신.pptx",
    )


@app.get("/api/health")
def health():
    return {"status": "ok"}
