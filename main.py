#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 채우기 전용 API 서버.
Make.com이 Iterator/Array Aggregator로 만든 [{"name":..,"content":..}, ...] 배열을
받아서 Template.pptx를 채운 .pptx 파일을 바이너리로 응답한다.
업로드/공유링크/메일 발송은 이 서버가 아니라 Make.com 기본 모듈이 담당한다.
"""
import os
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

from fill_template import build_pptx

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "Template.pptx"
SCRIPTS_DIR = BASE_DIR / "pptx_scripts"

# 한 요청에서 허용할 최대 인원 수 (비정상적으로 큰 요청으로 인한 과부하 방지)
MAX_RECORDS = 300

API_KEY = os.environ.get("PPTX_API_KEY")

app = FastAPI(title="kslife PPT fill API")


class Record(BaseModel):
    name: str = Field(min_length=1)
    content: str = Field(min_length=1)


class FillRequest(BaseModel):
    records: list[Record] = Field(min_length=1, max_length=MAX_RECORDS)


def _cleanup(path: Path) -> None:
    shutil.rmtree(path, ignore_errors=True)


@app.post("/api/fill-ppt")
def fill_ppt(body: FillRequest, x_api_key: str | None = Header(default=None)):
    if not API_KEY:
        raise HTTPException(status_code=500, detail="서버에 PPTX_API_KEY가 설정되어 있지 않습니다.")
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="유효하지 않은 API 키입니다.")

    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="Template.pptx를 찾을 수 없습니다.")

    work_dir = Path(tempfile.mkdtemp(prefix="pptx_req_"))
    output_path = work_dir / f"{uuid.uuid4().hex}.pptx"

    try:
        records = [r.model_dump() for r in body.records]
        build_pptx(
            template_path=str(TEMPLATE_PATH),
            records=records,
            output_path=str(output_path),
            scripts_dir=str(SCRIPTS_DIR),
            work_dir=str(work_dir),
        )
    except Exception as e:
        _cleanup(work_dir)
        raise HTTPException(status_code=500, detail=f"PPT 생성 실패: {e}")

    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="평가서.pptx",
        background=BackgroundTask(_cleanup, work_dir),
    )


@app.get("/api/health")
def health():
    return {"status": "ok"}
