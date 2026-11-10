#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
교육 평가서 자동화 API 서버.

- POST /api/webhook        : 프론트엔드(Cloudflare Worker 경유)가 호출. 참석자 메모를 받아
                              즉시 202로 응답하고, 백그라운드에서 Claude 호출 → PPT 생성 →
                              다운로드 페이지 링크를 Telegram으로 전송.
- POST /api/fill-ppt       : 이미 만들어진 {name, content} 배열을 받아 PPT 바이너리로 응답
                              (수동 테스트/디버깅용, Claude 호출 없음).
- GET  /files              : 생성된 PPT 목록 페이지 (비밀번호 로그인 필요). 날짜별로 선택해
                              다운로드하거나 삭제할 수 있음.
- GET  /files/login        : 비밀번호 입력 페이지. 맞으면 서명 쿠키를 심고 /files 로 이동.
- GET  /files/{id}         : 개별 PPT 다운로드 (로그인 쿠키 필요).
- POST /files/{id}/delete  : 개별 PPT 삭제 (로그인 쿠키 필요).
- GET  /latest.pptx        : 가장 최근에 생성된 PPT (고정 URL, 인증 없음). PC 북마크용.
"""
import hmac
import html
import json
import os
import re
import shutil
import tempfile
import time
import uuid
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from urllib.parse import parse_qs
from zoneinfo import ZoneInfo

from fastapi import BackgroundTasks, Depends, FastAPI, Header, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
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
LATEST_PPTX = DOWNLOADS_DIR / "latest.pptx"
LATEST_META = DOWNLOADS_DIR / "latest.json"

PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
KST = ZoneInfo("Asia/Seoul")

# 한 요청에서 허용할 최대 인원 수 (비정상적으로 큰 요청으로 인한 과부하 방지)
MAX_RECORDS = 300
# 생성된 파일 보존 기간(일). 목록 페이지에서 과거 파일을 다시 찾을 수 있어야 하므로 기본 90일.
RETENTION_DAYS = int(os.environ.get("RETENTION_DAYS", "90"))

API_KEY = os.environ.get("PPTX_API_KEY")
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "https://kslife-pptx.duckdns.org")
# 목록 페이지 비밀번호 (아이디 없음). 미설정 시 /files 계열은 500을 반환한다.
FILES_PASSWORD = os.environ.get("FILES_PASSWORD", "")
SESSION_COOKIE = "kslife_files"
SESSION_MAX_AGE = 30 * 86400  # 30일

FILE_ID_RE = re.compile(r"^[0-9a-f]{32}$")

app = FastAPI(title="kslife PPT fill API")


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


# --------------------------------------------------------------------------- 공통 유틸


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


def _meta_path(file_id: str) -> Path:
    return DOWNLOADS_DIR / f"{file_id}.json"


def _pptx_path(file_id: str) -> Path:
    return DOWNLOADS_DIR / f"{file_id}.pptx"


def _read_meta(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _remove_file(file_id: str) -> None:
    _pptx_path(file_id).unlink(missing_ok=True)
    _meta_path(file_id).unlink(missing_ok=True)


def _cleanup_old_downloads() -> None:
    cutoff = time.time() - RETENTION_DAYS * 86400
    for f in DOWNLOADS_DIR.glob("*.pptx"):
        if f.name != LATEST_PPTX.name and f.stat().st_mtime < cutoff:
            _remove_file(f.stem)


def _list_files() -> list[dict]:
    """다운로드 폴더의 PPT 목록을 최신순으로 반환한다. 메타데이터가 없는 옛 파일은 수정 시각으로 대체."""
    items = []
    for p in DOWNLOADS_DIR.glob("*.pptx"):
        if p.name == LATEST_PPTX.name or not FILE_ID_RE.match(p.stem):
            continue
        meta = _read_meta(_meta_path(p.stem))
        created_at = meta.get("created_at")
        if created_at:
            created = datetime.fromisoformat(created_at)
        else:
            created = datetime.fromtimestamp(p.stat().st_mtime, KST)
        items.append({"file_id": p.stem, "created": created, "date": meta.get("date", "")})
    items.sort(key=lambda i: i["created"], reverse=True)
    return items


def _refresh_latest(deleted_id: str) -> None:
    """삭제된 파일이 latest.pptx의 원본이었다면 남은 파일 중 최신으로 교체하거나 제거한다."""
    if _read_meta(LATEST_META).get("file_id") != deleted_id:
        return
    remaining = _list_files()
    if remaining:
        newest = remaining[0]["file_id"]
        shutil.copyfile(_pptx_path(newest), LATEST_PPTX)
        _write_json(LATEST_META, {"file_id": newest})
    else:
        LATEST_PPTX.unlink(missing_ok=True)
        LATEST_META.unlink(missing_ok=True)


# --------------------------------------------------------------------------- 로그인 (비밀번호 전용)


def _session_token() -> str:
    """비밀번호에서 파생한 서명값. 서버에 세션 저장소를 두지 않아도 되고, 비밀번호를 바꾸면 모두 무효화된다."""
    return hmac.new(FILES_PASSWORD.encode("utf-8"), b"kslife-files-session-v1", sha256).hexdigest()


def _is_logged_in(request: Request) -> bool:
    if not FILES_PASSWORD:
        return False
    cookie = request.cookies.get(SESSION_COOKIE, "")
    return hmac.compare_digest(cookie, _session_token())


def _require_login(request: Request) -> None:
    if not FILES_PASSWORD:
        raise HTTPException(status_code=500, detail="서버에 FILES_PASSWORD가 설정되어 있지 않습니다.")
    if not _is_logged_in(request):
        # HTML 페이지 요청은 로그인 화면으로 보내고, 그 외는 401
        raise HTTPException(
            status_code=303 if request.method == "GET" else 401,
            headers={"Location": "/files/login"} if request.method == "GET" else None,
            detail="로그인이 필요합니다.",
        )


def _set_session_cookie(response: RedirectResponse) -> None:
    response.set_cookie(
        key=SESSION_COOKIE,
        value=_session_token(),
        max_age=SESSION_MAX_AGE,
        httponly=True,
        secure=PUBLIC_BASE_URL.startswith("https://"),
        samesite="lax",
        path="/",
    )


# --------------------------------------------------------------------------- HTML


_PAGE_CSS = """
*{box-sizing:border-box}
body{margin:0;padding:16px;font-family:-apple-system,"Malgun Gothic",sans-serif;background:#f5f6f8;color:#222}
h1{font-size:1.25rem;margin:0 0 16px}
.card{background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.08);max-width:640px;margin:0 auto}
table{width:100%;border-collapse:collapse}
td{padding:12px 4px;border-bottom:1px solid #eee;vertical-align:middle}
td.actions{text-align:right;white-space:nowrap}
a.btn,button{display:inline-block;padding:8px 14px;border-radius:6px;border:0;font-size:.9rem;cursor:pointer;text-decoration:none}
a.btn{background:#2563eb;color:#fff}
button.del{background:#fff;color:#b91c1c;border:1px solid #fca5a5;margin-left:6px}
form.inline{display:inline}
input[type=password]{width:100%;padding:12px;font-size:1rem;border:1px solid #ccc;border-radius:6px;margin-bottom:12px}
button.primary{background:#2563eb;color:#fff;width:100%;padding:12px;font-size:1rem}
.err{color:#b91c1c;margin-bottom:12px}
.empty{color:#666;text-align:center;padding:24px 0}
.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
.top a{font-size:.85rem;color:#666}
.badge{font-size:.75rem;color:#2563eb;margin-left:6px}
"""


def _page(title: str, body: str) -> HTMLResponse:
    return HTMLResponse(
        "<!doctype html><html lang='ko'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        f"<title>{html.escape(title)}</title><style>{_PAGE_CSS}</style></head>"
        f"<body><div class='card'>{body}</div></body></html>"
    )


def _login_page(error: str = "") -> HTMLResponse:
    err = f"<div class='err'>{html.escape(error)}</div>" if error else ""
    return _page(
        "평가서 다운로드 - 로그인",
        "<h1>평가서 다운로드</h1>"
        f"{err}"
        "<form method='post' action='/files/login'>"
        "<input type='password' name='password' placeholder='비밀번호' autofocus required>"
        "<button class='primary' type='submit'>로그인</button>"
        "</form>",
    )


def _files_page(items: list[dict]) -> HTMLResponse:
    latest_id = _read_meta(LATEST_META).get("file_id")
    if not items:
        rows = "<div class='empty'>생성된 파일이 없습니다.</div>"
    else:
        trs = []
        for i in items:
            label = i["created"].astimezone(KST).strftime("%Y-%m-%d %H:%M")
            badge = "<span class='badge'>최신</span>" if i["file_id"] == latest_id else ""
            trs.append(
                f"<tr><td>{label}{badge}</td>"
                f"<td class='actions'><a class='btn' href='/files/{i['file_id']}'>다운로드</a>"
                f"<form class='inline' method='post' action='/files/{i['file_id']}/delete' "
                f"onsubmit=\"return confirm('{label} 파일을 삭제할까요?')\">"
                f"<button class='del' type='submit'>삭제</button></form></td></tr>"
            )
        rows = f"<table>{''.join(trs)}</table>"
    return _page(
        "평가서 다운로드",
        "<div class='top'><h1 style='margin:0'>평가서 다운로드</h1>"
        "<a href='/files/logout'>로그아웃</a></div>"
        f"{rows}",
    )


# --------------------------------------------------------------------------- 웹훅 처리


def _process_webhook(date: str, attendees: list[dict]) -> None:
    try:
        _cleanup_old_downloads()

        records = []
        for a in attendees:
            content = generate_evaluation(a["name"], a["notes"])
            records.append({"name": a["name"], "content": content})

        file_id = uuid.uuid4().hex
        output_path = _pptx_path(file_id)
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

        _write_json(
            _meta_path(file_id),
            {
                "file_id": file_id,
                "date": date,
                "count": len(records),
                "names": [r["name"] for r in records],
                "created_at": datetime.now(KST).isoformat(timespec="seconds"),
            },
        )
        shutil.copyfile(output_path, LATEST_PPTX)
        _write_json(LATEST_META, {"file_id": file_id})

        send_telegram_message(
            f"[{date}] 평가서 생성 완료 ({len(records)}명)\n"
            f"최신 파일: {PUBLIC_BASE_URL}/latest.pptx\n"
            f"전체 목록: {PUBLIC_BASE_URL}/files"
        )
    except Exception as e:
        print(f"[webhook 처리 실패] {e}")
        send_telegram_error(f"[{date}] 평가서 생성 실패: {e}")


# --------------------------------------------------------------------------- 라우트


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
        media_type=PPTX_MEDIA_TYPE,
        filename="평가서.pptx",
        background=BackgroundTask(_cleanup_dir, work_dir),
    )


@app.get("/latest.pptx")
def latest():
    if not LATEST_PPTX.exists():
        raise HTTPException(status_code=404, detail="아직 생성된 파일이 없습니다.")
    return FileResponse(path=LATEST_PPTX, media_type=PPTX_MEDIA_TYPE, filename="평가서_최신.pptx")


@app.get("/files/login")
def files_login_page(request: Request):
    if not FILES_PASSWORD:
        raise HTTPException(status_code=500, detail="서버에 FILES_PASSWORD가 설정되어 있지 않습니다.")
    if _is_logged_in(request):
        return RedirectResponse("/files", status_code=303)
    return _login_page()


@app.post("/files/login")
async def files_login(request: Request):
    if not FILES_PASSWORD:
        raise HTTPException(status_code=500, detail="서버에 FILES_PASSWORD가 설정되어 있지 않습니다.")
    # python-multipart 의존성을 피하기 위해 폼 본문을 직접 파싱한다.
    body = (await request.body()).decode("utf-8", errors="replace")
    password = parse_qs(body).get("password", [""])[0]
    if not hmac.compare_digest(password.encode("utf-8"), FILES_PASSWORD.encode("utf-8")):
        time.sleep(0.5)  # 무차별 대입 속도 제한
        return _login_page("비밀번호가 맞지 않습니다.")
    resp = RedirectResponse("/files", status_code=303)
    _set_session_cookie(resp)
    return resp


@app.get("/files/logout")
def files_logout():
    resp = RedirectResponse("/files/login", status_code=303)
    resp.delete_cookie(SESSION_COOKIE, path="/")
    return resp


@app.get("/files", dependencies=[Depends(_require_login)])
def files_page():
    return _files_page(_list_files())


@app.get("/files/{file_id}", dependencies=[Depends(_require_login)])
def files_download(file_id: str):
    if not FILE_ID_RE.match(file_id):
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    path = _pptx_path(file_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    meta = _read_meta(_meta_path(file_id))
    created_at = meta.get("created_at")
    created = datetime.fromisoformat(created_at) if created_at else datetime.fromtimestamp(path.stat().st_mtime, KST)
    filename = f"평가서_{created.astimezone(KST).strftime('%Y%m%d_%H%M')}.pptx"
    return FileResponse(path=path, media_type=PPTX_MEDIA_TYPE, filename=filename)


@app.post("/files/{file_id}/delete", dependencies=[Depends(_require_login)])
def files_delete(file_id: str):
    if not FILE_ID_RE.match(file_id) or not _pptx_path(file_id).exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    _remove_file(file_id)
    _refresh_latest(file_id)
    return RedirectResponse("/files", status_code=303)


@app.get("/api/health")
def health():
    return {"status": "ok"}
