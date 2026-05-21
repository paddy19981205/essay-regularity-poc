#!/usr/bin/env python3
"""Local standard-library server for the upload -> result -> history POC."""

from __future__ import annotations

import json
import mimetypes
import hmac
import os
import secrets
import shutil
import sys
import threading
import traceback
import uuid
import zipfile
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse


BACKEND_DIR = Path(__file__).resolve().parent
POC_ROOT = BACKEND_DIR.parent
REPO_ROOT = POC_ROOT.parent
FRONTEND_DIR = POC_ROOT / "frontend"
DIST_DIR = FRONTEND_DIR / "dist"
RUNS_DIR = Path(os.environ.get("POC_RUNS_DIR", BACKEND_DIR / "server_data" / "runs"))
INDEX_PATH = RUNS_DIR / "index.json"
MAX_UPLOAD_BYTES = 2 * 1024 * 1024 * 1024
UPLOAD_CHUNK_SIZE = 4 * 1024 * 1024
AUTH_USERNAME = os.environ.get("POC_USERNAME", "admin")
AUTH_PASSWORD = os.environ.get("POC_PASSWORD", "essay2026")
SESSION_COOKIE = "essay_poc_session"
SESSION_MAX_AGE = 60 * 60 * 12
COOKIE_SECURE = os.environ.get("POC_COOKIE_SECURE", "").lower() in {"1", "true", "yes"}
ALLOWED_FILE_TYPES = {
    "docx": ("manual.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
}

sys.path.insert(0, str(REPO_ROOT))
from tools.build_essay_manual import run_pipeline  # noqa: E402

INDEX_LOCK = threading.Lock()
ANALYSIS_SEMAPHORE = threading.Semaphore(1)
SESSION_LOCK = threading.Lock()
SESSIONS: dict[str, dict] = {}


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def make_session(username: str) -> str:
    token = secrets.token_urlsafe(32)
    with SESSION_LOCK:
        SESSIONS[token] = {"username": username, "createdAt": now_iso()}
    return token


def drop_session(token: str) -> None:
    if not token:
        return
    with SESSION_LOCK:
        SESSIONS.pop(token, None)


def get_cookie_value(cookie_header: str, name: str) -> str:
    for part in (cookie_header or "").split(";"):
        if "=" not in part:
            continue
        key, value = part.strip().split("=", 1)
        if key == name:
            return value
    return ""


def build_session_cookie(token: str) -> str:
    pieces = [
        f"{SESSION_COOKIE}={token}",
        "Path=/",
        f"Max-Age={SESSION_MAX_AGE}",
        "HttpOnly",
        "SameSite=Lax",
    ]
    if COOKIE_SECURE:
        pieces.append("Secure")
    return "; ".join(pieces)


def clear_session_cookie() -> str:
    return f"{SESSION_COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"


def constant_time_text_equals(left: str, right: str) -> bool:
    return hmac.compare_digest(left.encode("utf-8"), right.encode("utf-8"))


def make_progress(stage: str, current: int = 0, total: int = 0, message: str = "", current_file: str = "") -> dict:
    percent = round((current / total) * 100) if total else 0
    return {
        "stage": stage,
        "current": current,
        "total": total,
        "percent": max(0, min(100, percent)),
        "message": message,
        "currentFile": current_file,
        "updatedAt": now_iso(),
    }


def read_manual_preview(run_id: str) -> dict:
    units_path = RUNS_DIR / run_id / "output" / "manual_units.json"
    if not units_path.exists():
        return {"units": []}
    try:
        raw_units = json.loads(units_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"units": []}

    units = []
    if not isinstance(raw_units, list):
        return {"units": units}
    for unit in raw_units:
        if not isinstance(unit, dict):
            continue
        entries = unit.get("entries") if isinstance(unit.get("entries"), list) else []
        sample_titles = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            title = str(entry.get("title") or "").strip()
            if title:
                sample_titles.append(title[:80])
            if len(sample_titles) >= 3:
                break
        title = str(unit.get("title") or "").strip()
        if title or entries:
            units.append(
                {
                    "title": title or "未命名單元",
                    "entryCount": len(entries),
                    "sampleTitles": sample_titles,
                }
            )
    return {"units": units}


def with_manual_preview(record: dict) -> dict:
    if not isinstance(record, dict):
        return record
    if record.get("status") != "completed":
        return record
    enriched = dict(record)
    enriched["manualPreview"] = read_manual_preview(str(record.get("id") or ""))
    return enriched


def load_index() -> list[dict]:
    if not INDEX_PATH.exists():
        return []
    try:
        return json.loads(INDEX_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []


def save_index(runs: list[dict]) -> None:
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    tmp = INDEX_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(runs, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(INDEX_PATH)


def upsert_run(record: dict) -> None:
    with INDEX_LOCK:
        runs = [run for run in load_index() if run.get("id") != record.get("id")]
        runs.insert(0, record)
        save_index(runs)


def get_run(run_id: str) -> dict | None:
    with INDEX_LOCK:
        for run in load_index():
            if run.get("id") == run_id:
                return run
    return None


def recover_interrupted_runs() -> None:
    with INDEX_LOCK:
        runs = load_index()
        changed = False
        for run in runs:
            if run.get("status") in {"queued", "running"}:
                run.update(
                    {
                        "status": "failed",
                        "completedAt": run.get("completedAt") or now_iso(),
                        "error": "伺服器重新啟動前分析未完成，請重新上傳此批次。",
                        "progress": make_progress(
                            "failed",
                            current=(run.get("progress") or {}).get("current", 0),
                            total=(run.get("progress") or {}).get("total", 0),
                            message="伺服器重新啟動前分析未完成，請重新上傳此批次。",
                            current_file=(run.get("progress") or {}).get("currentFile", ""),
                        ),
                    }
                )
                result_path = RUNS_DIR / str(run.get("id")) / "result.json"
                if result_path.exists():
                    result_path.write_text(json.dumps(run, ensure_ascii=False, indent=2), encoding="utf-8")
                changed = True
        if changed:
            save_index(runs)


def sanitize_batch_name(name: str) -> str:
    name = unquote(name or "").strip()
    return name[:80] or "未命名批次"


def save_run_record(record: dict, record_path: Path) -> None:
    record_path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
    upsert_run(record)


def update_progress(record: dict, record_path: Path, payload: dict) -> None:
    progress = make_progress(
        payload.get("stage", "running"),
        current=int(payload.get("current") or 0),
        total=int(payload.get("total") or 0),
        message=payload.get("message") or "",
        current_file=payload.get("currentFile") or payload.get("current_file") or "",
    )
    record["progress"] = progress
    save_run_record(record, record_path)


def count_pdfs_in_zip(zip_path: Path) -> int:
    try:
        with zipfile.ZipFile(zip_path) as zf:
            return sum(
                1
                for member in zf.infolist()
                if not member.is_dir() and member.filename.lower().endswith(".pdf")
            )
    except zipfile.BadZipFile as exc:
        raise RuntimeError("ZIP 檔案無法讀取，請確認壓縮檔完整。") from exc


def write_request_body_to_file(source, target: Path, length: int) -> None:
    remaining = length
    with target.open("wb") as dst:
        while remaining > 0:
            chunk = source.read(min(UPLOAD_CHUNK_SIZE, remaining))
            if not chunk:
                raise RuntimeError("Upload interrupted before all bytes were received")
            dst.write(chunk)
            remaining -= len(chunk)


def safe_extract_zip(zip_path: Path, target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        for member in zf.infolist():
            dest = (target_dir / member.filename).resolve()
            if not str(dest).startswith(str(target_dir.resolve())):
                raise RuntimeError(f"Unsafe ZIP path: {member.filename}")
            if member.is_dir():
                dest.mkdir(parents=True, exist_ok=True)
            else:
                dest.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(member) as src, dest.open("wb") as dst:
                    shutil.copyfileobj(src, dst)


def run_analysis(run_id: str) -> None:
    run_dir = RUNS_DIR / run_id
    record_path = run_dir / "result.json"
    record = json.loads(record_path.read_text(encoding="utf-8"))
    try:
        update_progress(record, record_path, {"stage": "queued", "message": "等待分析工作開始"})
        with ANALYSIS_SEMAPHORE:
            record["status"] = "running"
            record["startedAt"] = now_iso()
            update_progress(record, record_path, {"stage": "validating_zip", "message": "正在檢查 ZIP 內容"})

            extracted_dir = run_dir / "extracted"
            output_dir = run_dir / "output"
            pdf_count = count_pdfs_in_zip(run_dir / "source.zip")
            if pdf_count == 0:
                raise RuntimeError("ZIP 內沒有 PDF 檔案。")

            update_progress(
                record,
                record_path,
                {
                    "stage": "extracting",
                    "current": 0,
                    "total": pdf_count,
                    "message": f"正在解壓 ZIP，找到 {pdf_count} 份 PDF",
                },
            )
            safe_extract_zip(run_dir / "source.zip", extracted_dir)

            summary = run_pipeline(
                extracted_dir,
                output_dir,
                record["batchName"],
                progress_callback=lambda payload: update_progress(record, record_path, payload),
            )
            record.update(
                {
                    "status": "completed",
                    "completedAt": now_iso(),
                    "summary": summary,
                    "manualPreview": read_manual_preview(run_id),
                    "error": None,
                    "progress": make_progress("completed", pdf_count, pdf_count, "分析完成"),
                }
            )
    except Exception as exc:  # POC server should preserve failure history.
        previous = record.get("progress") or {}
        record.update(
            {
                "status": "failed",
                "completedAt": now_iso(),
                "error": str(exc),
                "traceback": traceback.format_exc(),
                "progress": make_progress(
                    "failed",
                    int(previous.get("current") or 0),
                    int(previous.get("total") or 0),
                    str(exc),
                    previous.get("currentFile") or "",
                ),
            }
        )
    finally:
        save_run_record(record, record_path)


class Handler(BaseHTTPRequestHandler):
    server_version = "EssayRegularityPOC/0.1"

    def log_message(self, fmt: str, *args) -> None:
        print(f"[{now_iso()}] {self.address_string()} {fmt % args}")

    def send_json(self, payload: dict | list, status: int = 200, headers: dict[str, str] | None = None) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,DELETE,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type,X-Batch-Name,X-File-Name")
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self) -> None:
        self.send_json({})

    def current_session(self) -> dict | None:
        token = get_cookie_value(self.headers.get("Cookie", ""), SESSION_COOKIE)
        if not token:
            return None
        with SESSION_LOCK:
            return SESSIONS.get(token)

    def is_authenticated(self) -> bool:
        return self.current_session() is not None

    def require_auth(self) -> bool:
        if self.is_authenticated():
            return True
        self.send_json({"error": "請先登入。"}, HTTPStatus.UNAUTHORIZED)
        return False

    def read_json_body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        try:
            return json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return {}

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/api/session":
            session = self.current_session()
            self.send_json(
                {
                    "authenticated": bool(session),
                    "username": session.get("username") if session else None,
                }
            )
            return
        if path == "/api/runs":
            if not self.require_auth():
                return
            self.send_json({"runs": [with_manual_preview(record) for record in load_index()]})
            return
        if path.startswith("/api/runs/"):
            if not self.require_auth():
                return
            parts = [part for part in path.split("/") if part]
            if len(parts) == 3:
                record = get_run(parts[2])
                if not record:
                    self.send_json({"error": "Run not found"}, HTTPStatus.NOT_FOUND)
                    return
                self.send_json(with_manual_preview(record))
                return
            if len(parts) == 5 and parts[3] == "files":
                self.send_run_file(parts[2], parts[4])
                return
        self.send_static(path)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/session":
            payload = self.read_json_body()
            username = str(payload.get("username", ""))
            password = str(payload.get("password", ""))
            if constant_time_text_equals(username, AUTH_USERNAME) and constant_time_text_equals(password, AUTH_PASSWORD):
                token = make_session(username)
                self.send_json(
                    {"authenticated": True, "username": username},
                    headers={"Set-Cookie": build_session_cookie(token)},
                )
                return
            self.send_json({"error": "帳號或密碼錯誤。"}, HTTPStatus.UNAUTHORIZED)
            return

        if parsed.path != "/api/runs":
            self.send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)
            return
        if not self.require_auth():
            return

        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            self.send_json({"error": "No upload body received"}, HTTPStatus.BAD_REQUEST)
            return
        if length > MAX_UPLOAD_BYTES:
            self.send_json({"error": "上傳檔案太大，請改用較小批次或聯絡管理者。"}, HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return

        query = parse_qs(parsed.query)
        batch_name = sanitize_batch_name(
            self.headers.get("X-Batch-Name")
            or query.get("batchName", [""])[0]
            or self.headers.get("X-File-Name")
            or "未命名批次"
        )
        file_name = self.headers.get("X-File-Name", "upload.zip")
        if not file_name.lower().endswith(".zip"):
            self.send_json({"error": "Only .zip uploads are supported"}, HTTPStatus.BAD_REQUEST)
            return

        run_id = datetime.now().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:8]
        run_dir = RUNS_DIR / run_id
        run_dir.mkdir(parents=True, exist_ok=True)
        zip_path = run_dir / "source.zip"
        try:
            write_request_body_to_file(self.rfile, zip_path, length)
        except RuntimeError as exc:
            shutil.rmtree(run_dir, ignore_errors=True)
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return

        record = {
            "id": run_id,
            "batchName": batch_name,
            "sourceFile": file_name,
            "status": "queued",
            "createdAt": now_iso(),
            "startedAt": None,
            "completedAt": None,
            "summary": None,
            "error": None,
            "progress": make_progress("queued", message="已收到 ZIP，等待背景分析"),
        }
        (run_dir / "result.json").write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
        upsert_run(record)
        threading.Thread(target=run_analysis, args=(run_id,), daemon=True).start()
        self.send_json(record, HTTPStatus.ACCEPTED)

    def do_DELETE(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/session":
            self.send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)
            return
        token = get_cookie_value(self.headers.get("Cookie", ""), SESSION_COOKIE)
        drop_session(token)
        self.send_json({"authenticated": False}, headers={"Set-Cookie": clear_session_cookie()})

    def send_run_file(self, run_id: str, file_type: str) -> None:
        record = get_run(run_id)
        if not record:
            self.send_json({"error": "Run not found"}, HTTPStatus.NOT_FOUND)
            return
        file_info = ALLOWED_FILE_TYPES.get(file_type)
        if not file_info:
            self.send_json({"error": "Unsupported file type"}, HTTPStatus.BAD_REQUEST)
            return
        file_name, content_type = file_info
        file_path = RUNS_DIR / run_id / "output" / file_name
        if file_type == "docx" and not file_path.exists():
            docx_files = sorted((RUNS_DIR / run_id / "output").glob("*.docx"))
            if docx_files:
                file_path = docx_files[0]
                file_name = file_path.name
        if not file_path.exists():
            self.send_json({"error": "File not ready"}, HTTPStatus.NOT_FOUND)
            return
        body = file_path.read_bytes()
        download_name = f"{record.get('batchName') or run_id}-{file_name}"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(download_name)}")
        self.end_headers()
        self.wfile.write(body)

    def send_static(self, path: str) -> None:
        if path in ("", "/"):
            target = DIST_DIR / "index.html"
        else:
            target = (DIST_DIR / path.lstrip("/")).resolve()
            if not str(target).startswith(str(DIST_DIR.resolve())) or not target.exists():
                target = DIST_DIR / "index.html"

        if not target.exists():
            self.send_json({"error": "Frontend not built. Run npm run build in product-poc."}, HTTPStatus.NOT_FOUND)
            return
        body = target.read_bytes()
        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    if not INDEX_PATH.exists():
        save_index([])
    recover_interrupted_runs()
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", "8787"))
    host = os.environ.get("POC_HOST", "127.0.0.1")
    httpd = ThreadingHTTPServer((host, port), Handler)
    print(f"POC server running at http://{host}:{port}")
    httpd.serve_forever()


if __name__ == "__main__":
    main()
