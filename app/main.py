"""FastAPI app: serves the UI and converts uploaded PDF/PPTX files to DOCX.

Each conversion runs in a separate worker process (see app/worker.py) with a timeout, behind a
queue that allows one conversion at a time and a few waiting requests. Uploads are capped
before they are read, clients are rate limited per IP, and results expire after a short TTL.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
from contextlib import asynccontextmanager
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from app.config import settings
from app.guards import SECURITY_HEADERS, BodyLimitMiddleware, RateLimitMiddleware, SecurityHeadersMiddleware
from app.ratelimit import RateLimits

logging.basicConfig(level=settings.log_level, format="%(levelname)s %(name)s: %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).resolve().parent.parent
STATIC = Path(__file__).parent / "static"
SUPPORTED = {".pdf", ".pptx"}
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
GENERIC_FAIL = "The file could not be converted. It may be damaged or use features this converter does not handle."
JOB_ID_RE = re.compile(r"^[0-9a-f]{32}$")

_jobs: dict[str, dict] = {}
limits = RateLimits()


class Busy(Exception):
    pass


class ConvertQueue:
    """One conversion at a time; at most `queue_max_waiting` requests wait for their turn.
    Thread based because the conversion itself blocks a threadpool thread anyway."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._sem = threading.Semaphore(1)
        self.running = False
        self.waiting = 0

    def full(self) -> bool:
        return self.running and self.waiting >= settings.queue_max_waiting

    def enter(self) -> None:
        with self._lock:
            if self.full():
                raise Busy()
            self.waiting += 1
        ok = self._sem.acquire(timeout=settings.queue_wait_timeout_s)
        with self._lock:
            self.waiting -= 1
            if ok:
                self.running = True
        if not ok:
            raise Busy()

    def leave(self) -> None:
        with self._lock:
            self.running = False
        self._sem.release()


queue = ConvertQueue()


# ---------- housekeeping ----------


def _job_dirs() -> list[Path]:
    """Job folders this app created in the work dir (named by uuid4 hex). Nothing else in there is ever
    deleted, so a WORK_DIR that points at a shared or existing folder cannot wipe it."""
    try:
        return [
            d
            for d in settings.work_dir.iterdir()
            if JOB_ID_RE.match(d.name) and not d.is_symlink() and not d.is_junction() and d.is_dir()
        ]
    except OSError:
        return []


def _remove_job_dirs() -> None:
    for d in _job_dirs():
        shutil.rmtree(d, ignore_errors=True)


def _cleanup_once() -> None:
    now = time.time()
    for job_id, job in list(_jobs.items()):
        if now - job["created"] > settings.result_ttl_s:
            shutil.rmtree(job["dir"], ignore_errors=True)
            _jobs.pop(job_id, None)
    orphan_age = 2 * settings.convert_timeout_s + settings.queue_wait_timeout_s
    for d in _job_dirs():
        try:
            if d.name not in _jobs and now - d.stat().st_mtime > orphan_age:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass
    limits.evict()


async def _cleanup_loop() -> None:
    while True:
        await asyncio.sleep(settings.cleanup_interval_s)
        try:
            await run_in_threadpool(_cleanup_once)
        except Exception:
            log.exception("cleanup failed")


@asynccontextmanager
async def lifespan(_: FastAPI):
    settings.work_dir.mkdir(parents=True, exist_ok=True)
    _remove_job_dirs()  # a hard kill skips the shutdown branch
    _jobs.clear()
    task = asyncio.create_task(_cleanup_loop())
    try:
        yield
    finally:
        task.cancel()
        _remove_job_dirs()


app = FastAPI(title="PDF / PPTX to Word", lifespan=lifespan, docs_url=None, redoc_url=None, openapi_url=None)
# The last one added is the outermost.
app.add_middleware(BodyLimitMiddleware)
app.add_middleware(RateLimitMiddleware, limits=limits, queue_full=lambda: queue.full())
app.add_middleware(SecurityHeadersMiddleware)


@app.exception_handler(Exception)
async def _unhandled(request: Request, exc: Exception):
    log.exception("unhandled error on %s %s", request.method, request.url.path)
    return JSONResponse({"error": "internal", "detail": "Internal error."}, status_code=500, headers=SECURITY_HEADERS)


# ---------- worker process ----------

_ENV_KEEP = (
    "PATH",
    "SYSTEMROOT",
    "SYSTEMDRIVE",
    "WINDIR",
    "COMSPEC",
    "TEMP",
    "TMP",
    "TMPDIR",
    "HOME",
    "USERPROFILE",
    "LANG",
    "LC_ALL",
    "LC_CTYPE",
    "PYTHONUTF8",
)


def _worker_argv(job_dir: Path) -> list[str]:
    return [sys.executable, "-m", "app.worker", str(job_dir)]


def _worker_env() -> dict[str, str]:
    env = {k: v for k, v in os.environ.items() if k.upper() in _ENV_KEEP}
    env.update(
        {
            "OMP_NUM_THREADS": "1",
            "OPENBLAS_NUM_THREADS": "1",
            "OPENCV_FOR_THREADS_NUM": "1",
            "MALLOC_ARENA_MAX": "2",
            "PYTHONDONTWRITEBYTECODE": "1",
        }
    )
    return env


def _log_tail(job_dir: Path, limit: int = 4096) -> str:
    try:
        data = (job_dir / "worker.log").read_bytes()
    except OSError:
        return ""
    return data[-limit:].decode("utf-8", "replace")


def _run_job(job_dir: Path, spec: dict) -> dict:
    """Blocking: waits for a queue slot, then runs the worker process with a timeout."""
    queue.enter()
    try:
        kwargs = {"start_new_session": True} if sys.platform != "win32" else {}
        try:
            proc = subprocess.run(
                _worker_argv(job_dir),
                input=(json.dumps(spec) + "\n").encode(),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                cwd=str(ROOT),
                env=_worker_env(),
                timeout=settings.convert_timeout_s,
                **kwargs,
            )
        except subprocess.TimeoutExpired:
            return {"ok": False, "error": "timeout"}
    finally:
        queue.leave()

    result_path = job_dir / "result.json"
    if not result_path.exists():
        log.error("worker crashed job=%s rc=%s log=%r", job_dir.name, proc.returncode, _log_tail(job_dir))
        return {"ok": False, "error": "memory" if proc.returncode in (-9, 137) else "crashed"}
    try:
        return json.loads(result_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        log.error("worker result unreadable job=%s: %s", job_dir.name, exc)
        return {"ok": False, "error": "crashed"}


def _error_response(result: dict, kind: str) -> JSONResponse:
    code = result.get("error")
    if code == "password":
        return JSONResponse({"error": "password", "detail": "This PDF is password protected."}, status_code=401)
    if code == "too_many_pages":
        unit = "slides" if kind == ".pptx" else "pages"
        cap = settings.max_slides if kind == ".pptx" else settings.max_pages
        return JSONResponse(
            {"error": "too_many_pages", "detail": f"The file has {result.get('pages')} {unit}; the limit is {cap}."},
            status_code=413,
        )
    if code == "timeout":
        return JSONResponse(
            {
                "error": "timeout",
                "detail": f"Conversion took longer than {int(settings.convert_timeout_s)} s. Try a smaller or simpler file.",
            },
            status_code=504,
        )
    if code == "memory":
        return JSONResponse(
            {"error": "failed", "detail": "This file needs more memory than this service allows."}, status_code=422
        )
    return JSONResponse({"error": "failed", "detail": GENERIC_FAIL}, status_code=422)


# ---------- routes ----------


def _kind(filename: str, head: bytes) -> str:
    ext = Path(filename).suffix.lower()
    if ext == ".pdf" or head.startswith(b"%PDF"):
        return ".pdf"
    return ext


def _save_upload(upload: UploadFile, dst: Path) -> int:
    """Copy the (already spooled) upload to the job dir in 1 MB chunks; second line of defence
    for the size cap. Returns the byte count."""
    size = 0
    upload.file.seek(0)
    with open(dst, "wb") as out:
        while True:
            chunk = upload.file.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > settings.max_upload_bytes:
                raise HTTPException(413, f"File is larger than {settings.max_upload_bytes // (1024 * 1024)} MB.")
            out.write(chunk)
    return size


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html", headers={"Cache-Control": "no-cache"})


@app.get("/privacy")
def privacy():
    return FileResponse(STATIC / "privacy.html", headers={"Cache-Control": "no-cache"})


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.get("/api/limits")
def api_limits():
    return settings.public_limits()


@app.post("/api/convert")
async def convert(file: UploadFile = File(...), password: str | None = Form(None), notes: bool = Form(False)):
    name = Path(file.filename or "upload").name
    head = await file.read(8)
    if not head:
        raise HTTPException(422, "The file is empty.")
    kind = _kind(name, head)
    if kind not in SUPPORTED:
        raise HTTPException(415, "Only .pdf and .pptx files are supported.")

    job_id = uuid.uuid4().hex
    job_dir = settings.work_dir / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    src = job_dir / f"input{kind}"
    dst = job_dir / "output.docx"
    out_name = (Path(name).stem or "converted") + ".docx"
    spec = {
        "kind": kind,
        "src": str(src),
        "dst": str(dst),
        "password": password or None,
        "notes": bool(notes),
        "max_pages": settings.max_pages,
        "max_slides": settings.max_slides,
        "memory_mb": settings.worker_memory_mb,
    }
    try:
        await run_in_threadpool(_save_upload, file, src)
        result = await run_in_threadpool(_run_job, job_dir, spec)
    except Busy:
        shutil.rmtree(job_dir, ignore_errors=True)
        return JSONResponse(
            {"error": "busy", "detail": "The server is busy. Try again in a moment."},
            status_code=503,
            headers={"Retry-After": "15"},
        )
    except Exception:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    finally:
        src.unlink(missing_ok=True)

    if not result.get("ok"):
        if result.get("error") not in ("password", "too_many_pages"):
            log.warning(
                "conversion failed job=%s kind=%s error=%s detail=%s",
                job_id,
                kind,
                result.get("error"),
                (result.get("detail") or "")[:2000],
            )
        shutil.rmtree(job_dir, ignore_errors=True)
        return _error_response(result, kind)

    (job_dir / "worker.log").unlink(missing_ok=True)
    (job_dir / "result.json").unlink(missing_ok=True)
    _jobs[job_id] = {"dir": job_dir, "path": dst, "filename": out_name, "created": time.time()}
    return {"id": job_id, "filename": out_name, "size": dst.stat().st_size, "report": result["report"]}


@app.get("/api/result/{job_id}")
def result(job_id: str):
    job = _jobs.get(job_id)
    if not job or time.time() - job["created"] > settings.result_ttl_s or not job["path"].exists():
        raise HTTPException(404, "This result has expired. Convert the file again.")
    stem = re.sub(r"[^A-Za-z0-9._ -]+", "_", Path(job["filename"]).stem).strip(" .")[:100] or "converted"
    headers = {
        "Content-Disposition": f"attachment; filename=\"{stem}.docx\"; filename*=UTF-8''{quote(job['filename'])}",
        "Cache-Control": "no-store",
    }
    return FileResponse(job["path"], media_type=DOCX_MIME, headers=headers)


app.mount("/static", StaticFiles(directory=STATIC), name="static")
