"""FastAPI app: serves the UI and converts uploaded PDF/PPTX files to DOCX."""

from __future__ import annotations

import atexit
import logging
import shutil
import tempfile
import threading
import time
import uuid
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from app.converters.pdf_to_docx import convert_pdf
from app.converters.pptx_to_docx import convert_pptx
from app.report import CorruptFileError, PasswordError

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
log = logging.getLogger(__name__)

STATIC = Path(__file__).parent / "static"
WORK = Path(tempfile.gettempdir()) / "pdftoword"
MAX_BYTES = 100 * 1024 * 1024
JOB_TTL_S = 60 * 60
SUPPORTED = {".pdf", ".pptx"}
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

app = FastAPI(title="PDF / PPTX to Word")
_jobs: dict[str, dict] = {}
_convert_lock = threading.Lock()  # pdf2docx is CPU-bound; one conversion at a time


def _sweep() -> None:
    now = time.time()
    for job_id, job in list(_jobs.items()):
        if now - job["created"] > JOB_TTL_S:
            shutil.rmtree(job["dir"], ignore_errors=True)
            _jobs.pop(job_id, None)


atexit.register(lambda: shutil.rmtree(WORK, ignore_errors=True))


def _kind(filename: str, head: bytes) -> str:
    ext = Path(filename).suffix.lower()
    if ext == ".pdf" or head.startswith(b"%PDF"):
        return ".pdf"
    return ext


def _run(kind: str, src: Path, dst: Path, password: str | None, notes: bool):
    with _convert_lock:
        if kind == ".pdf":
            return convert_pdf(src, dst, password=password)
        return convert_pptx(src, dst, include_notes=notes)


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.post("/api/convert")
async def convert(file: UploadFile = File(...), password: str | None = Form(None), notes: bool = Form(False)):
    _sweep()
    data = await file.read()
    if len(data) > MAX_BYTES:
        raise HTTPException(413, "File is larger than 100 MB.")
    if not data:
        raise HTTPException(422, "The file is empty.")
    name = Path(file.filename or "upload").name
    kind = _kind(name, data[:8])
    if kind not in SUPPORTED:
        raise HTTPException(415, "Only .pdf and .pptx files are supported.")

    job_id = uuid.uuid4().hex
    job_dir = WORK / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    src = job_dir / f"input{kind}"
    dst = job_dir / "output.docx"
    src.write_bytes(data)
    out_name = (Path(name).stem or "converted") + ".docx"
    try:
        report = await run_in_threadpool(_run, kind, src, dst, password or None, notes)
    except PasswordError:
        shutil.rmtree(job_dir, ignore_errors=True)
        return JSONResponse({"error": "password", "detail": "This PDF is password protected."}, status_code=401)
    except CorruptFileError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(422, str(exc))
    except Exception as exc:
        log.exception("conversion failed for %s", name)
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(500, f"Conversion failed: {exc.__class__.__name__}: {exc}")
    finally:
        src.unlink(missing_ok=True)

    _jobs[job_id] = {"dir": job_dir, "path": dst, "filename": out_name, "created": time.time()}
    return {"id": job_id, "filename": out_name, "size": dst.stat().st_size, "report": report.to_dict()}


@app.get("/api/result/{job_id}")
def result(job_id: str):
    _sweep()
    job = _jobs.get(job_id)
    if not job or not job["path"].exists():
        raise HTTPException(404, "This result has expired. Convert the file again.")
    ascii_name = job["filename"].encode("ascii", "replace").decode()
    headers = {"Content-Disposition": f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(job['filename'])}"}
    return FileResponse(job["path"], media_type=DOCX_MIME, headers=headers)


app.mount("/static", StaticFiles(directory=STATIC), name="static")
