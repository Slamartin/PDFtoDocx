"""Conversion worker: one process per job, so the server can kill it on timeout and survive
a crash inside MuPDF or OpenCV.

Protocol
    argv:    python -m app.worker <job_dir>
    stdin:   one JSON object {kind, src, dst, password, notes, max_pages, max_slides, memory_mb}
    stdout/stderr: redirected to <job_dir>/worker.log (library chatter, never sent to clients)
    result:  <job_dir>/result.json, written atomically:
             {"ok": true, "report": {...}}  or
             {"ok": false, "error": <code>, "detail": <text for the server log>, "pages": <n>}
    exit:    0 when result.json was written (also for expected errors), 1 for an unexpected
             exception, anything else is a native crash.
"""

from __future__ import annotations

import json
import os
import re
import sys
import traceback
import zipfile
from pathlib import Path

SLIDE_RE = re.compile(r"^ppt/slides/slide\d+\.xml$")
ZIP_MAX_ENTRIES = 5000
ZIP_MAX_UNCOMPRESSED = 300 * 1024 * 1024
MAX_IMAGE_PIXELS = 40_000_000


def _apply_memory_limit(memory_mb: int) -> None:
    """Cap the heap on POSIX. RLIMIT_DATA (Linux >= 4.7 counts heap and private mmaps) is
    preferred over RLIMIT_AS, which counts shared library mappings and would kill the import."""
    if memory_mb <= 0 or sys.platform == "win32":
        return
    import resource

    limit = memory_mb * 1024 * 1024
    for name in ("RLIMIT_DATA", "RLIMIT_AS"):
        res = getattr(resource, name, None)
        if res is None:
            continue
        try:
            resource.setrlimit(res, (limit, limit))
            break
        except (ValueError, OSError):
            continue
    try:
        resource.setrlimit(resource.RLIMIT_CORE, (0, 0))
    except (ValueError, OSError):
        pass


def _redirect_output(log_path: Path) -> None:
    sys.stdout.flush()
    sys.stderr.flush()
    fd = os.open(str(log_path), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    os.dup2(fd, 1)
    os.dup2(fd, 2)
    os.close(fd)


def pdf_page_count(src: Path, password: str | None) -> int:
    """Cheap check before the full conversion: password and page count only."""
    import pymupdf

    from app.report import CorruptFileError, PasswordError

    try:
        doc = pymupdf.open(str(src))
    except Exception as exc:
        raise CorruptFileError(f"Not a readable PDF: {exc}") from exc
    try:
        if doc.needs_pass and (not password or not doc.authenticate(password)):
            raise PasswordError("wrong or missing password")
        return doc.page_count
    finally:
        doc.close()


def pptx_slide_count(src: Path) -> int:
    """Zip-level sanity check and slide count without loading python-pptx."""
    from app.report import CorruptFileError

    class ZipBomb(CorruptFileError):
        pass

    try:
        with zipfile.ZipFile(src) as zf:
            infos = zf.infolist()
            if len(infos) > ZIP_MAX_ENTRIES or sum(i.file_size for i in infos) > ZIP_MAX_UNCOMPRESSED:
                raise ZipBomb("archive too large when unpacked")
            return sum(1 for i in infos if SLIDE_RE.match(i.filename))
    except zipfile.BadZipFile as exc:
        raise CorruptFileError(f"Not a readable PPTX: {exc}") from exc


def _convert(spec: dict) -> dict:
    from PIL import Image

    Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS

    from app.report import CorruptFileError, PasswordError

    src, dst = Path(spec["src"]), Path(spec["dst"])
    kind = spec["kind"]
    password = spec.get("password") or None
    try:
        if kind == ".pdf":
            pages = pdf_page_count(src, password)
            if pages > spec["max_pages"]:
                return {"ok": False, "error": "too_many_pages", "pages": pages, "detail": ""}
            from app.converters.pdf_to_docx import convert_pdf

            report = convert_pdf(src, dst, password=password)
        else:
            pages = pptx_slide_count(src)
            if pages > spec["max_slides"]:
                return {"ok": False, "error": "too_many_pages", "pages": pages, "detail": ""}
            from app.converters.pptx_to_docx import convert_pptx

            report = convert_pptx(src, dst, include_notes=bool(spec.get("notes")))
    except PasswordError as exc:
        return {"ok": False, "error": "password", "detail": str(exc)}
    except CorruptFileError as exc:
        code = "zip_bomb" if exc.__class__.__name__ == "ZipBomb" else "corrupt"
        return {"ok": False, "error": code, "detail": f"{exc.__class__.__name__}: {exc}"}
    except MemoryError:
        return {"ok": False, "error": "memory", "detail": "MemoryError"}
    if not dst.exists():
        return {"ok": False, "error": "internal", "detail": "converter produced no output"}
    return {"ok": True, "report": report.to_dict()}


def _write_result(job_dir: Path, result: dict) -> None:
    tmp = job_dir / "result.json.tmp"
    tmp.write_text(json.dumps(result), encoding="utf-8")
    os.replace(tmp, job_dir / "result.json")


def main(argv: list[str]) -> int:
    job_dir = Path(argv[1])
    spec = json.loads(sys.stdin.readline())
    _apply_memory_limit(int(spec.get("memory_mb") or 0))
    _redirect_output(job_dir / "worker.log")
    try:
        result = _convert(spec)
    except Exception:
        result = {"ok": False, "error": "internal", "detail": traceback.format_exc()}
    _write_result(job_dir, result)
    return 1 if result.get("error") == "internal" else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
