"""The worker process on its own, without the web layer."""

import json
import subprocess
import sys
import zipfile

from app import worker
from app.main import ROOT, _worker_env
from tests.make_fixtures import FIX


def _spec(kind, src, dst, **extra):
    base = {
        "kind": kind,
        "src": str(src),
        "dst": str(dst),
        "password": None,
        "notes": False,
        "max_pages": 80,
        "max_slides": 80,
        "memory_mb": 0,
    }
    base.update(extra)
    return base


def test_worker_process_roundtrip(tmp_path):
    spec = _spec(".pdf", FIX / "sample.pdf", tmp_path / "out.docx")
    proc = subprocess.run(
        [sys.executable, "-m", "app.worker", str(tmp_path)],
        input=(json.dumps(spec) + "\n").encode(),
        cwd=str(ROOT),
        env=_worker_env(),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        timeout=120,
    )
    assert proc.returncode == 0
    result = json.loads((tmp_path / "result.json").read_text())
    assert result["ok"] and result["report"]["pages"] == 2
    assert (tmp_path / "out.docx").exists() and (tmp_path / "worker.log").exists()


def test_worker_password_and_caps(tmp_path):
    res = worker._convert(_spec(".pdf", FIX / "locked.pdf", tmp_path / "o.docx"))
    assert res == {"ok": False, "error": "password", "detail": "wrong or missing password"}
    res = worker._convert(_spec(".pdf", FIX / "locked.pdf", tmp_path / "o.docx", password="secret", max_pages=1))
    assert res["error"] == "too_many_pages" and res["pages"] == 2
    res = worker._convert(_spec(".pptx", FIX / "sample.pptx", tmp_path / "o.docx", max_slides=3))
    assert res["error"] == "too_many_pages" and res["pages"] == 4


def test_worker_zip_checks(tmp_path):
    assert worker.pptx_slide_count(FIX / "sample.pptx") == 4
    bomb = tmp_path / "bomb.pptx"
    with zipfile.ZipFile(bomb, "w") as zf:
        for i in range(worker.ZIP_MAX_ENTRIES + 1):
            zf.writestr(f"ppt/slides/slide{i}.xml", "x")
    res = worker._convert(_spec(".pptx", bomb, tmp_path / "o.docx"))
    assert res["error"] == "zip_bomb"
    junk = tmp_path / "junk.pptx"
    junk.write_bytes(b"not a zip at all")
    res = worker._convert(_spec(".pptx", junk, tmp_path / "o.docx"))
    assert res["error"] == "corrupt"


def test_worker_rejects_non_pdf_named_pdf(tmp_path):
    """MuPDF would open these through its HTML, SVG and image handlers; only real PDFs may reach pdf2docx."""
    import pymupdf

    png = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 8, 8), False).tobytes("png")
    samples = {
        "html": b"<html><body><p>hello</p></body></html>",
        "svg": b'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><text y="5">x</text></svg>',
        "png": png,
    }
    for name, data in samples.items():
        src = tmp_path / f"{name}.pdf"
        src.write_bytes(data)
        res = worker._convert(_spec(".pdf", src, tmp_path / "o.docx"))
        assert res["ok"] is False and res["error"] == "corrupt", (name, res)
        assert not (tmp_path / "o.docx").exists()
