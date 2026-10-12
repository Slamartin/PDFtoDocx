"""Public-service guards: size and page caps, rate limit, queue, timeout, crash handling, cleanup."""

import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

from fastapi.testclient import TestClient

from app import main
from app.config import settings
from app.main import app
from tests.conftest import post_file
from tests.make_fixtures import FIX

SLEEP_WORKER = [sys.executable, "-c", "import time; time.sleep(10)"]


def test_size_cap_content_length(client, monkeypatch):
    monkeypatch.setattr(settings, "max_upload_bytes", 1000)
    r = client.post("/api/convert", files={"file": ("big.pdf", b"%PDF" + b"x" * 70_000)})
    assert r.status_code == 413 and r.json()["error"] == "too_large"
    assert not any(settings.work_dir.iterdir())


def test_size_cap_chunked_body(client, monkeypatch):
    """No Content-Length: the bytes are counted while the body streams in."""
    monkeypatch.setattr(settings, "max_upload_bytes", 1000)
    boundary = "pdftowordtest"

    def body():
        yield (
            f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="a.pdf"\r\n'
            f"Content-Type: application/pdf\r\n\r\n"
        ).encode()
        for _ in range(200):
            yield b"%PDF" + b"x" * 1000
        yield f"\r\n--{boundary}--\r\n".encode()

    r = client.post(
        "/api/convert", content=body(), headers={"content-type": f"multipart/form-data; boundary={boundary}"}
    )
    assert r.status_code == 413 and r.json()["error"] == "too_large"


def test_client_side_limit_endpoint(client):
    r = client.get("/api/limits")
    assert r.status_code == 200
    body = r.json()
    assert body["max_upload_bytes"] == settings.max_upload_bytes and "source_url" in body


def test_page_caps(client, monkeypatch):
    monkeypatch.setattr(settings, "max_pages", 1)
    r = post_file(client, FIX / "sample.pdf")
    assert r.status_code == 413 and r.json()["error"] == "too_many_pages" and "2 pages" in r.json()["detail"]
    monkeypatch.setattr(settings, "max_slides", 3)
    r = post_file(client, FIX / "sample.pptx")
    assert r.status_code == 413 and "4 slides" in r.json()["detail"]
    # The password check comes before the count, so a locked file still asks for the password.
    r = post_file(client, FIX / "locked.pdf")
    assert r.status_code == 401 and r.json()["error"] == "password"
    assert not any(settings.work_dir.iterdir())


def test_rate_limit(client, monkeypatch):
    monkeypatch.setattr(settings, "rate_per_minute", 1)
    monkeypatch.setattr(settings, "rate_burst", 2)
    main.limits.reset()
    # Tokens are taken before the file type is checked, so cheap 415s exercise the limiter.
    assert post_file(client, FIX / "not-a-doc.txt").status_code == 415
    assert post_file(client, FIX / "not-a-doc.txt").status_code == 415
    r = post_file(client, FIX / "not-a-doc.txt")
    assert r.status_code == 429 and r.json()["error"] == "rate_limited"
    assert int(r.headers["retry-after"]) >= 1
    assert client.get("/healthz").status_code == 200  # other routes are not limited


def test_timeout_kills_worker(client, monkeypatch):
    monkeypatch.setattr(main, "_worker_argv", lambda job_dir: SLEEP_WORKER)
    monkeypatch.setattr(settings, "convert_timeout_s", 1)
    t0 = time.monotonic()
    r = post_file(client, FIX / "sample.pdf")
    assert r.status_code == 504 and r.json()["error"] == "timeout"
    assert time.monotonic() - t0 < 6
    assert not any(settings.work_dir.iterdir())
    assert not main._jobs


def test_queue_full_returns_503(client, monkeypatch):
    monkeypatch.setattr(main, "_worker_argv", lambda job_dir: [sys.executable, "-c", "import time; time.sleep(3)"])
    monkeypatch.setattr(settings, "convert_timeout_s", 10)
    monkeypatch.setattr(settings, "queue_max_waiting", 0)
    with ThreadPoolExecutor(2) as pool:
        first = pool.submit(post_file, client, FIX / "sample.pdf")
        time.sleep(0.8)  # let the first one occupy the slot
        second = pool.submit(post_file, client, FIX / "sample.pdf")
        statuses = sorted([first.result().status_code, second.result().status_code])
    assert statuses == [422, 503], statuses  # sleeping worker writes no result -> 422; the other is turned away


def test_crash_is_sanitised(client, monkeypatch):
    monkeypatch.setattr(main, "_worker_argv", lambda job_dir: [sys.executable, "-c", "import os; os.abort()"])
    r = post_file(client, FIX / "sample.pdf")
    assert r.status_code == 422 and r.json()["error"] == "failed"
    assert "Traceback" not in r.text and "abort" not in r.text


def test_junk_pdf_message_is_generic(client):
    r = client.post("/api/convert", files={"file": ("x.pdf", b"%PDF-1.4 junk junk junk")})
    assert r.status_code == 422
    assert r.json()["detail"] == main.GENERIC_FAIL
    assert "mupdf" not in r.text.lower() and "exception" not in r.text.lower()


def test_unexpected_error_is_generic(work_dir, monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("secret internals")

    monkeypatch.setattr(main, "_run_job", boom)
    with TestClient(app, raise_server_exceptions=False) as client:
        r = post_file(client, FIX / "sample.pdf")
        assert r.status_code == 500 and r.json() == {"error": "internal", "detail": "Internal error."}
        assert not any(settings.work_dir.iterdir())


def test_cleanup_expires_results_and_orphans(client, monkeypatch):
    r = post_file(client, FIX / "sample.pptx")
    assert r.status_code == 200, r.text
    job_id = r.json()["id"]
    assert client.get(f"/api/result/{job_id}").status_code == 200
    job_dir = settings.work_dir / job_id
    assert sorted(p.name for p in job_dir.iterdir()) == ["output.docx"]  # input and worker log are gone

    orphan = settings.work_dir / "deadbeef"
    orphan.mkdir()
    old = time.time() - 10 * 24 * 3600
    os.utime(orphan, (old, old))
    fresh = settings.work_dir / "fresh"
    fresh.mkdir()

    monkeypatch.setattr(settings, "result_ttl_s", 0)
    main._cleanup_once()
    assert client.get(f"/api/result/{job_id}").status_code == 404
    assert not job_dir.exists() and not orphan.exists() and fresh.exists()


def test_startup_wipes_work_dir(work_dir):
    work_dir.mkdir(parents=True)
    (work_dir / "stale").mkdir()
    with TestClient(app) as c:
        assert c.get("/healthz").json() == {"status": "ok"}
        assert not (work_dir / "stale").exists()


def test_security_headers_and_download_name(client):
    r = client.get("/")
    assert (
        r.headers["x-content-type-options"] == "nosniff" and "frame-ancestors" in r.headers["content-security-policy"]
    )
    r = post_file(client, FIX / "sample.pdf", name="wéird; náme.pdf")
    assert r.status_code == 200, r.text
    dl = client.get(f"/api/result/{r.json()['id']}")
    cd = dl.headers["content-disposition"]
    assert 'filename="w_ird_ n_me.docx"' in cd and "filename*=UTF-8''w%C3%A9ird%3B%20n%C3%A1me.docx" in cd
    assert dl.headers["cache-control"] == "no-store"
    assert client.get("/docs").status_code == 404
