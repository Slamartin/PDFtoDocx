"""Public-service guards: size and page caps, rate limit, queue, timeout, crash handling, cleanup."""

import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

from fastapi.testclient import TestClient

from app import main
from app.config import settings
from app.guards import SECURITY_HEADERS
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
    assert body["max_upload_bytes"] == settings.max_upload_bytes
    assert body["source_url"] == "https://github.com/Slamartin/PDFtoDocx"


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


def test_one_conversion_in_flight_per_client(client, monkeypatch):
    """One address cannot hold the running slot and the waiting slots at once."""
    monkeypatch.setattr(main, "_worker_argv", lambda job_dir: [sys.executable, "-c", "import time; time.sleep(2)"])
    monkeypatch.setattr(settings, "convert_timeout_s", 10)
    monkeypatch.setattr(settings, "rate_per_minute", 1)
    monkeypatch.setattr(settings, "rate_burst", 2)
    monkeypatch.setattr(settings, "trusted_proxy", True)
    main.limits.reset()
    pdf = (FIX / "sample.pdf").read_bytes()

    def post(ip=None):
        headers = {"fly-client-ip": ip} if ip else {}
        return client.post("/api/convert", files={"file": ("sample.pdf", pdf)}, headers=headers)

    with ThreadPoolExecutor(2) as pool:
        first = pool.submit(post)
        time.sleep(0.8)  # let the first one start
        r = post()
        assert r.status_code == 429 and r.json()["error"] == "rate_limited", r.text
        assert int(r.headers["retry-after"]) >= 1
        other = pool.submit(post, "203.0.113.7")  # a different address still gets in line
        assert first.result().status_code == 422  # sleeping worker writes no result
        assert other.result().status_code == 422
    # The turned-away request cost no token (burst 2), and the finished one no longer counts.
    assert post().status_code == 422
    assert post().status_code == 429


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


def test_html_disguised_as_pdf_is_rejected(client):
    r = client.post("/api/convert", files={"file": ("x.pdf", b"<html><body><p>not a pdf</p></body></html>")})
    assert r.status_code == 422 and r.json()["detail"] == main.GENERIC_FAIL
    assert not any(settings.work_dir.iterdir())


def test_unexpected_error_is_generic(work_dir, monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("secret internals")

    monkeypatch.setattr(main, "_run_job", boom)
    with TestClient(app, raise_server_exceptions=False) as client:
        r = post_file(client, FIX / "sample.pdf")
        assert r.status_code == 500 and r.json() == {"error": "internal", "detail": "Internal error."}
        assert r.headers["content-security-policy"] == SECURITY_HEADERS["content-security-policy"]
        assert not any(settings.work_dir.iterdir())


def test_cleanup_expires_results_and_orphans(client, monkeypatch):
    r = post_file(client, FIX / "sample.pptx")
    assert r.status_code == 200, r.text
    job_id = r.json()["id"]
    assert client.get(f"/api/result/{job_id}").status_code == 200
    job_dir = settings.work_dir / job_id
    assert sorted(p.name for p in job_dir.iterdir()) == ["output.docx"]  # input and worker log are gone

    old = time.time() - 10 * 24 * 3600
    orphan = settings.work_dir / ("deadbeef" * 4)
    foreign = settings.work_dir / "not-a-job"  # an old folder the app did not create stays
    for d in (orphan, foreign):
        d.mkdir()
        os.utime(d, (old, old))
    fresh = settings.work_dir / ("0f" * 16)
    fresh.mkdir()

    monkeypatch.setattr(settings, "result_ttl_s", 0)
    main._cleanup_once()
    assert client.get(f"/api/result/{job_id}").status_code == 404
    assert not job_dir.exists() and not orphan.exists() and fresh.exists() and foreign.exists()


def test_startup_and_shutdown_remove_only_job_dirs(work_dir):
    """WORK_DIR may point at an existing folder; only the job folders in it are ever deleted."""
    stale = work_dir / ("ab" * 16)
    stale.mkdir(parents=True)
    keep_dir, keep_file = work_dir / "my-files", work_dir / "notes.txt"
    keep_dir.mkdir()
    keep_file.write_text("keep")
    with TestClient(app) as c:
        assert c.get("/healthz").json() == {"status": "ok"}
        assert not stale.exists() and keep_dir.exists() and keep_file.exists()
        (work_dir / ("cd" * 16)).mkdir()
    assert work_dir.exists() and keep_dir.exists() and keep_file.read_text() == "keep"
    assert not (work_dir / ("cd" * 16)).exists()


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


def test_security_headers_on_guard_answers(client, monkeypatch):
    """The 413/429/503 come from middlewares outside the app; they must carry the headers as well."""

    def assert_headers(r, status):
        assert r.status_code == status, r.text
        for name, value in SECURITY_HEADERS.items():
            assert r.headers[name] == value, (status, name)

    assert_headers(client.get("/"), 200)
    assert "max-age=" in client.get("/").headers["strict-transport-security"]
    monkeypatch.setattr(settings, "max_upload_bytes", 1000)
    assert_headers(client.post("/api/convert", files={"file": ("big.pdf", b"%PDF" + b"x" * 70_000)}), 413)
    monkeypatch.setattr(settings, "rate_per_minute", 1)
    monkeypatch.setattr(settings, "rate_burst", 1)
    main.limits.reset()
    assert post_file(client, FIX / "not-a-doc.txt").status_code == 415
    assert_headers(post_file(client, FIX / "not-a-doc.txt"), 429)
    monkeypatch.setattr(main.queue, "full", lambda: True)
    assert_headers(post_file(client, FIX / "not-a-doc.txt"), 503)


def test_validation_errors_are_generic(client):
    """FastAPI's default 422 lists the failing fields and echoes the submitted value back."""
    r = client.post("/api/convert", files={"file": ("a.pdf", b"%PDF-1.4")}, data={"notes": "<b>echo-me</b>"})
    assert r.status_code == 422
    assert "echo-me" not in r.text, r.text
    assert r.json() == {"error": "invalid", "detail": main.INVALID_REQUEST}
    r = client.post("/api/convert", data={"file": "not-an-upload"})
    assert r.status_code == 422 and "not-an-upload" not in r.text
    assert r.headers["x-content-type-options"] == "nosniff"
