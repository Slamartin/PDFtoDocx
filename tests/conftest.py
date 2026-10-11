import pytest
from fastapi.testclient import TestClient

from app.config import settings
from app.main import app, limits
from tests.make_fixtures import FIX, make_pdf, make_pptx


@pytest.fixture(scope="session", autouse=True)
def fixtures():
    FIX.mkdir(exist_ok=True)
    if not (FIX / "sample.pdf").exists():
        make_pdf(FIX / "sample.pdf")
    if not (FIX / "locked.pdf").exists():
        make_pdf(FIX / "locked.pdf", password="secret")
    if not (FIX / "sample.pptx").exists():
        make_pptx(FIX / "sample.pptx")
    (FIX / "not-a-doc.txt").write_text("hello")


@pytest.fixture
def work_dir(tmp_path, monkeypatch):
    d = tmp_path / "work"
    monkeypatch.setattr(settings, "work_dir", d)
    return d


@pytest.fixture
def client(work_dir):
    """Client with the lifespan running (startup wipe, cleanup task) and fresh rate-limit state."""
    limits.reset()
    with TestClient(app) as c:
        yield c
    limits.reset()


def post_file(client, path, name=None, **data):
    with open(path, "rb") as f:
        return client.post("/api/convert", files={"file": (name or path.name, f)}, data=data)
