# Spec: PDFtoWord

Source of truth for the project. Update it before changing behaviour, not after.
Status: approved 2026-10-12.

## Objective

A web app that converts PDF and PPTX files into editable Word documents (.docx) while keeping as much
formatting as a flowing Word document can hold. It runs in two profiles from one codebase:

- **Local**: Martin starts it on his PC with `run.ps1`; nothing leaves the machine.
- **Hosted**: the same code on one small Fly.io machine, open to anonymous users, protected by upload,
  page, rate and time limits. Free; a paid "more usage" tier is a deferred phase 2 and out of scope here.

Users: anyone with a PDF or PowerPoint file who wants editable text, tables and pictures in Word, without
installing Office. They drop files on a page, see a status per file, download the .docx.

Non-goals: OCR of scanned pages, accounts or logins, exact pixel reproduction of slide layouts, Office COM
or LibreOffice at runtime, more than one conversion at a time.

Acceptance criteria (the product is "working" when all hold):

- A PDF with text, tables, images and multi-column text converts to a .docx that keeps paragraphs, fonts,
  sizes, colours, bold/italic, tables, images, page size and margins.
- A PPTX converts to one Word page per slide at the slide's size, with titles, bullets, numbering, theme
  fonts and colours resolved through layout/master/theme, filled boxes, tables with merged cells, pictures
  with crop and rotation, and optional speaker notes.
- Password-protected PDFs are detected, the page asks for the password, the right password converts.
- Hosted profile: oversize uploads get 413 before the body is read, too many pages get 413 after a cheap
  inspection, a flood from one IP gets 429 with `Retry-After`, a full queue gets 503, a hung or crashed
  worker gets 504/422 and the server stays up, results disappear after the retention time, and no error
  response contains library text, tracebacks or the user's filename.
- The page makes no third-party requests, sets no cookies, links to the AGPL source and a privacy page.

## Tech Stack

- Python >= 3.12, managed with uv (`uv.lock` is authoritative; local venv runs CPython 3.14).
- FastAPI 0.143 + uvicorn 0.54 (one worker, one event loop), python-multipart.
- PDF: pdf2docx 0.5.13 on PyMuPDF 1.28 (AGPL-3.0; this is why the project is AGPL).
- PPTX: python-pptx 1.0, python-docx 1.2, Pillow 12, lxml.
- Front end: a single static `index.html` with inline CSS and vanilla JS, system fonts only.
- Tests: pytest 9 with FastAPI `TestClient` (httpx).
- Runtime: Docker image on `python:3.14-slim`, non-root, deployed to Fly.io (`fly.toml`), built remotely.
- CI: GitHub Actions (`.github/workflows/ci.yml`): tests, image build + container smoke test, deploy on main.
- Licence: AGPL-3.0-only.

## Commands

```powershell
# Local run (installs deps, starts http://127.0.0.1:8765, opens the browser)
.\run.ps1
# Same thing by hand, any OS
uv sync
uv run uvicorn app.main:app --host 127.0.0.1 --port 8765

# Tests (fixtures are generated on first run by tests/conftest.py)
uv run pytest -q
uv run pytest -q tests/test_api_limits.py -k rate_limit     # one area
uv run python tests/make_fixtures.py                        # regenerate fixtures

# Lint and format (ruff is added as the first task after this spec is approved)
uv run ruff check .
uv run ruff format .

# Container (Docker Desktop installed 2026-10-11)
docker build -t pdftoword .
docker run --rm -p 8080:8080 pdftoword
curl -F file=@tests/fixtures/sample.pdf http://localhost:8080/api/convert

# Deploy (flyctl not installed yet; see README)
fly deploy --remote-only
fly logs
```

Public-profile run on the PC for manual checks (rate limiting is off by default locally):

```powershell
$env:PUBLIC='1'; $env:RATE_PER_MINUTE='2'; uv run uvicorn app.main:app --port 8766
```

## Project Structure

```
app/
  main.py              FastAPI app: routes, worker launch, queue, cleanup, lifespan. No converter imports.
  config.py            Settings dataclass read from env vars once; read attributes at call time.
  guards.py            Pure ASGI middlewares: rate limit, body limit, security headers.
  ratelimit.py         Token buckets per client IP, client_ip() with proxy trust.
  worker.py            One process per job: limits, cheap inspection, converter call, result.json.
  report.py            ConversionReport dataclass, PasswordError, CorruptFileError.
  converters/
    pdf_to_docx.py     pdf2docx wrapper + opacity patch + inspect_pdf().
    pptx_to_docx.py    slide -> Word page rebuild: theme resolver, layout rows, tables, pictures, notes.
  static/
    index.html         the whole UI (CSS + JS inline)
    privacy.html       privacy and terms page served at /privacy
tests/
  conftest.py          fixture generation, `client` (lifespan on, fresh limiter), post_file()
  make_fixtures.py     builds sample.pdf, locked.pdf, sample.pptx into tests/fixtures/ (gitignored)
  test_convert.py      converter output assertions + API roundtrip + basic API errors
  test_api_limits.py   size/page caps, rate limit, queue, timeout, crash, cleanup, headers
  test_ratelimit.py    token bucket and client_ip unit tests
  test_worker.py       worker process and its checks without the web layer
Dockerfile, fly.toml, .dockerignore, .github/workflows/ci.yml
README.md (user-facing), SPEC.md (this file), LICENSE (AGPL-3.0), run.ps1
```

Rules of placement: anything that touches untrusted bytes before parsing lives in `guards.py` or
`worker.py`; anything that changes DOCX output lives in `converters/`; `main.py` orchestrates and must
stay importable without cv2/numpy (the server process idles at ~57 MB because of this).

## Code Style

Example of the house style (from `app/guards.py`):

```python
class BodyLimitMiddleware:
    """Rejects oversize uploads early.

    1. Content-Length over the cap: answer 413 at once, nothing is read.
    2. Chunked or lying clients: count the bytes handed to the app; once over the cap, report a
       disconnect to the app (so it stops reading) and rewrite whatever it answers into a 413.
    """

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if not _is_convert(scope):
            await self.app(scope, receive, send)
            return
        cap = settings.body_limit_bytes
        ...
```

Conventions:

- English everywhere: code, comments, log messages, UI text, commit messages.
- Line length 120. `from __future__ import annotations` at the top of every module. Type hints on
  function signatures; no hints for obvious locals.
- Module docstring says what the file is for; function docstrings only where the *why* is not obvious.
  Comments explain reasons and gotchas (see the opacity patch), never restate the code.
- Private helpers are `_snake_case`; module-level singletons (`settings`, `queue`, `limits`) are lowercase.
- Settings are read as `settings.x` at call time, never copied into module constants, so tests can
  monkeypatch them.
- Errors to clients are generic strings chosen in `_error_response`; library text goes to the log with a
  job id, never with the user's filename.
- Front end: ES5-style vanilla JS inside an IIFE, DOM built with `textContent` (never `innerHTML`),
  CSS variables for theming with a dark-mode block, no external resources.
- Ruff with `line-length = 120`, default rule set plus `I` (import sorting); format with `ruff format`.
- Commits: imperative subject line, body explains why, ends with the Claude co-author line when Claude
  wrote it.

## Testing Strategy

- Framework: pytest. Run `uv run pytest -q`; CI runs the same on Ubuntu, then builds the Docker image and
  posts `sample.pdf` to the running container.
- Fixtures are generated, not committed (`tests/make_fixtures.py`), so a test can assert exact content
  (2 PDF pages, 4 slides, password `secret`).
- Levels:
  - **Converter tests** (`test_convert.py`): open the produced .docx with python-docx and assert
    structure (styles, colours, merged cells, inline shapes, page size). This is the fidelity regression net.
  - **API tests** with `TestClient`: every guard and error path has a test that asserts status code, JSON
    `error` code and that nothing sensitive leaks. Worker behaviour is simulated by monkeypatching
    `main._worker_argv` to a sleeping or aborting `python -c` script.
  - **Unit tests** for pure logic (token bucket with a fake clock, `client_ip`, zip checks).
  - **Worker tests**: the real worker process via `subprocess.run`, and `_convert` in-process.
- Bar: all tests pass on Windows and in CI; every bug fix adds a regression test; every new endpoint, guard
  or error code adds an API test. No numeric coverage threshold.
- Fidelity changes in `converters/` are not done until the visual loop has been run on both real decks:
  export the original with PowerPoint COM to PDF, export the DOCX with Word COM to PDF, render both with
  PyMuPDF, compare page by page, and check the Word page count equals the slide count. Record the result
  in the commit message.
- Manual hosted check after a deploy: a 30 MB upload -> 413, six quick uploads -> 429, `fly status` shows
  the machine alive after a crafted bad file.

## Boundaries

- **Always**
  - Run `uv run pytest -q` before every commit; run `ruff check` and `ruff format` once ruff is in.
  - Keep the pure-Python stack: no Office COM, LibreOffice or other external binaries at runtime.
  - Keep converter imports out of `app/main.py`; everything heavy runs in `app/worker.py`.
  - Keep the worker's thread env vars (`OMP_NUM_THREADS=1` etc.); without them the memory limit kills
    the import (measured: 2.8 GB vs 114 MB).
  - Validate and cap every untrusted input before parsing it; return generic error text.
  - Write tests for new behaviour; update this spec and the README when behaviour or limits change.
  - Keep the UI free of third-party requests and the privacy page truthful.
  - Keep dependencies AGPL-compatible and pinned through `uv.lock` (`uv add`, never hand-edit).
- **Ask first**
  - Changing public limits or Fly configuration: Dockerfile env defaults, `fly.toml`, machine size or
    region, rate-limit numbers, timeouts, retention.
  - Changing the CI workflow or deploy pipeline (`.github/workflows/ci.yml`, deploy job, secrets).
  - Anything that would make the service non-anonymous or add persistence (accounts, databases, volumes).
- **Never**
  - Commit secrets, tokens or `.env` files; never log passwords or file contents.
  - Reintroduce reading the whole upload into memory or running a conversion on a server thread.
  - Return library exception text, tracebacks or the uploaded filename in an HTTP response.
  - Remove or weaken a failing test to get CI green.
  - Replace the licence or add a dependency that conflicts with AGPL-3.0.
  - Edit `uv.lock` by hand or commit `tests/fixtures/`.

## Success Criteria

- `uv run pytest -q` passes (26 tests today) on Windows and in CI.
- `docker build` succeeds and the container converts `sample.pdf` and both real decks with the 700 MB
  worker limit (verified 2026-10-11: 23-page, 148-image PDF peaks at 273 MB).
- Hosted checks above (413, 429, 503, 504, generic 422, retention) behave as listed.
- The page shows real limits from `/api/limits`, retries on 429/503, and never contacts a third party.
- Fly deploy succeeds and a 90 s conversion is not cut by the proxy (or the timeout is lowered and this spec
  updated).

## Open Questions

1. Fly proxy timeout for a silent long response: unverified until the first deploy. If it cuts at ~60 s,
   either lower `CONVERT_TIMEOUT_S` or move to a `202` + `GET /api/jobs/{id}` polling API (also what the
   paid tier needs).
2. Resolved 2026-10-12: the repository is `https://github.com/Slamartin/PDFtoDocx` (set in `config.py`,
   `fly.toml`, `pyproject.toml` and `index.html`).
3. Resolved 2026-10-12: an oversize (413) upload counts against the rate limit, on purpose.
4. Custom domain for the hosted service, or keep `*.fly.dev`.
