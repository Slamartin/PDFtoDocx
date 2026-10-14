# PDF / PPTX to Word

Web app that turns PDF and PowerPoint files into editable Word documents.
A Python server does the conversion; the browser page is only the front end.

It runs in two profiles:

- **Local**: start it on your own PC, nothing leaves your machine.
- **Hosted**: the same code behind a public URL, with upload caps, rate limits, a conversion timeout and
  short result retention. Files are deleted right after conversion; results expire after 15 minutes.
  Live at **https://pdftoword.fly.dev**. See [privacy and terms](app/static/privacy.html).

## Run locally

Requires Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```powershell
.\run.ps1
```

This installs dependencies, starts the server on http://127.0.0.1:8765 and opens the page.
Drop files on the page, wait for the "Ready" badge, click **Save .docx**.
On Linux or macOS: `uv sync && uv run uvicorn app.main:app --host 127.0.0.1 --port 8765`.

## What it does

| Input | Engine | Kept |
|-------|--------|------|
| PDF | pdf2docx (PyMuPDF) | paragraphs, fonts, sizes, colours, bold/italic, tables, images, multi-column text, page size and margins |
| PPTX | python-pptx + python-docx | one slide per page at the slide's size; fonts, sizes, colours and alignment resolved through layout, master and theme; bullets and numbering; shapes placed side by side become columns; filled boxes keep their background; tables with merged cells; pictures with crop and rotation applied; optional speaker notes |

## Limits

- Scanned PDFs have no text layer; those pages are kept as pictures. No OCR.
- PDF hyperlinks, footnotes and form fields become plain text.
- PPTX charts, SmartArt and embedded objects become a placeholder line. EMF/WMF/SVG and linked pictures are skipped.
- PPTX animations, transitions, slide backgrounds and layout artwork (logos, footers) are not carried over.
- Exact slide positions cannot be reproduced in a flowing Word document; vertical gaps and side-by-side columns approximate them.
- Password-protected PDFs are supported; the page asks for the password.

## How the hosted profile protects itself

- Every conversion runs in its own worker process (`app/worker.py`) with a wall-clock timeout and, on
  Linux, a memory limit. A crash inside MuPDF or OpenCV kills only that job.
- One conversion runs at a time; a few requests may wait, the rest get `503` with `Retry-After`.
- Uploads are capped before the body is read (`413`), and PDFs/PPTX files over the page or slide cap are
  rejected after a cheap inspection, before the full conversion. That inspection also rejects files that
  are not real PDFs, because MuPDF would otherwise open HTML, SVG or images through its other parsers.
- Per-IP token buckets (per minute and per day) on the convert endpoint return `429` with `Retry-After`.
- Error responses are generic; library error text goes to the server log with a job id only.
- The page makes no third-party requests, sets no cookies and has no analytics.

## Configuration

All settings are environment variables with defaults (`app/config.py`). The Docker image sets the public
profile; the local launcher uses the defaults.

| Variable | Default | Docker | Meaning |
|---|---|---|---|
| `HOST` / `PORT` | `127.0.0.1` / `8765` | `0.0.0.0` / `8080` | bind address |
| `MAX_UPLOAD_BYTES` | 25 MB | 25 MB | upload cap |
| `MAX_PAGES` / `MAX_SLIDES` | 80 / 80 | 80 / 80 | page and slide caps |
| `CONVERT_TIMEOUT_S` | 90 | 90 | worker wall clock |
| `WORKER_MEMORY_MB` | 0 (off) | 700 | RLIMIT for the worker, POSIX only |
| `QUEUE_MAX_WAITING` / `QUEUE_WAIT_TIMEOUT_S` | 3 / 60 | 3 / 60 | queue depth and max wait |
| `RESULT_TTL_S` / `CLEANUP_INTERVAL_S` | 900 / 60 | 900 / 60 | result retention |
| `RATE_PER_MINUTE` / `RATE_BURST` / `RATE_DAILY` | 0 (off) / 5 / 40 | 2 / 5 / 40 | per-IP limits; `0` disables |
| `TRUSTED_PROXY` | 0 | 1 | trust `Fly-Client-IP` / `X-Forwarded-For` |
| `WORK_DIR` | `<tmp>/pdftoword` | `/tmp/pdftoword` | job folders; only folders named like a job id are ever deleted |
| `PUBLIC` | 0 | 1 | hosted wording in the UI |
| `SOURCE_URL` | GitHub URL | | footer link |
| `LOG_LEVEL` | `INFO` | | |

## Self-host with Docker

```sh
docker build -t pdftoword .
docker run --rm -p 8080:8080 -e MAX_UPLOAD_BYTES=52428800 pdftoword
```

## Deploy on Fly.io

The repo contains `fly.toml` (one `shared-cpu-1x` machine with 1 GB in `ams`, stops when idle). Images are
built on Fly's remote builder, so Docker is not needed locally. Fly asks for billing details before it
creates an app.

```sh
fly apps create pdftoword            # another name: change `app` in fly.toml too
fly ips allocate-v6                  # free
fly ips allocate-v4 --shared         # free; a dedicated IPv4 costs extra
fly deploy --remote-only --ha=false  # one machine instead of the default two
fly logs
```

`fly launch` works too, but it asks questions and may rewrite `fly.toml`.

GitHub Actions (`.github/workflows/ci.yml`) runs the tests, builds and smoke-tests the image on every
push, and deploys `main` with the `FLY_API_TOKEN` repository secret. The live app uses a deploy token
that only covers this app and expires after a year:

```sh
gh secret set FLY_API_TOKEN --body "$(fly tokens create deploy -a pdftoword -x 8760h)"
```

Measured on the live machine (2026-10-13): the Fly proxy delivers a response that stays silent for the
full 90 s timeout, so no polling API is needed for the free tier. The shared CPU is about three times
slower than a desktop: a dense synthetic PDF converts at roughly 3.5 s per page, so heavy PDFs of 20+
pages end in the `504` well before the page cap. If the worker runs out of memory, `fly logs` shows
`worker crashed rc=-9`; raise `WORKER_MEMORY_MB` or the machine size.

Measured in the image (python 3.14-slim, one CPU): the server process idles at about 57 MB; the worker
imports pdf2docx with 114 MB of data segment and converts a 23-page, 148-image PDF with a 273 MB peak
under the 700 MB `RLIMIT_DATA`. The thread variables the server passes to the worker
(`OMP_NUM_THREADS=1`, `OPENBLAS_NUM_THREADS=1`, `OPENCV_FOR_THREADS_NUM=1`, `MALLOC_ARENA_MAX=2`) are
what make this possible: without them OpenBLAS and OpenCV reserve about 2.8 GB at import and the limit
kills the import.

## Tests

```powershell
uv run python tests/make_fixtures.py
uv run pytest
```

## Licence

AGPL-3.0-only (see `LICENSE`). The PDF path depends on [PyMuPDF](https://pymupdf.readthedocs.io/) and
[pdf2docx](https://github.com/ArtifexSoftware/pdf2docx), which are AGPL/GPL licensed; anyone who runs
this as a network service must offer its source to the users, which the page's footer link does.
