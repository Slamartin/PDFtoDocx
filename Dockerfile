# Public-service image. Built remotely by `fly deploy --remote-only` or GitHub Actions.
FROM python:3.14-slim

COPY --from=ghcr.io/astral-sh/uv:0.12.10 /uv /bin/uv

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_NO_DEV=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Dependencies first, so code changes do not invalidate this layer.
COPY pyproject.toml uv.lock ./
RUN --mount=type=cache,target=/root/.cache/uv uv sync --locked --no-dev

COPY app ./app

# The code and .venv stay root-owned: the worker runs as `app` and parses untrusted files, so an exploit in
# MuPDF/OpenCV/Pillow must not be able to patch site-packages that every later worker imports.
RUN useradd -r -u 10001 -d /app -s /usr/sbin/nologin app \
    && mkdir -p /tmp/pdftoword \
    && chown app:app /tmp/pdftoword

USER app

# Public profile: binds to all interfaces, rate limits on. TRUSTED_PROXY stays off here and is set in
# fly.toml: outside Fly, clients could send their own Fly-Client-IP and pick their rate-limit bucket.
ENV PATH="/app/.venv/bin:$PATH" \
    HOST=0.0.0.0 \
    PORT=8080 \
    WORK_DIR=/tmp/pdftoword \
    PUBLIC=1 \
    MAX_UPLOAD_BYTES=26214400 \
    MAX_PAGES=80 \
    MAX_SLIDES=80 \
    CONVERT_TIMEOUT_S=90 \
    WORKER_MEMORY_MB=700 \
    QUEUE_MAX_WAITING=3 \
    RESULT_TTL_S=900 \
    RATE_PER_MINUTE=2 \
    RATE_BURST=5 \
    RATE_DAILY=40

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s \
    CMD python -c "import sys, urllib.request; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8080/healthz', timeout=3).status == 200 else 1)"

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080", "--workers", "1", "--timeout-keep-alive", "5", "--no-server-header"]
