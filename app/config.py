"""Runtime settings, read once from environment variables.

`settings` is a plain (mutable) dataclass so tests can monkeypatch single attributes.
Everything in the app reads attributes at call time, never copies them into module constants.
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

_TRUE = {"1", "true", "yes", "on"}


def _int(name: str, default: int) -> int:
    return int(os.environ.get(name, default))


def _float(name: str, default: float) -> float:
    return float(os.environ.get(name, default))


def _bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    return default if raw is None else raw.strip().lower() in _TRUE


@dataclass
class Settings:
    host: str = "127.0.0.1"
    port: int = 8765
    # Uploads and work size
    max_upload_bytes: int = 25 * 1024 * 1024
    max_pages: int = 80  # PDF pages
    max_slides: int = 80  # PPTX slides
    # Worker process
    convert_timeout_s: float = 90.0
    worker_memory_mb: int = 0  # 0 = no RLIMIT in the child (POSIX only)
    # Queue: one conversion at a time, a few may wait
    queue_max_waiting: int = 3
    queue_wait_timeout_s: float = 60.0
    # Results
    result_ttl_s: float = 15 * 60
    cleanup_interval_s: float = 60.0
    # Per-IP rate limits on POST /api/convert; rate_per_minute <= 0 disables limiting
    rate_per_minute: float = 0.0
    rate_burst: int = 5
    rate_daily: int = 40
    rate_max_inflight: int = 1  # conversions in progress per client while limiting is on; 0 = no cap
    trusted_proxy: bool = False  # trust Fly-Client-IP; only safe behind the Fly proxy, which overwrites it
    # Misc
    work_dir: Path = field(default_factory=lambda: Path(tempfile.gettempdir()) / "pdftoword")
    public: bool = False  # switches UI copy from "runs on this PC" to hosted
    source_url: str = "https://github.com/Slamartin/PDFtoDocx"
    log_level: str = "INFO"

    @classmethod
    def from_env(cls) -> "Settings":
        d = cls()
        return cls(
            host=os.environ.get("HOST", d.host),
            port=_int("PORT", d.port),
            max_upload_bytes=_int("MAX_UPLOAD_BYTES", d.max_upload_bytes),
            max_pages=_int("MAX_PAGES", d.max_pages),
            max_slides=_int("MAX_SLIDES", d.max_slides),
            convert_timeout_s=_float("CONVERT_TIMEOUT_S", d.convert_timeout_s),
            worker_memory_mb=_int("WORKER_MEMORY_MB", d.worker_memory_mb),
            queue_max_waiting=_int("QUEUE_MAX_WAITING", d.queue_max_waiting),
            queue_wait_timeout_s=_float("QUEUE_WAIT_TIMEOUT_S", d.queue_wait_timeout_s),
            result_ttl_s=_float("RESULT_TTL_S", d.result_ttl_s),
            cleanup_interval_s=_float("CLEANUP_INTERVAL_S", d.cleanup_interval_s),
            rate_per_minute=_float("RATE_PER_MINUTE", d.rate_per_minute),
            rate_burst=_int("RATE_BURST", d.rate_burst),
            rate_daily=_int("RATE_DAILY", d.rate_daily),
            rate_max_inflight=_int("RATE_MAX_INFLIGHT", d.rate_max_inflight),
            trusted_proxy=_bool("TRUSTED_PROXY", d.trusted_proxy),
            work_dir=Path(os.environ.get("WORK_DIR", str(d.work_dir))),
            public=_bool("PUBLIC", d.public),
            source_url=os.environ.get("SOURCE_URL", d.source_url),
            log_level=os.environ.get("LOG_LEVEL", d.log_level).upper(),
        )

    @property
    def body_limit_bytes(self) -> int:
        """Multipart framing adds a little on top of the file itself."""
        return self.max_upload_bytes + 64 * 1024

    def public_limits(self) -> dict:
        return {
            "max_upload_bytes": self.max_upload_bytes,
            "max_pages": self.max_pages,
            "max_slides": self.max_slides,
            "convert_timeout_s": self.convert_timeout_s,
            "result_ttl_s": self.result_ttl_s,
            "rate_per_minute": self.rate_per_minute,
            "rate_burst": self.rate_burst,
            "rate_daily": self.rate_daily,
            "public": self.public,
            "source_url": self.source_url,
        }


settings = Settings.from_env()
