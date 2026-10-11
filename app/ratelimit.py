"""In-memory token-bucket rate limiting keyed by client IP.

Single-instance only: the state lives in this process. Good enough for one Fly machine.
"""

from __future__ import annotations

import ipaddress
import math
import time
from dataclasses import dataclass
from typing import Callable

from app.config import settings


@dataclass
class _Bucket:
    tokens: float
    updated: float


class TokenBucketLimiter:
    def __init__(self, rate_per_s: float, burst: int, clock: Callable[[], float] = time.monotonic, max_keys: int = 10_000):
        self.rate = rate_per_s
        self.burst = max(1, burst)
        self.clock = clock
        self.max_keys = max_keys
        self._buckets: dict[str, _Bucket] = {}

    def check(self, key: str) -> tuple[bool, int]:
        """Consume one token for `key`. Returns (allowed, retry_after_seconds)."""
        now = self.clock()
        b = self._buckets.get(key)
        if b is None:
            if len(self._buckets) >= self.max_keys:
                self.evict(idle_s=0)
            b = _Bucket(tokens=float(self.burst), updated=now)
            self._buckets[key] = b
        else:
            b.tokens = min(float(self.burst), b.tokens + (now - b.updated) * self.rate)
            b.updated = now
        if b.tokens >= 1:
            b.tokens -= 1
            return True, 0
        retry = math.ceil((1 - b.tokens) / self.rate) if self.rate > 0 else 60
        return False, max(1, retry)

    def evict(self, idle_s: float = 3600) -> None:
        """Drop buckets that are full again (nothing to remember) and idle; if still over
        capacity, drop the least recently used ones."""
        now = self.clock()
        for key, b in list(self._buckets.items()):
            refilled = min(float(self.burst), b.tokens + (now - b.updated) * self.rate)
            if refilled >= self.burst and now - b.updated >= idle_s:
                del self._buckets[key]
        if len(self._buckets) >= self.max_keys:
            for key, _ in sorted(self._buckets.items(), key=lambda kv: kv[1].updated)[: len(self._buckets) // 2]:
                del self._buckets[key]

    def __len__(self) -> int:
        return len(self._buckets)


class RateLimits:
    """Minute and daily buckets per client; both must allow the request.
    Rebuilds its buckets whenever the settings change (tests monkeypatch them)."""

    def __init__(self, clock: Callable[[], float] = time.monotonic):
        self.clock = clock
        self._key: tuple | None = None
        self.minute: TokenBucketLimiter | None = None
        self.daily: TokenBucketLimiter | None = None

    def _sync(self) -> None:
        key = (settings.rate_per_minute, settings.rate_burst, settings.rate_daily)
        if key != self._key:
            self._key = key
            if settings.rate_per_minute > 0:
                self.minute = TokenBucketLimiter(settings.rate_per_minute / 60, settings.rate_burst, self.clock)
                self.daily = TokenBucketLimiter(settings.rate_daily / 86400, settings.rate_daily, self.clock)
            else:
                self.minute = self.daily = None

    @property
    def enabled(self) -> bool:
        self._sync()
        return self.minute is not None

    def check(self, key: str) -> int:
        """0 when allowed, otherwise the number of seconds to wait."""
        self._sync()
        if self.minute is None or self.daily is None:
            return 0
        ok_m, wait_m = self.minute.check(key)
        if not ok_m:
            return wait_m
        ok_d, wait_d = self.daily.check(key)
        if not ok_d:  # give the minute token back; the day is what blocks
            self.minute._buckets[key].tokens += 1
            return wait_d
        return 0

    def evict(self) -> None:
        self._sync()
        for lim in (self.minute, self.daily):
            if lim is not None:
                lim.evict()

    def reset(self) -> None:
        self._key = None
        self._sync()


def client_ip(scope: dict) -> str:
    """Best-effort client key. Behind the Fly proxy the socket peer is the proxy, so the
    real address comes from headers; those are trusted only when TRUSTED_PROXY is on."""
    ip = ""
    if settings.trusted_proxy:
        headers = {k.decode("latin-1").lower(): v.decode("latin-1") for k, v in scope.get("headers", [])}
        ip = headers.get("fly-client-ip", "") or headers.get("x-forwarded-for", "").split(",")[0].strip()
    if not ip:
        client = scope.get("client")
        ip = client[0] if client else "unknown"
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return ip
    if addr.version == 6:  # one subscriber usually owns a whole /64
        return str(ipaddress.ip_network(f"{addr}/64", strict=False))
    return str(addr)
