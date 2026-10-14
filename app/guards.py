"""Pure ASGI middlewares that run before the upload body is parsed.

Order (outermost first): SecurityHeadersMiddleware -> RateLimitMiddleware -> BodyLimitMiddleware -> app.
"""

from __future__ import annotations

import asyncio
import json
import time
from typing import Callable

from app.config import settings
from app.ratelimit import RateLimits, client_ip

CONVERT_PATH = "/api/convert"


def _json_bytes(payload: dict) -> bytes:
    return json.dumps(payload).encode()


async def _send_json(send, status: int, payload: dict, extra_headers: list[tuple[bytes, bytes]] | None = None) -> None:
    body = _json_bytes(payload)
    headers = [(b"content-type", b"application/json"), (b"content-length", str(len(body)).encode())]
    headers += extra_headers or []
    await send({"type": "http.response.start", "status": status, "headers": headers})
    await send({"type": "http.response.body", "body": body})


def _is_convert(scope: dict) -> bool:
    return scope["type"] == "http" and scope["method"] == "POST" and scope["path"] == CONVERT_PATH


class RateLimitMiddleware:
    """429 when the client IP is over its token buckets or already has a conversion in progress, 503 when
    the conversion queue is full. All are decided before a single byte of the upload is read."""

    def __init__(self, app, limits: RateLimits, queue_full: Callable[[], bool]):
        self.app = app
        self.limits = limits
        self.queue_full = queue_full

    async def __call__(self, scope, receive, send):
        if not _is_convert(scope):
            await self.app(scope, receive, send)
            return
        if self.queue_full():
            await _send_json(
                send,
                503,
                {"error": "busy", "detail": "The server is busy. Try again in a moment."},
                [(b"retry-after", b"15")],
            )
            return
        key = client_ip(scope)
        # Checked before the token buckets, so a request turned away here costs no token.
        if not self.limits.begin(key):
            await _send_json(
                send,
                429,
                {"error": "rate_limited", "detail": "A conversion from your address is still running. Try again soon."},
                [(b"retry-after", b"15")],
            )
            return
        try:
            wait = self.limits.check(key)
            if wait:
                await _send_json(
                    send,
                    429,
                    {"error": "rate_limited", "detail": f"Too many conversions. Try again in {wait} s."},
                    [(b"retry-after", str(wait).encode())],
                )
                return
            await self.app(scope, receive, send)
        finally:
            self.limits.end(key)


class BodyLimitMiddleware:
    """Rejects oversize and overlong uploads early.

    1. Content-Length over the cap: answer 413 at once, nothing is read.
    2. Chunked or lying clients: count the bytes handed to the app; once over the cap, report a
       disconnect to the app (so it stops reading) and rewrite whatever it answers into a 413.
    3. A body still incomplete after `upload_timeout_s`: same trick, answered with a 408. The Fly proxy
       streams the body through, so a client sending a byte every few seconds would otherwise hold a
       request slot (and its address's one in-progress conversion) for as long as it likes.
    """

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if not _is_convert(scope):
            await self.app(scope, receive, send)
            return
        cap = settings.body_limit_bytes
        too_large = {
            "error": "too_large",
            "detail": f"File is larger than {settings.max_upload_bytes // (1024 * 1024)} MB.",
        }

        headers = {k.lower(): v for k, v in scope.get("headers", [])}
        length = headers.get(b"content-length")
        if length is not None and length.isdigit() and int(length) > cap:
            await _send_json(send, 413, too_large)
            return

        timeout = settings.upload_timeout_s
        deadline = time.monotonic() + timeout
        seen = 0
        body_done = False
        rejected: tuple[int, dict] | None = None

        async def limited_receive():
            nonlocal seen, body_done, rejected
            if rejected:
                return {"type": "http.disconnect"}
            if body_done or timeout <= 0:
                message = await receive()
            else:
                try:
                    message = await asyncio.wait_for(receive(), max(deadline - time.monotonic(), 0))
                except TimeoutError:
                    detail = f"The upload took longer than {int(timeout)} s. Try again on a faster connection."
                    rejected = (408, {"error": "upload_timeout", "detail": detail})
                    return {"type": "http.disconnect"}
            if message["type"] == "http.request":
                seen += len(message.get("body", b""))
                if seen > cap:
                    rejected = (413, too_large)
                    return {"type": "http.disconnect"}
                body_done = not message.get("more_body", False)
            return message

        async def rewriting_send(message):
            if rejected:
                if message["type"] == "http.response.start":
                    await _send_json(send, *rejected)
                return  # drop the app's own answer to the rejected request
            await send(message)

        await self.app(scope, limited_receive, rewriting_send)


SECURITY_HEADERS = {
    "x-content-type-options": "nosniff",
    "referrer-policy": "no-referrer",
    "x-frame-options": "DENY",
    "content-security-policy": (
        "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; "
        "img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; form-action 'self'; base-uri 'self'"
    ),
    # Browsers ignore HSTS on plain HTTP, so the local http://127.0.0.1 run is unaffected.
    "strict-transport-security": "max-age=31536000",
    "cross-origin-opener-policy": "same-origin",
    "cross-origin-resource-policy": "same-origin",
}


class SecurityHeadersMiddleware:
    """Outermost, so the 413/429/503 answers of the other guards carry the headers too. Starlette's
    ServerErrorMiddleware sits outside every user middleware, so the 500 handler adds them itself."""

    HEADERS = [(k.encode(), v.encode()) for k, v in SECURITY_HEADERS.items()]

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        async def send_with_headers(message):
            if message["type"] == "http.response.start":
                message["headers"] = list(message.get("headers", [])) + self.HEADERS
            await send(message)

        await self.app(scope, receive, send_with_headers)
