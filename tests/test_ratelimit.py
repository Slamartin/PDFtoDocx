from app.config import settings
from app.ratelimit import RateLimits, TokenBucketLimiter, client_ip


class Clock:
    def __init__(self):
        self.t = 1000.0

    def __call__(self):
        return self.t


def test_bucket_burst_and_refill():
    clock = Clock()
    lim = TokenBucketLimiter(rate_per_s=1 / 30, burst=2, clock=clock)
    assert lim.check("a") == (True, 0)
    assert lim.check("a") == (True, 0)
    ok, wait = lim.check("a")
    assert not ok and wait == 30
    clock.t += 30
    assert lim.check("a") == (True, 0)
    assert lim.check("b") == (True, 0)  # other keys are independent


def test_bucket_eviction():
    clock = Clock()
    lim = TokenBucketLimiter(rate_per_s=1, burst=1, clock=clock, max_keys=4)
    for k in "abcd":
        lim.check(k)
    clock.t += 10
    lim.evict(idle_s=5)  # all refilled and idle
    assert len(lim) == 0
    for k in "abcd":
        lim.check(k)
    lim.check("e")  # over capacity: idle eviction with idle_s=0 clears the refilled ones
    assert len(lim) <= 4


def test_rate_limits_follow_settings(monkeypatch):
    clock = Clock()
    rl = RateLimits(clock=clock)
    monkeypatch.setattr(settings, "rate_per_minute", 0)
    assert not rl.enabled and rl.check("x") == 0
    monkeypatch.setattr(settings, "rate_per_minute", 60)
    monkeypatch.setattr(settings, "rate_burst", 1)
    monkeypatch.setattr(settings, "rate_daily", 2)
    assert rl.enabled
    assert rl.check("x") == 0
    assert rl.check("x") == 1  # minute bucket empty, refills one per second
    clock.t += 1
    assert rl.check("x") == 0
    clock.t += 1
    assert rl.check("x") >= 3600  # daily bucket (2) is exhausted; minute token was refunded
    assert rl.minute._buckets["x"].tokens >= 1


def test_client_ip_trust(monkeypatch):
    scope = {
        "client": ("10.0.0.1", 1234),
        "headers": [(b"fly-client-ip", b"203.0.113.9"), (b"x-forwarded-for", b"198.51.100.1, 10.0.0.1")],
    }
    monkeypatch.setattr(settings, "trusted_proxy", False)
    assert client_ip(scope) == "10.0.0.1"
    monkeypatch.setattr(settings, "trusted_proxy", True)
    assert client_ip(scope) == "203.0.113.9"
    # X-Forwarded-For is client controlled (its first entry is spoofable), so it never picks the key.
    scope["headers"] = [(b"x-forwarded-for", b"198.51.100.1, 10.0.0.1")]
    assert client_ip(scope) == "10.0.0.1"
    scope["headers"] = [(b"fly-client-ip", b"2001:db8:abcd:1234:5555:6666:7777:8888")]
    assert client_ip(scope) == "2001:db8:abcd:1234::/64"
    scope["headers"] = [(b"fly-client-ip", b"not-an-ip")]
    assert client_ip(scope) == "not-an-ip"
