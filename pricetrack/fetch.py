"""
Resilient fetching.

  * retries with exponential backoff + jitter (Retry-After honoured) for timeouts / 429 / 5xx
  * per-domain politeness: one request in flight per domain and a minimum interval between requests
  * circuit breaker: after N consecutive blocks/failures a domain is skipped for the rest of the run
    (blocks also persist a short cool-down in the state file so the next run does not hammer it)
  * conditional GET (ETag / Last-Modified) with the PARSED result cached in the state file
  * strategy memory: the client that worked last time for a domain (plain requests vs curl_cffi
    Chrome impersonation vs browser render) is tried first next time
  * every result carries an Outcome category instead of a bare failure
"""
from __future__ import annotations

import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from typing import Optional

import requests

from .models import Outcome, utcnow
from .urls import host_of
from .extract import looks_blocked

try:                                               # Real-Chrome TLS fingerprint client (optional)
    from curl_cffi import requests as cffi_requests
except Exception:                                  # pragma: no cover - not installed
    cffi_requests = None

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/128.0 Safari/537.36")
BROWSER_HEADERS = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9",
                   "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                   "Accept-Encoding": "gzip, deflate", "Connection": "keep-alive",
                   "Upgrade-Insecure-Requests": "1",
                   "Sec-Fetch-Dest": "document", "Sec-Fetch-Mode": "navigate", "Sec-Fetch-Site": "none",
                   "Sec-Fetch-User": "?1",
                   "sec-ch-ua": '"Chromium";v="128", "Google Chrome";v="128", "Not-A.Brand";v="99"',
                   "sec-ch-ua-mobile": "?0", "sec-ch-ua-platform": '"Windows"'}
JSON_HEADERS = {"User-Agent": USER_AGENT, "Accept": "application/json,text/javascript,*/*;q=0.5",
                "Accept-Language": "en-US,en;q=0.9"}

PAGE_TIMEOUT = (10, 30)
API_TIMEOUT = 20
MAX_RETRIES = 2                  # extra attempts after the first
BACKOFF_BASE, BACKOFF_CAP = 1.5, 20.0
DOMAIN_MIN_INTERVAL = 2.0        # seconds between requests to the same retail domain
API_MIN_INTERVAL = 0.25          # official APIs (Best Buy 5 req/s, eBay, SerpApi)
CIRCUIT_THRESHOLD = 3            # consecutive blocks/failures that open a domain's circuit
CIRCUIT_COOLDOWN_HOURS = 6       # persisted cool-down after a run ended with the circuit open on blocks
HTTP_CACHE_MAX = 400             # cached parsed responses kept in the state file
BLOCK_STATUSES = {202, 401, 403, 429, 503}
API_HOSTS = ("api.bestbuy.com", "api.ebay.com", "serpapi.com")


@dataclass
class FetchResult:
    url: str
    status: int = 0
    text: str = ""
    content: bytes = b""
    headers: dict = field(default_factory=dict)
    final_url: str = ""
    outcome: str = Outcome.NETWORK
    detail: str = ""
    method: str = ""
    not_modified: bool = False
    data: object = None            # parsed JSON when requested

    @property
    def ok(self) -> bool:
        return self.outcome == Outcome.SUCCESS


class BrowserRenderer:
    """Playwright (Chromium) - LAST RESORT for JS-only product pages. All Playwright calls run on one
    dedicated thread (its sync API is not thread-safe). Heavy resources (images/fonts/media) are blocked."""

    def __init__(self, enabled: bool = True, timeout_ms: int = 30000):
        self.enabled = enabled
        self.timeout_ms = timeout_ms
        self.reason = ""
        self._pool = None
        self._pw = self._browser = self._ctx = None
        self.renders = 0
        if enabled:
            try:
                import playwright.sync_api  # noqa: F401
            except Exception:
                self.enabled, self.reason = False, "playwright not installed (pip install playwright && playwright install chromium)"

    def _start(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
        self._ctx = self._browser.new_context(user_agent=USER_AGENT, locale="en-US", viewport={"width": 1366, "height": 900})
        self._ctx.route("**/*", lambda r: r.abort() if r.request.resource_type in ("image", "media", "font")
                        else r.continue_())

    def _render(self, url: str):
        if self._ctx is None:
            self._start()
        page = self._ctx.new_page()
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
            try:
                page.wait_for_load_state("networkidle", timeout=min(12000, self.timeout_ms))
            except Exception:
                pass
            # Let client-side price widgets settle; JSON-LD injected by JS is now in the DOM.
            page.wait_for_timeout(1500)
            return page.content(), page.url, (resp.status if resp else 0)
        finally:
            page.close()

    def render(self, url: str):
        """Returns (html, final_url, status). Raises on failure."""
        if not self.enabled:
            raise RuntimeError(self.reason or "browser disabled")
        if self._pool is None:
            self._pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="browser")
        self.renders += 1
        return self._pool.submit(self._render, url).result(timeout=self.timeout_ms / 1000 + 30)

    def close(self):
        if self._pool is None:
            return

        def _stop():
            for obj, fn in ((self._ctx, "close"), (self._browser, "close"), (self._pw, "stop")):
                try:
                    if obj is not None:
                        getattr(obj, fn)()
                except Exception:
                    pass
        try:
            self._pool.submit(_stop).result(timeout=30)
        except Exception:
            pass
        self._pool.shutdown(wait=False)
        self._pool = None


class Fetcher:
    def __init__(self, state: Optional[dict] = None, session=None, browser: Optional[BrowserRenderer] = None,
                 sleep=None, use_cffi: bool = True, max_retries: int = MAX_RETRIES,
                 min_interval: float = DOMAIN_MIN_INTERVAL):
        self.state = state if state is not None else {}
        for k in ("http_cache", "strategies", "circuits"):
            self.state.setdefault(k, {})
        self.session = session or requests.Session()
        self.browser = browser
        self.sleep = sleep or (lambda s: time.sleep(s))
        self.use_cffi = use_cffi and cffi_requests is not None
        self.max_retries = max_retries
        self.min_interval = min_interval
        self._locks: dict = {}
        self._last: dict = {}
        self._fails: dict = {}
        self._open: dict = {}
        self._guard = threading.Lock()
        self.stats = {"requests": 0, "cache_hits": 0, "renders": 0, "blocked": 0}

    # ---- politeness ------------------------------------------------------------------------------
    def _lock(self, dom: str) -> threading.Lock:
        with self._guard:
            return self._locks.setdefault(dom, threading.Lock())

    def _wait_turn(self, dom: str, url: str):
        gap = API_MIN_INTERVAL if any(h in url for h in API_HOSTS) else self.min_interval
        last = self._last.get(dom)
        if last is not None:
            wait = gap - (time.monotonic() - last)
            if wait > 0:
                self.sleep(wait)
        self._last[dom] = time.monotonic()

    # ---- circuit breaker -------------------------------------------------------------------------
    def circuit_open(self, dom: str) -> Optional[str]:
        if dom in self._open:
            return self._open[dom]
        c = self.state["circuits"].get(dom)
        if c:
            try:
                until = datetime.fromisoformat(c["until"])
            except (KeyError, ValueError):
                until = None
            if until and until > utcnow():
                return f"cool-down until {until:%Y-%m-%d %H:%M} UTC after repeated blocks ({c.get('why', '')})"
            self.state["circuits"].pop(dom, None)
        return None

    def _record(self, dom: str, outcome: str, why: str):
        if outcome == Outcome.SUCCESS:
            self._fails[dom] = 0
            return
        if outcome in (Outcome.BLOCKED, Outcome.NETWORK):
            n = self._fails.get(dom, 0) + 1
            self._fails[dom] = n
            if outcome == Outcome.BLOCKED:
                self.stats["blocked"] += 1
            if n >= CIRCUIT_THRESHOLD and dom not in self._open:
                self._open[dom] = f"circuit open after {n} consecutive failures ({why})"
                if outcome == Outcome.BLOCKED:
                    self.state["circuits"][dom] = {
                        "until": (utcnow() + timedelta(hours=CIRCUIT_COOLDOWN_HOURS)).isoformat(), "why": why[:80]}

    # ---- cache -----------------------------------------------------------------------------------
    def cached(self, url: str) -> Optional[dict]:
        return self.state["http_cache"].get(url)

    def store(self, url: str, res: FetchResult, parsed) -> None:
        """Remember validators + the parsed result (never the raw body) for conditional GETs."""
        et, lm = res.headers.get("ETag") or res.headers.get("etag"), \
            res.headers.get("Last-Modified") or res.headers.get("last-modified")
        if not (et or lm):
            return
        cache = self.state["http_cache"]
        cache[url] = {"etag": et, "last_modified": lm, "parsed": parsed, "ts": utcnow().isoformat()}
        if len(cache) > HTTP_CACHE_MAX:
            for k in sorted(cache, key=lambda k: cache[k].get("ts", ""))[:len(cache) - HTTP_CACHE_MAX]:
                cache.pop(k, None)

    # ---- strategy memory -------------------------------------------------------------------------
    def strategy(self, dom: str) -> dict:
        return self.state["strategies"].setdefault(dom, {})

    def remember(self, dom: str, **kw):
        s = self.strategy(dom)
        s.update({k: v for k, v in kw.items() if v})
        s["ts"] = utcnow().isoformat()

    # ---- classification --------------------------------------------------------------------------
    @staticmethod
    def classify(status: int, text: str) -> tuple:
        if status == 304:
            return Outcome.SUCCESS, "not modified"
        if 200 <= status < 300 and status != 202:
            if looks_blocked(text) and len(text) < 60000:
                return Outcome.BLOCKED, "bot challenge page"
            return Outcome.SUCCESS, ""
        if status in (404, 410):
            return Outcome.NO_MATCH, f"HTTP {status} (page gone - check the URL)"
        if status in BLOCK_STATUSES:
            return Outcome.BLOCKED, f"site blocked automated access (HTTP {status})"
        return Outcome.NETWORK, f"HTTP {status}"

    @staticmethod
    def _retry_after(headers: dict) -> Optional[float]:
        v = (headers or {}).get("Retry-After") or (headers or {}).get("retry-after")
        if not v:
            return None
        try:
            return min(30.0, float(v))
        except ValueError:
            try:
                return min(30.0, max(0.0, (parsedate_to_datetime(v).replace(tzinfo=None) - utcnow()).total_seconds()))
            except Exception:
                return None

    def _backoff(self, attempt: int, retry_after: Optional[float] = None):
        delay = retry_after if retry_after is not None else \
            min(BACKOFF_CAP, BACKOFF_BASE * (2 ** attempt)) * random.uniform(0.5, 1.0)
        self.sleep(delay)

    # ---- main entry ------------------------------------------------------------------------------
    def get(self, url: str, *, params: Optional[dict] = None, headers: Optional[dict] = None, kind: str = "page",
            conditional: bool = False, json_body: bool = False, retries: Optional[int] = None,
            allow_cffi: bool = True) -> FetchResult:
        """GET with politeness, retries, circuit breaking and (optionally) a conditional request.
        kind='page' sends browser headers; kind='api'/'json' sends JSON headers and never impersonates."""
        dom = host_of(url)
        res = FetchResult(url=url)
        why = self.circuit_open(dom)
        if why:
            res.outcome, res.detail = Outcome.BLOCKED, why
            return res
        hdrs = dict(BROWSER_HEADERS if kind == "page" else JSON_HEADERS)
        hdrs.update(headers or {})
        cache = self.cached(url) if conditional else None
        if cache:
            if cache.get("etag"):
                hdrs["If-None-Match"] = cache["etag"]
            if cache.get("last_modified"):
                hdrs["If-Modified-Since"] = cache["last_modified"]
        retries = self.max_retries if retries is None else retries
        timeout = PAGE_TIMEOUT if kind == "page" else API_TIMEOUT
        prefer_cffi = kind == "page" and self.use_cffi and allow_cffi and self.strategy(dom).get("client") == "cffi"

        with self._lock(dom):
            if not prefer_cffi:
                for attempt in range(retries + 1):
                    self._wait_turn(dom, url)
                    self.stats["requests"] += 1
                    try:
                        r = self.session.get(url, params=params, headers=hdrs, timeout=timeout, allow_redirects=True)
                    except requests.exceptions.Timeout as e:
                        res.outcome, res.detail = Outcome.NETWORK, f"timeout ({type(e).__name__})"
                        if attempt < retries:
                            self._backoff(attempt)
                        continue
                    except Exception as e:
                        res.outcome, res.detail = Outcome.NETWORK, f"network error ({type(e).__name__})"
                        if attempt < retries:
                            self._backoff(attempt)
                        continue
                    self._fill(res, r, "requests")
                    if res.ok:
                        break
                    if res.status in (401, 403) or res.outcome == Outcome.NO_MATCH:
                        break                                # same client won't do better; maybe cffi will
                    if res.status in (429, 503) or res.status >= 500 or res.status == 202:
                        if attempt < retries:
                            self._backoff(attempt, self._retry_after(res.headers))
                        continue
                    break
            if (not res.ok and res.outcome in (Outcome.BLOCKED, Outcome.NETWORK) and kind == "page"
                    and self.use_cffi and allow_cffi):
                self._wait_turn(dom, url)
                self.stats["requests"] += 1
                try:
                    r = cffi_requests.get(url, params=params, impersonate="chrome", timeout=PAGE_TIMEOUT[1],
                                          headers={k: v for k, v in hdrs.items() if k.lower().startswith(("if-", "accept-lang"))})
                    prev = res.detail
                    self._fill(res, r, "cffi")
                    if not res.ok and prev and prev not in res.detail:
                        res.detail = f"{prev}; Chrome impersonation: {res.detail}"
                    if res.ok:
                        self.remember(dom, client="cffi")
                except Exception as e:
                    res.detail = (res.detail + "; " if res.detail else "") + f"Chrome impersonation {type(e).__name__}"
            elif res.ok and kind == "page" and not prefer_cffi and self.strategy(dom).get("client") == "cffi":
                self.remember(dom, client="requests")
        if res.ok and json_body and not res.not_modified:
            try:
                res.data = (r.json() if hasattr(r, "json") else None)
            except Exception:
                res.outcome, res.detail = Outcome.PARSER, "response was not JSON"
        self._record(dom, res.outcome, res.detail)
        if res.not_modified:
            self.stats["cache_hits"] += 1
        return res

    def _fill(self, res: FetchResult, r, method: str):
        res.status = int(getattr(r, "status_code", 0) or 0)
        res.headers = dict(getattr(r, "headers", {}) or {})
        try:
            res.text = r.text if res.status != 304 else ""
        except Exception:
            res.text = ""
        res.content = getattr(r, "content", b"") or b""
        res.final_url = str(getattr(r, "url", "") or res.url)
        res.method = method
        res.outcome, res.detail = self.classify(res.status, res.text)
        res.not_modified = res.status == 304

    def render(self, url: str) -> FetchResult:
        """Browser render (last resort). Uses the same politeness / circuit rules."""
        dom = host_of(url)
        res = FetchResult(url=url, method="browser")
        if not self.browser or not self.browser.enabled:
            res.outcome, res.detail = Outcome.SKIPPED, (self.browser.reason if self.browser else "browser disabled")
            return res
        why = self.circuit_open(dom)
        if why:
            res.outcome, res.detail = Outcome.BLOCKED, why
            return res
        with self._lock(dom):
            self._wait_turn(dom, url)
            self.stats["renders"] += 1
            try:
                html, final, status = self.browser.render(url)
            except Exception as e:
                res.outcome, res.detail = Outcome.NETWORK, f"browser render failed ({type(e).__name__}: {str(e)[:80]})"
                self._record(dom, res.outcome, res.detail)
                return res
        res.text, res.final_url, res.status = html or "", final or url, status or 200
        res.outcome, res.detail = self.classify(res.status, res.text)
        self._record(dom, res.outcome, res.detail)
        return res


def is_bot_or_empty(text: str) -> bool:
    return not text or looks_blocked(text)


def strip_html(s: str) -> str:
    return re.sub(r"<[^>]+>", " ", s or "")
