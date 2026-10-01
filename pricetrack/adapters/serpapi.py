"""Optional SerpApi (Google Shopping) - DISCOVERY / CORROBORATION layer only.

Rows are market snapshots (evidence = market_snapshot): Google's merchant feeds, not a page we fetched.
They never establish a verified baseline on their own. Responses are cached (see State.serp_*) so a
retried / repeated run inside SERP_CACHE_HOURS spends no credits, and the search is skipped entirely
when enough verified merchant pages already priced the item (SERP_REDISCOVER_DAYS)."""
from __future__ import annotations

import re
from typing import Optional

import requests

from ..models import Item, Listing, MARKET_SNAPSHOT, Outcome, SourceResult, utcnow
from ..text import detect_condition, norm_text, parse_price
from ..urls import host_of, normalize_url
from .base import AdapterContext, SourceAdapter

SERPAPI_SEARCH_URL = "https://serpapi.com/search.json"
SERPAPI_ACCOUNT_URL = "https://serpapi.com/account.json"
SERPAPI_RESERVE = 3
HTTP_TIMEOUT = 15

_CONDITIONAL_PATTERNS = [
    ("coupon", re.compile(r"coupon|promo code|with code|clip", re.I)),
    ("membership", re.compile(r"member|prime\b|club price|plus price|my best buy|walmart\+|circle", re.I)),
    ("subscription", re.compile(r"subscribe|subscription|auto[- ]?ship", re.I)),
    ("financing", re.compile(r"/\s*mo\b|per month|monthly|apr\b|affirm|klarna|afterpay|installment", re.I)),
    ("trade_in", re.compile(r"trade[- ]?in", re.I)),
    ("carrier_activation", re.compile(r"activation|with carrier|unlock", re.I)),
    ("in_cart", re.compile(r"in cart|add to cart to see", re.I)),
]


def conditional_kind(*texts) -> tuple:
    blob = " | ".join(str(t) for t in texts if t)
    for kind, rx in _CONDITIONAL_PATTERNS:
        m = rx.search(blob)
        if m:
            return kind, blob[max(0, m.start() - 30):m.end() + 30].strip()
    return "", ""


def shipping_of(text: str) -> Optional[float]:
    t = (text or "").lower()
    if not t:
        return None
    if re.search(r"free (?:delivery|shipping)|free\s*$|\$0(?:\.00)? (?:delivery|shipping)", t):
        return 0.0
    m = re.search(r"\$\s*([\d,]+(?:\.\d{2})?)\s*(?:delivery|shipping)", t) or re.search(r"(?:shipping|delivery)\s*\$\s*([\d,]+(?:\.\d{2})?)", t)
    return parse_price(m.group(1)) if m else None


class SerpApiClient:
    """Thin SerpApi wrapper that degrades gracefully (never raises)."""

    def __init__(self, key: Optional[str]):
        self.key = key
        self.disabled = not key
        self.reason = "SERPAPI_KEY not set" if not key else ""
        self.credits_left: Optional[float] = None
        self.calls = 0
        self.session = requests.Session()
        self.log = print

    def _disable(self, why: str) -> None:
        if not self.disabled:
            self.log(f"  [SerpApi] disabled for rest of run: {why}")
        self.disabled, self.reason = True, why

    def check_credits(self) -> None:
        if self.disabled:
            return
        try:
            r = self.session.get(SERPAPI_ACCOUNT_URL, params={"api_key": self.key}, timeout=HTTP_TIMEOUT)
            if r.status_code in (401, 403):
                self._disable("invalid API key")
            elif r.ok:
                d = r.json()
                left = d.get("total_searches_left", d.get("plan_searches_left"))
                if isinstance(left, (int, float)):
                    self.credits_left = left
                    self.log(f"  [SerpApi] credits left this month: {int(left)}")
                    if left <= SERPAPI_RESERVE:
                        self._disable(f"only {int(left)} credits left (reserve={SERPAPI_RESERVE})")
        except Exception as e:
            self.log(f"  [SerpApi] account check skipped ({type(e).__name__})")

    def shopping(self, query: str) -> Optional[list]:
        """Raw shopping_results ([] = no results, None = source unavailable)."""
        if self.disabled:
            return None
        try:
            r = self.session.get(SERPAPI_SEARCH_URL, timeout=HTTP_TIMEOUT * 2, params={
                "engine": "google_shopping", "q": query, "gl": "us", "hl": "en",
                "google_domain": "google.com", "api_key": self.key})
            self.calls += 1
            if r.status_code in (401, 403):
                self._disable("invalid API key")
                return None
            if r.status_code == 429:
                self._disable("rate limited / out of searches")
                return None
            data = r.json()
            err = str(data.get("error", "")).lower()
            if err:
                if "hasn't returned any results" in err or "no results" in err:
                    return []
                if any(w in err for w in ("run out", "out of searches", "limit", "exceeded", "plan", "invalid api key")):
                    self._disable(data.get("error", "quota error"))
                else:
                    self.log(f"  [SerpApi] error: {data.get('error')}")
                return None
            if self.credits_left is not None:
                self.credits_left -= 1
                if self.credits_left <= SERPAPI_RESERVE:
                    self._disable("credit reserve reached")
            return data.get("shopping_results") or []
        except Exception as e:
            self.log(f"  [SerpApi] request failed ({type(e).__name__}: {e})")
            return None


def parse_serpapi(results: list) -> list:
    """Google Shopping rows -> market_snapshot Listings (regular/old price, shipping, conditional
    pricing and a direct merchant link are kept when Google provides them)."""
    out = []
    for r in results or []:
        title = r.get("title") or ""
        price = r.get("extracted_price")
        if price is None:
            price = parse_price(r.get("price"))
        if not title or not price or price <= 0:
            continue
        vendor = re.sub(r"^from\s+", "", r.get("source") or "", flags=re.I).strip()
        comment = ""
        if r.get("rating"):
            comment = (f"Google Shopping product rating {r['rating']}/5"
                       + (f" ({r['reviews']:,} reviews)" if isinstance(r.get("reviews"), int) else "")
                       + "; product-level, not seller-level")
        old = r.get("extracted_old_price") or parse_price(r.get("old_price"))
        ext = r.get("extensions") or []
        kind, detail = conditional_kind(r.get("price") if isinstance(r.get("price"), str) and "/mo" in r.get("price", "") else "",
                                        " ".join(map(str, ext)) if isinstance(ext, list) else ext, r.get("tag"),
                                        r.get("badge"), r.get("snippet") if "coupon" in str(r.get("snippet", "")).lower() else "")
        link = r.get("link") or ""
        direct = link if link.startswith("http") and "google." not in host_of(link) else ""
        google = r.get("product_link") or (link if not direct else "")
        out.append(Listing(
            title=title, url=normalize_url(direct) if direct else google, price=float(price), vendor=vendor,
            source="serpapi", evidence=MARKET_SNAPSHOT,
            condition=detect_condition(title, r.get("second_hand_condition") or ""),
            regular_price=old if old and old > price else None, shipping=shipping_of(r.get("delivery") or ""),
            merchant_item_id=str(r.get("product_id") or ""), conditional=kind, conditional_detail=detail[:120],
            seller_comment=comment, method="Google Shopping (SerpApi)", alt_url=google if direct else "",
            retrieved_at=utcnow()))
    return out


def build_query(item: Item, include_specs: bool = True) -> str:
    """Product name + SKU/MPN (+ specs), skipping words already in the name."""
    parts, have = [item.product], set(norm_text(item.product).split())
    extras = list(item.all_mpns) + (list(item.spec_phrases) if include_specs else [])
    for e in extras:
        toks = norm_text(e).split()
        if toks and not all(t in have for t in toks):
            parts.append(e)
            have.update(toks)
    return " ".join(parts)[:150]


class SerpApiAdapter(SourceAdapter):
    name = "serpapi"
    evidence = MARKET_SNAPSHOT

    def __init__(self, client: SerpApiClient):
        self.client = client

    def available(self) -> tuple:
        return (not self.client.disabled), self.client.reason

    def search(self, item: Item, ctx: AdapterContext, include_specs: bool = True, cache_hours: float = 20,
               cache_only: bool = False, **kw) -> SourceResult:
        """cache_only=True: reuse a cached response if one exists (0 credits), never query."""
        q1 = build_query(item, include_specs)
        rows, notes = [], []
        for q in [q1] + ([build_query(item, False)] if build_query(item, False) != q1 else []):
            cached = ctx.state.serp_get(q, cache_hours)
            if cached is not None:
                rows += cached
                notes.append("cached response")
            elif cache_only:
                if not rows:
                    return SourceResult(self.name, "Google Shopping", Outcome.SKIPPED, "no cached response")
                break
            else:
                if self.client.disabled:
                    if not rows:
                        return SourceResult(self.name, "Google Shopping", Outcome.API, self.client.reason)
                    break
                raw = self.client.shopping(q)
                if raw is None:
                    if not rows:
                        return SourceResult(self.name, "Google Shopping", Outcome.API if self.client.disabled else Outcome.NETWORK,
                                            self.client.reason or "request failed")
                    break
                ctx.state.serp_put(q, raw)
                rows += raw
            if len(parse_serpapi(rows)) >= 3:
                break                       # specs may over-narrow: second (broader) query only if needed
        ls = parse_serpapi(rows)
        if not cache_only:
            ctx.state.serp_mark(item)
        return SourceResult(self.name, "Google Shopping", Outcome.SUCCESS if ls else Outcome.NO_MATCH,
                            ("; ".join(dict.fromkeys(notes))) if notes else "", ls)
