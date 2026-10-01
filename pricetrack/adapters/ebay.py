"""OPTIONAL eBay Browse API (used / refurbished + seller ratings).

API access may be pending; with no keys (or a rejected keyset) this adapter reports api_unavailable
and every other source runs normally. Searches by GTIN when one is known (identifier-first)."""
from __future__ import annotations

import base64
from typing import Optional

import requests

from ..identity import normalize_gtin
from ..models import Item, Listing, Outcome, SourceResult, VERIFIED_DIRECT, utcnow
from ..text import detect_condition, parse_price
from ..urls import normalize_url
from .base import AdapterContext, SourceAdapter

EBAY_TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
EBAY_SEARCH_URL = "https://api.ebay.com/buy/browse/v1/item_summary/search"
EBAY_LIMIT = 40
EBAY_MIN_FEEDBACK_PCT = 95.0
EBAY_MIN_FEEDBACK_COUNT = 10
EBAY_CONDITION_IDS = "1000|1500|2000|2010|2020|2030|2500|2750|3000|4000"
HTTP_TIMEOUT = 15


class EbayClient:
    def __init__(self, client_id: Optional[str], client_secret: Optional[str]):
        self.cid, self.secret = client_id, client_secret
        self.disabled = not (client_id and client_secret)
        self.reason = "EBAY_CLIENT_ID/EBAY_CLIENT_SECRET not set (optional - API access pending)" if self.disabled else ""
        self.token: Optional[str] = None
        self.session = requests.Session()
        self.log = print

    def _disable(self, why: str) -> None:
        if not self.disabled:
            self.log(f"  [eBay] disabled for rest of run: {why}")
        self.disabled, self.reason = True, why

    def _get_token(self) -> bool:
        try:
            basic = base64.b64encode(f"{self.cid}:{self.secret}".encode()).decode()
            r = self.session.post(EBAY_TOKEN_URL, timeout=HTTP_TIMEOUT, headers={
                "Authorization": f"Basic {basic}", "Content-Type": "application/x-www-form-urlencoded"},
                data={"grant_type": "client_credentials", "scope": "https://api.ebay.com/oauth/api_scope"})
            if r.status_code == 200:
                self.token = r.json().get("access_token")
                return bool(self.token)
            self._disable(f"token request failed HTTP {r.status_code} (production keyset active yet?)")
        except Exception as e:
            self._disable(f"token request error {type(e).__name__}")
        return False

    def search(self, query: str, gtin: Optional[str] = None, _retry: bool = True) -> Optional[list]:
        """Raw itemSummaries ([] = none, None = unavailable)."""
        if self.disabled:
            return None
        if not self.token and not self._get_token():
            return None
        params = {"limit": EBAY_LIMIT,
                  "filter": f"buyingOptions:{{FIXED_PRICE}},conditionIds:{{{EBAY_CONDITION_IDS}}},priceCurrency:USD"}
        if gtin:
            params["gtin"] = gtin
        else:
            params["q"] = query
        try:
            r = self.session.get(EBAY_SEARCH_URL, timeout=HTTP_TIMEOUT, params=params, headers={
                "Authorization": f"Bearer {self.token}", "X-EBAY-C-MARKETPLACE-ID": "EBAY_US"})
            if r.status_code == 401 and _retry and self._get_token():
                return self.search(query, gtin, _retry=False)
            if r.status_code in (403, 429):
                self._disable(f"HTTP {r.status_code}")
                return None
            if r.status_code != 200:
                self.log(f"  [eBay] HTTP {r.status_code}")
                return None
            return r.json().get("itemSummaries") or []
        except Exception as e:
            self.log(f"  [eBay] request failed ({type(e).__name__}: {e})")
            return None


def parse_ebay(items: list) -> list:
    """eBay itemSummaries -> Listings, with shipping cost when eBay returns it."""
    out = []
    for it in items or []:
        p = it.get("price") or {}
        price = parse_price(p.get("value"))
        if not price or p.get("currency", "USD") != "USD" or not it.get("title"):
            continue
        s = it.get("seller") or {}
        pct, cnt = parse_price(s.get("feedbackPercentage")), s.get("feedbackScore")
        if pct is not None and isinstance(cnt, int):
            comment = f"eBay seller '{s.get('username', '?')}': {pct:g}% positive, {cnt:,} total feedback ratings"
            ok = pct >= EBAY_MIN_FEEDBACK_PCT and cnt >= EBAY_MIN_FEEDBACK_COUNT
        else:
            comment, ok = "N/A (seller feedback not returned)", False
        ship = None
        opts = it.get("shippingOptions") or []
        if opts and isinstance(opts[0], dict):
            sc = opts[0].get("shippingCost") or {}
            ship = parse_price(sc.get("value")) if sc else None
            if ship is None and str(opts[0].get("shippingCostType", "")).upper() == "FREE":
                ship = 0.0
        orig = ((it.get("marketingPrice") or {}).get("originalPrice") or {}).get("value")
        g = normalize_gtin(it.get("gtin"))
        out.append(Listing(
            title=it["title"], url=normalize_url(it.get("itemWebUrl", "")), price=price, vendor="eBay", source="ebay",
            evidence=VERIFIED_DIRECT, condition=detect_condition(it["title"], it.get("condition", "")),
            in_stock=True, availability="in_stock", seller_comment=comment, seller_ok=ok, shipping=ship,
            regular_price=parse_price(orig) if parse_price(orig) and parse_price(orig) > price else None,
            merchant_item_id=str(it.get("itemId") or it.get("legacyItemId") or ""), gtins={g} if g else set(),
            method="eBay Browse API", retrieved_at=utcnow()))
    return out


class EbayAdapter(SourceAdapter):
    name = "ebay"
    evidence = VERIFIED_DIRECT

    def __init__(self, client: EbayClient):
        self.client = client

    def available(self) -> tuple:
        return (not self.client.disabled), self.client.reason

    def search(self, item: Item, ctx: AdapterContext, query: str = "", **kw) -> SourceResult:
        if self.client.disabled:
            return SourceResult(self.name, "eBay", Outcome.API, self.client.reason)
        raw, how = None, "keywords"
        for g in sorted(item.all_gtins)[:1]:
            raw, how = self.client.search(query, gtin=g.lstrip("0").zfill(12)), "GTIN"
            if raw:
                break
        if not raw:
            raw, how = self.client.search(query), "keywords"
        if raw is None:
            return SourceResult(self.name, "eBay", Outcome.API, self.client.reason or "request failed")
        ls = parse_ebay(raw)
        return SourceResult(self.name, "eBay", Outcome.SUCCESS if ls else Outcome.NO_MATCH, f"search by {how}", ls)
