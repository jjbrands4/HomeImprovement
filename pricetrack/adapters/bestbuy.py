"""Best Buy Products API (free key: developer.bestbuy.com -> BESTBUY_API_KEY).

fetch(url)      : SKU taken from the product URL
search(item)    : known Best Buy SKU (cached) -> UPC (from GTIN) -> model number (MPN) -> keyword search
Official API => evidence verified_direct. Never scraped, never bot-blocked."""
from __future__ import annotations

import os
import re
from typing import Optional

from ..identity import classify, has_identifier_hit, normalize_gtin
from ..models import Item, Listing, Outcome, SourceResult, VERIFIED_DIRECT, utcnow
from ..text import detect_condition, norm_text, parse_price
from ..urls import normalize_url
from .base import AdapterContext, SourceAdapter

API = "https://api.bestbuy.com/v1/products({q})"
SHOW = ("sku,name,upc,modelNumber,manufacturer,salePrice,regularPrice,onSale,onlineAvailability,orderable,"
        "url,condition,freeShipping,shippingCost,color,active")
STOP = {"the", "and", "with", "for", "smart", "new", "inch", "pack"}


def bestbuy_sku(url: str) -> Optional[str]:
    m = re.search(r"(?:skuId=|/sku/|/)(\d{7})(?:\.p\b|\b)", url or "")
    return m.group(1) if m else None


class BestBuyAdapter(SourceAdapter):
    name = "bestbuy_api"
    evidence = VERIFIED_DIRECT

    def __init__(self, key: Optional[str] = None):
        self.key = key if key is not None else os.getenv("BESTBUY_API_KEY")
        self.disabled_reason = "" if self.key else "BESTBUY_API_KEY not set (free key at developer.bestbuy.com)"

    def available(self) -> tuple:
        return (not self.disabled_reason), self.disabled_reason

    def _query(self, q: str, ctx: AdapterContext, page_size: int = 10) -> tuple:
        if self.disabled_reason:
            return None, SourceResult(self.name, q, Outcome.API, self.disabled_reason)
        r = ctx.fetcher.get(API.format(q=q), kind="api", json_body=True, retries=1,
                            params={"apiKey": self.key, "format": "json", "show": SHOW, "pageSize": page_size})
        if r.status in (401, 403) or (r.status == 400 and "key" in (r.text or "").lower()):
            self.disabled_reason = f"Best Buy API HTTP {r.status} (invalid key or over quota)"
            return None, SourceResult(self.name, q, Outcome.API, self.disabled_reason)
        if r.status == 429:
            return None, SourceResult(self.name, q, Outcome.API, "Best Buy API rate limited (HTTP 429)")
        if not r.ok or not isinstance(r.data, dict):
            return None, SourceResult(self.name, q, r.outcome if not r.ok else Outcome.PARSER, r.detail or "bad response")
        return r.data.get("products") or [], None

    def _listing(self, p: dict, vendor: str, url_hint: str = "") -> Optional[Listing]:
        price = parse_price(p.get("salePrice")) or parse_price(p.get("regularPrice"))
        if not price:
            return None
        reg = parse_price(p.get("regularPrice"))
        ship = 0.0 if p.get("freeShipping") else parse_price(p.get("shippingCost"))
        g = normalize_gtin(p.get("upc"))
        url = normalize_url(url_hint or p.get("url") or f"https://www.bestbuy.com/site/{p.get('sku')}.p?skuId={p.get('sku')}")
        avail = p.get("onlineAvailability")
        return Listing(
            title=p.get("name") or "", url=url, price=price, vendor=vendor or "Best Buy", source=self.name,
            evidence=VERIFIED_DIRECT, condition=detect_condition(p.get("name") or "", p.get("condition") or ""),
            in_stock=bool(avail) if avail is not None else None,
            availability=("in_stock" if avail else "out_of_stock") if avail is not None else "",
            regular_price=reg if reg and reg > price else None, shipping=ship,
            merchant_item_id=str(p.get("sku") or ""), gtins={g} if g else set(),
            mpns={p["modelNumber"]} if p.get("modelNumber") else set(), brand=p.get("manufacturer") or "",
            color=p.get("color") or "", method="Best Buy API", retrieved_at=utcnow())

    def fetch(self, url: str, item: Item, ctx: AdapterContext, vendor: str = "Best Buy", **kw) -> SourceResult:
        sku = bestbuy_sku(url)
        if not sku:
            return SourceResult(self.name, url[:60], Outcome.SKIPPED, "no SKU in URL")
        prods, err = self._query(f"sku={sku}", ctx, 1)
        if err:
            return err
        if not prods:
            return SourceResult(self.name, f"sku {sku}", Outcome.NO_MATCH, f"SKU {sku} not found (discontinued?)")
        l = self._listing(prods[0], vendor, url)
        if not l:
            return SourceResult(self.name, f"sku {sku}", Outcome.PARSER, "API returned no price")
        out = Outcome.SUCCESS if l.in_stock is not False else Outcome.UNAVAILABLE
        return SourceResult(self.name, f"sku {sku}", out, "Best Buy API", [l])

    def search(self, item: Item, ctx: AdapterContext, vendor: str = "Best Buy", **kw) -> SourceResult:
        """Identifier-first resolution; keyword search results must still pass product matching."""
        tried = []
        known = (ctx.state.discovery_get(item, "bestbuy.com") or {}).get("item_id")
        attempts = []
        if known:
            attempts.append(("cached SKU", f"sku={known}"))
        for g in sorted(item.all_gtins)[:2]:
            attempts.append(("UPC", f"upc={g.lstrip('0').zfill(12)}"))
        for m in item.all_mpns[:2]:
            if re.fullmatch(r"[A-Za-z0-9-]{4,}", m):
                attempts.append(("model", f"modelNumber={m}"))
        words = [w for w in norm_text(item.product).split() if w not in STOP and not w.replace(".", "").isdigit()][:4]
        if words:
            attempts.append(("keywords", "&".join(f"search={w}" for w in words)))
        for how, q in attempts:
            prods, err = self._query(q, ctx)
            if err:
                if err.outcome == Outcome.API:
                    return err
                tried.append(f"{how}: {err.outcome}")
                continue
            ls = [l for l in (self._listing(p, vendor) for p in prods) if l]
            if how == "keywords":
                ls = [l for l in ls if classify(item, l.title, gtins=l.gtins, mpns=l.mpns).confidence != "Low"
                      or has_identifier_hit(item, l.title, l.gtins, l.mpns)]
            if ls:
                best = ls[0]
                if how != "keywords" or classify(item, best.title, gtins=best.gtins, mpns=best.mpns).confidence == "High":
                    ctx.state.discovery_put(item, "bestbuy.com", url=best.url, item_id=best.merchant_item_id, ok=True)
                for l in ls:
                    l.method = f"Best Buy API ({how})"
                return SourceResult(self.name, how, Outcome.SUCCESS, f"resolved by {how}", ls[:5])
            tried.append(f"{how}: none")
        return SourceResult(self.name, "search", Outcome.NO_MATCH, "; ".join(tried) or "no identifiers or keywords")
