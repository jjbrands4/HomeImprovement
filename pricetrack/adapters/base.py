"""SourceAdapter interface + helpers shared by adapters.

    SourceAdapter.search(item, ctx) -> SourceResult(listings=[Listing ...], outcome=...)
    SourceAdapter.fetch(url, item, ctx) -> SourceResult

Adapters never raise: run_safely() converts any exception into a parser_failure outcome so one
broken source/parser can never stop other sources or items (Apprise-style plugin isolation)."""
from __future__ import annotations

import re
import traceback
from dataclasses import dataclass, field
from typing import Callable, Optional

from ..extract import PageOffer, ProductPage
from ..fetch import Fetcher
from ..models import Item, Listing, MARKET_SNAPSHOT, Outcome, SourceResult, utcnow
from ..text import extract_pack_qty
from ..urls import merchant_item_id, normalize_url


@dataclass
class AdapterContext:
    fetcher: Fetcher
    state: object                    # history.State
    primary: list = field(default_factory=list)
    log: Callable = print
    no_browser: bool = False
    force_discovery: bool = False


class SourceAdapter:
    name = "base"
    evidence = MARKET_SNAPSHOT

    def available(self) -> tuple:
        return True, ""

    def search(self, item: Item, ctx: AdapterContext, **kw) -> SourceResult:
        return SourceResult(self.name, item.product[:40], Outcome.SKIPPED, "search not supported")

    def fetch(self, url: str, item: Item, ctx: AdapterContext, **kw) -> SourceResult:
        return SourceResult(self.name, url, Outcome.SKIPPED, "fetch not supported")


def run_safely(fn, source: str, target: str, *a, **kw) -> SourceResult:
    try:
        r = fn(*a, **kw)
        return r if isinstance(r, SourceResult) else SourceResult(source, target, Outcome.PARSER, "adapter returned nothing")
    except Exception as e:                                    # isolate: report, never propagate
        tb = traceback.extract_tb(e.__traceback__)
        where = f" @ {tb[-1].name}:{tb[-1].lineno}" if tb else ""
        return SourceResult(source, target, Outcome.PARSER, f"{type(e).__name__}: {str(e)[:120]}{where}")


def offer_gtins(page: ProductPage, offer: PageOffer) -> set:
    if offer.gtin:
        return {offer.gtin}
    return set(page.gtins) if len(page.offers) == 1 or len(page.gtins) == 1 else set()


def page_listings(page: ProductPage, url: str, vendor: str, source: str, evidence: str,
                  want_variant: Optional[str] = None) -> list:
    """ProductPage -> one Listing per offer/variant (exact duplicates collapsed)."""
    out, seen = [], set()
    base_title = page.name or page.title
    offers = page.offers
    if want_variant:
        pick = [o for o in offers if want_variant in (o.variant_id or "") or want_variant in (o.url or "")
                or want_variant in (o.sku or "")]
        offers = pick or offers
    for o in offers:
        title = base_title
        if o.name and o.name.lower() not in (base_title or "").lower():
            title = f"{base_title} - {o.name}" if base_title else o.name
        if o.variant and o.variant.lower() not in title.lower():
            title = f"{title} - {o.variant}"
        key = (round(o.price, 2), o.variant, o.sku, o.condition, o.availability)
        if key in seen:
            continue
        seen.add(key)
        lurl = normalize_url(o.url) if o.url and o.url.startswith("http") else normalize_url(url)
        pack = o.pack_qty
        if pack is None and o.variant:
            q = extract_pack_qty(o.variant)
            pack = q if q > 1 else None
        out.append(Listing(
            title=title or url, url=lurl, price=o.price, vendor=vendor, source=source, evidence=evidence,
            condition=o.condition or "new", in_stock=o.in_stock, availability=o.availability or "",
            regular_price=o.regular_price, shipping=o.shipping, pack_qty_hint=pack, variant=o.variant,
            variant_id=o.variant_id or o.sku, merchant_item_id=o.sku or merchant_item_id(lurl),
            gtins=offer_gtins(page, o), mpns=set(page.mpns) | ({o.mpn} if o.mpn else set()), brand=page.brand,
            color=o.color or page.color, conditional=o.conditional, conditional_detail=o.conditional_detail,
            method=page.method, retrieved_at=utcnow()))
    return out


def tokens_for_query(text: str, limit: int = 6) -> list:
    return [t for t in re.split(r"[^A-Za-z0-9.]+", text or "") if t][:limit]
