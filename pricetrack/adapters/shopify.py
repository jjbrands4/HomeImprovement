"""Shopify adapter - works for ANY Shopify storefront (detected, not hand-marked).

    /search/suggest.json        product search (no key, rarely blocked)
    /products/<handle>.js       every variant: id, option values, price + compare_at (cents), available,
                                sku, barcode (GTIN) -> one Listing per variant (1/2/4-pack, colour ...)
    /products/<handle>.json     fallback (dollar strings, no stock flag)
Detection result is cached per domain in the state file."""
from __future__ import annotations

import re
from typing import Optional
from urllib.parse import urljoin

from ..identity import classify, has_identifier_hit, normalize_gtin
from ..models import Item, Listing, Outcome, SourceResult, VERIFIED_DISCOVERED, utcnow
from ..text import extract_pack_qty, parse_price
from ..urls import normalize_url
from .base import AdapterContext, SourceAdapter

SHOPIFY_MAX_PRODUCTS = 3
_PACK_OPT = re.compile(r"pack|quantity|qty|count|bundle|units?|size", re.I)
_QTY_OPT = re.compile(r"pack|quantity|qty|count|bundle|units?", re.I)       # bare numbers mean 'how many'

_COLOR_OPT = re.compile(r"colou?r|finish", re.I)


def _cents(v) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return round(v / 100.0, 2)
    return parse_price(v)


class ShopifyAdapter(SourceAdapter):
    name = "shopify"
    evidence = VERIFIED_DISCOVERED

    # ---- detection ------------------------------------------------------------------------------
    def is_shopify(self, domain: str, ctx: AdapterContext) -> Optional[bool]:
        known = ctx.state.shopify_get(domain)
        if known is not None:
            return known
        base = f"https://{domain}"
        r = ctx.fetcher.get(f"{base}/products.json", params={"limit": 1}, kind="json", json_body=True, retries=0,
                            allow_cffi=False)
        verdict = None
        if r.ok and isinstance(r.data, dict) and isinstance(r.data.get("products"), list):
            verdict = True
        elif r.ok or r.outcome == Outcome.NO_MATCH:
            r2 = ctx.fetcher.get(base + "/", retries=0)
            verdict = bool(r2.ok and re.search(r"cdn\.shopify\.com|Shopify\.theme|shopify-section", r2.text or ""))
        if verdict is not None:                 # blocked/network -> undecided, ask again next run
            ctx.state.shopify_put(domain, verdict)
        return verdict

    # ---- product variants -----------------------------------------------------------------------
    def product(self, base: str, handle: str, item: Item, ctx: AdapterContext, vendor: str = "",
                evidence: str = VERIFIED_DISCOVERED, variant: Optional[str] = None) -> SourceResult:
        f = ctx.fetcher
        url = f"{base}/products/{handle}.js"
        r = f.get(url, kind="json", json_body=True, conditional=True)
        data, how = None, "Shopify .js"
        if r.not_modified and f.cached(url):
            data, how = f.cached(url)["parsed"], "Shopify .js (HTTP 304 cache)"
        elif r.ok and isinstance(r.data, dict):
            data = r.data
            f.store(url, r, data)
        if not data:
            r2 = f.get(f"{base}/products/{handle}.json", kind="json", json_body=True)
            if r2.ok and isinstance(r2.data, dict) and isinstance(r2.data.get("product"), dict):
                data, how = r2.data["product"], "Shopify .json"
            else:
                return SourceResult(self.name, handle, r.outcome if not r.ok else Outcome.PARSER,
                                    r.detail or "not a Shopify product")
        ls = self._variants(data, base, handle, vendor, evidence, how)
        if variant:
            pick = [l for l in ls if l.variant_id == str(variant)]
            ls = pick or ls
        if not ls:
            return SourceResult(self.name, handle, Outcome.PARSER, "Shopify product has no priced variants")
        f.remember(base.split("//")[-1], extractor=how)
        oos = all(l.in_stock is False for l in ls)
        return SourceResult(self.name, handle, Outcome.UNAVAILABLE if oos else Outcome.SUCCESS, how, ls)

    @staticmethod
    def _variants(data: dict, base: str, handle: str, vendor: str, evidence: str, how: str) -> list:
        title = data.get("title") or handle
        brand = data.get("vendor") or ""
        opts = data.get("options") or []
        names = [(o.get("name") if isinstance(o, dict) else str(o)) or "" for o in opts]
        variants = data.get("variants") or []
        out = []
        for v in variants:
            cents = how.endswith(".js") or "(HTTP 304" in how
            price = _cents(v.get("price")) if cents else parse_price(v.get("price"))
            if not price:
                continue
            cmp_ = _cents(v.get("compare_at_price")) if cents else parse_price(v.get("compare_at_price"))
            values = [v.get(f"option{i}") for i in (1, 2, 3)]
            vt = str(v.get("title") or "").strip()
            pack, color = None, ""
            for n, val in zip(names, values):
                if not val:
                    continue
                if _PACK_OPT.search(n):
                    sval = str(val)
                    pack = int(sval) if (sval.strip().isdigit() and _QTY_OPT.search(n)) else (extract_pack_qty(sval) if extract_pack_qty(sval) > 1
                                                                     else (1 if re.search(r"\b(?:1|one|single)\b", sval, re.I) else None))
                if _COLOR_OPT.search(n):
                    color = str(val)
            if pack is None:
                q = extract_pack_qty(vt)
                pack = q if q > 1 else None
            full = title if vt.lower() in ("", "default title") else f"{title} - {vt}"
            vid = str(v.get("id") or "")
            url = normalize_url(f"{base}/products/{handle}" + (f"?variant={vid}" if vid and len(variants) > 1 else ""))
            avail = v.get("available")
            g = normalize_gtin(v.get("barcode"))
            out.append(Listing(
                title=full, url=url, price=price, vendor=vendor or base.split("//")[-1], source="shopify",
                evidence=evidence, in_stock=avail if isinstance(avail, bool) else None,
                availability=("in_stock" if avail else "out_of_stock") if isinstance(avail, bool) else "",
                regular_price=cmp_ if cmp_ and cmp_ > price else None, pack_qty_hint=pack,
                variant=vt if vt.lower() != "default title" else "", variant_id=vid, merchant_item_id="v" + vid if vid else "",
                gtins={g} if g else set(), mpns={v["sku"]} if v.get("sku") and len(str(v["sku"])) >= 4 else set(),
                brand=brand, color=color, method=how, retrieved_at=utcnow()))
        return out

    # ---- search ---------------------------------------------------------------------------------
    def search(self, item: Item, ctx: AdapterContext, domain: str = "", vendor: str = "",
               evidence: str = VERIFIED_DISCOVERED, **kw) -> SourceResult:
        base = f"https://{domain}"
        queries = []
        for g in sorted(item.all_gtins)[:1]:
            queries.append(g.lstrip("0").zfill(12))
        for m in item.all_mpns[:1]:
            queries.append(m)
        queries.append(item.product)
        prods, seen, last = [], set(), None
        for q in queries:
            r = ctx.fetcher.get(f"{base}/search/suggest.json", kind="json", json_body=True, retries=1,
                                params={"q": q, "resources[type]": "product", "resources[limit]": 10})
            last = r
            if not r.ok or not isinstance(r.data, dict):
                continue
            for p in ((r.data.get("resources") or {}).get("results") or {}).get("products") or []:
                h = p.get("handle") or (re.search(r"/products/([^/?#]+)", p.get("url") or "") or [None, None])[1]
                if h and h not in seen:
                    seen.add(h)
                    prods.append((p, h))
            if prods and q != item.product:
                break                                  # an identifier query hit: no need to search by name
        if not prods:
            if last is not None and not last.ok:
                return SourceResult(self.name, domain, last.outcome, last.detail)
            return SourceResult(self.name, domain, Outcome.NO_MATCH, "Shopify search: no products")
        matched = [(p, h) for p, h in prods
                   if classify(item, p.get("title", "")).confidence in ("High", "Medium")
                   or has_identifier_hit(item, p.get("title", ""))]
        out, notes = [], []
        for p, h in matched[:SHOPIFY_MAX_PRODUCTS]:
            r = self.product(base, h, item, ctx, vendor=vendor, evidence=evidence)
            if r.listings:
                out += r.listings
            else:
                notes.append(f"{h}: {r.outcome}")
                price = parse_price(p.get("price"))
                if price:
                    out.append(Listing(title=p.get("title", ""), url=normalize_url(urljoin(base, (p.get("url") or "").split("?")[0])),
                                       price=price, vendor=vendor or domain, source="shopify", evidence=evidence,
                                       in_stock=p.get("available") if isinstance(p.get("available"), bool) else None,
                                       method="Shopify search result", retrieved_at=utcnow()))
        detail = f"Shopify search {len(prods)} results, {len(matched)} matching product(s), {len(out)} priced option(s)"
        if notes:
            detail += " (" + "; ".join(notes[:2]) + ")"
        return SourceResult(self.name, domain, Outcome.SUCCESS if out else Outcome.NO_MATCH, detail, out)
