"""Generic exact product page adapter (structured data first; browser render only as last resort)."""
from __future__ import annotations

from urllib.parse import urlparse

from ..extract import ProductPage, extract_product, looks_js_rendered
from ..models import Item, Outcome, SourceResult, VERIFIED_DIRECT
from ..text import norm_text
from ..urls import host_of, looks_like_listing_page, normalize_url, shopify_handle, shopify_variant
from .base import AdapterContext, SourceAdapter, page_listings


_LOCALE = {"en", "us", "www", "com", "html", "htm", "p"}


def _slug_tokens(url: str) -> set:
    return set(norm_text(urlparse(url).path.replace("-", " ").replace("_", " ")).split())


class ProductPageAdapter(SourceAdapter):
    name = "page"
    evidence = VERIFIED_DIRECT

    def __init__(self, shopify=None, bestbuy=None):
        self.shopify = shopify
        self.bestbuy = bestbuy

    def fetch(self, url: str, item: Item, ctx: AdapterContext, vendor: str = "", evidence: str = VERIFIED_DIRECT,
              allow_browser: bool = True, **kw) -> SourceResult:
        nurl = normalize_url(url)
        host = host_of(nurl)
        tag = host or url[:40]
        # 1) official merchant API (Best Buy) beats scraping the same page
        if self.bestbuy and "bestbuy.com" in host:
            r = self.bestbuy.fetch(nurl, item, ctx, vendor=vendor)
            if r.listings:
                return r
            api_note = r.detail
        else:
            api_note = ""
        # 2) Shopify product JSON (every variant / pack size with its own price + stock)
        handle = shopify_handle(nurl)
        if self.shopify and handle:
            base = f"{urlparse(nurl).scheme}://{urlparse(nurl).netloc}"
            r = self.shopify.product(base, handle, item, ctx, vendor=vendor, evidence=evidence,
                                     variant=shopify_variant(nurl))
            if r.listings:
                return r
        # 3) the page itself
        f = ctx.fetcher
        res = f.get(nurl, conditional=True)
        page, how = None, ""
        if res.not_modified and f.cached(nurl):
            page, how = ProductPage.from_dict(f.cached(nurl)["parsed"]), "cached (HTTP 304)"
        elif res.ok:
            page = extract_product(res.text, item.target, item.product)
            how = res.method
        final = res.final_url or nurl
        # stale-link detection: a product URL that now lands on a category/home/search page (or on an
        # unrelated path with no product data) is reported as an identity mismatch, never priced
        if res.ok and final and normalize_url(final) != nurl and not looks_like_listing_page(nurl):
            req, got = _slug_tokens(nurl) - _LOCALE, _slug_tokens(final) - _LOCALE
            overlap = len(req & got) / max(1, len(req))
            if looks_like_listing_page(final) or (not (page and page.offers) and overlap < 0.3):
                return SourceResult(self.name, tag, Outcome.IDENTITY_MISMATCH,
                                    f"stale link: redirected to a non-product page ({final[:90]})")
        # 4) last resort: browser render when HTTP could not expose product data
        need_render = (res.ok and not (page and page.offers) and looks_js_rendered(res.text)) or \
                      (res.ok and not (page and page.offers)) or res.outcome == Outcome.BLOCKED
        if need_render and allow_browser and not ctx.no_browser and f.browser and f.browser.enabled:
            rr = f.render(nurl)
            if rr.ok:
                p2 = extract_product(rr.text, item.target, item.product)
                if p2 and p2.offers:
                    page, how, final = p2, "browser render", rr.final_url or final
                    res = rr
                    f.remember(host, renderer="browser")
            elif not res.ok:
                res.detail = f"{res.detail}; browser: {rr.detail}"
        if not res.ok and not (page and page.offers):
            extra = f" | Best Buy API: {api_note}" if api_note else ""
            return SourceResult(self.name, tag, res.outcome, res.detail + extra)
        if not page or not page.offers:
            why = "no price data on page"
            if res.ok and looks_js_rendered(res.text):
                why += " (JavaScript-rendered" + ("; browser unavailable)" if not (f.browser and f.browser.enabled) else ")")
            return SourceResult(self.name, tag, Outcome.PARSER, why)
        if not res.not_modified:
            f.store(nurl, res, page.to_dict())
            f.remember(host, extractor=page.method)
        if final and normalize_url(final) != nurl:
            page.canonical = page.canonical or final
        page_slug = _slug_tokens(final)
        ls = page_listings(page, nurl, vendor or host, "page" if how != "browser render" else "browser", evidence,
                           want_variant=kw.get("want_variant"))
        for l in ls:
            l.method = f"{page.method}{' via ' + how if how and how not in ('requests',) else ''}"
            l.page_slug = " ".join(sorted(page_slug))[:200]     # used by Product URL identity validation
        oos = all(l.in_stock is False for l in ls)
        detail = f"{page.method}" + (f" ({how})" if how and how != "requests" else "")
        if final and normalize_url(final) != nurl:
            detail += f"; redirected to {final[:80]}"
        return SourceResult(self.name, tag, Outcome.UNAVAILABLE if oos else Outcome.SUCCESS, detail, ls)
