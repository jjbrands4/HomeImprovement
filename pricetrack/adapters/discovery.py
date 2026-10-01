"""Retailer discovery: find the merchant's own product page, then price it from that page.

Order per retailer domain (cheapest / most reliable first):
    1. discovery cache  - a page validated on an earlier run is simply re-fetched (no rediscovery)
    2. Shopify search   - when the domain is a Shopify store (auto-detected)
    3. site search      - the retailer's own search page (template per known retailer)
    4. sitemap          - robots.txt -> sitemap index -> product sitemaps (rarely bot-blocked)
    5. search engines   - 'site:' query on DuckDuckGo / Bing (DTC vendors only; often rate-limited)
Discovered pages are fetched directly and must pass product matching -> evidence verified_discovered.
A negative result is cached too, so a retailer that doesn't carry the product isn't re-crawled every run."""
from __future__ import annotations

import base64
import gzip
import re
import time
from html import unescape as html_unescape
from urllib.parse import quote_plus, unquote, urlparse

from ..identity import ACCESSORY_WORDS, classify, has_identifier_hit
from ..models import Item, Outcome, SourceResult, VERIFIED_DISCOVERED, Vendor
from ..text import norm_text
from ..urls import host_of, looks_like_listing_page, normalize_url, same_site, us_locale_ok
from .base import AdapterContext, SourceAdapter

SITEMAP_MAX_FILES = 6
DIRECT_MAX_PAGES = 2
DISCOVERY_TTL_DAYS = 30          # re-validate (re-fetch) every run; rediscover only after this
NEGATIVE_TTL_DAYS = 10           # 'retailer does not carry it' is re-checked after this many days

# Known retailer domains (vendor key -> domain). Extend via a 'Domain' column in the vendor tab.
KNOWN_DOMAINS = {
    "philipshue": "www.philips-hue.com", "sonos": "www.sonos.com", "thirdreality": "www.thirdreality.com",
    "bhphotovideo": "www.bhphotovideo.com", "superbrightleds": "www.superbrightleds.com", "dell": "www.dell.com",
    "microcenter": "www.microcenter.com", "homedepot": "www.homedepot.com", "lowes": "www.lowes.com",
    "bestbuy": "www.bestbuy.com", "amazon": "www.amazon.com",
}
# Retailers whose pages are never crawled here: Best Buy is priced through its official API instead,
# Amazon forbids automated access (its prices come from Product URLs you add or Google Shopping).
NO_CRAWL = {"amazon.com", "bestbuy.com", "walmart.com", "target.com", "ebay.com"}
SITE_SEARCH = {
    "bhphotovideo.com": "https://www.bhphotovideo.com/c/search?q={q}",
    "microcenter.com": "https://www.microcenter.com/search/search_results.aspx?Ntt={q}",
    "homedepot.com": "https://www.homedepot.com/s/{q}",
    "lowes.com": "https://www.lowes.com/search?searchTerm={q}",
    "dell.com": "https://www.dell.com/en-us/search/{q}",
    "superbrightleds.com": "https://www.superbrightleds.com/catalogsearch/result/?q={q}",
}
# Big-box sitemaps are thousands of multi-MB files with no product ordering - site search only.
SITEMAP_SKIP = {"homedepot.com", "lowes.com", "dell.com", "microcenter.com", "bhphotovideo.com", "walmart.com",
                "target.com", "bestbuy.com", "amazon.com"}
PRODUCT_PATH_HINT = {
    "homedepot.com": "/p/", "lowes.com": "/pd/", "bhphotovideo.com": "/c/product/", "microcenter.com": "/product/",
    "dell.com": "/apd/",
}
SEARCH_ENGINES = [
    ("DuckDuckGo", "post", "https://html.duckduckgo.com/html/", lambda q: {"data": {"q": q}}),
    ("DuckDuckGo Lite", "post", "https://lite.duckduckgo.com/lite/", lambda q: {"data": {"q": q}}),
    ("Bing", "get", "https://www.bing.com/search", lambda q: {"params": {"q": q, "setlang": "en-US", "cc": "US"}}),
]
_BLOCKED = {202, 401, 403, 429, 503}


def rank_product_urls(urls: list, item: Item, vendor: Vendor, limit: int = DIRECT_MAX_PAGES) -> list:
    """Pick the URLs whose path best matches the product (brand words implied by the domain are
    ignored). An identifier in the path (SKU / GTIN) counts double."""
    brand = set(re.split(r"[^a-z0-9]+", (vendor.domain + " " + vendor.name).lower()))
    name_tokens = [t for t in norm_text(item.product).split() if t not in brand] or norm_text(item.product).split()
    ids = [re.sub(r"[^a-z0-9]", "", k.lower()) for k in item.all_mpns] + [g.lstrip("0") for g in item.all_gtins]
    hint = PRODUCT_PATH_HINT.get(host_of("https://" + vendor.domain), "")
    scored = []
    for u in dict.fromkeys(urls):
        path = unquote(urlparse(u).path).lower()
        if not us_locale_ok(u) or path in ("", "/") or looks_like_listing_page(u):
            continue
        toks = set(norm_text(path.replace("-", " ").replace("_", " ")).split())
        comp = norm_text(path).replace(" ", "")
        if (toks & ACCESSORY_WORDS) - set(name_tokens):
            continue
        cov = sum(1 for t in name_tokens if t in toks or (len(t) >= 4 and t in comp)) / len(name_tokens)
        id_hit = any(k and len(k) >= 5 and k in comp for k in ids)
        if cov >= 0.5 or id_hit:
            scored.append((cov + (2 if id_hit else 0) + (0.2 if hint and hint in path else 0), -len(path), u))
    scored.sort(reverse=True)
    return [u for _, _, u in scored[:limit]]


def _extract_links(page_html: str, domain: str) -> list:
    """Same-domain links from a search-results page (DuckDuckGo / Bing redirect links decoded)."""
    bare, out = host_of("https://" + domain), []
    for raw in re.findall(r"""href=["']([^"']+)["']""", page_html or ""):
        u = html_unescape(raw)
        m = re.search(r"uddg=([^&]+)", u)
        if m:
            u = unquote(m.group(1))
        elif "bing.com/ck/a" in u:
            m = re.search(r"[?&]u=a1([^&]+)", u)
            if not m:
                continue
            b64 = m.group(1) + "=" * (-len(m.group(1)) % 4)
            try:
                u = base64.urlsafe_b64decode(b64).decode("utf-8", "ignore")
            except ValueError:
                continue
        if u.startswith("//"):
            u = "https:" + u
        elif u.startswith("/"):
            u = f"https://{domain}{u}"
        if u.startswith("http") and same_site(host_of(u), bare) and u not in out:
            out.append(u)
    return out


class RetailerDiscoveryAdapter(SourceAdapter):
    name = "discovered"
    evidence = VERIFIED_DISCOVERED

    def __init__(self, page_adapter, shopify):
        self.page = page_adapter
        self.shopify = shopify
        self._sitemaps: dict = {}

    # ---- link sources ---------------------------------------------------------------------------
    def sitemap_links(self, domain: str, ctx: AdapterContext) -> list:
        if domain in self._sitemaps:
            return self._sitemaps[domain]
        f = ctx.fetcher
        urls, fetched = [], 0
        r = f.get(f"https://{domain}/robots.txt", retries=0, allow_cffi=False)
        queue = re.findall(r"(?im)^\s*sitemap:\s*(\S+)", r.text) if r.ok else []
        queue = queue or [f"https://{domain}/sitemap.xml"]
        while queue and fetched < SITEMAP_MAX_FILES:
            sm = queue.pop(0)
            fetched += 1
            r = f.get(sm, retries=0, allow_cffi=False)
            if not r.ok:
                continue
            body = r.content or (r.text or "").encode()
            if sm.endswith(".gz") or body[:2] == b"\x1f\x8b":
                try:
                    body = gzip.decompress(body)
                except OSError:
                    continue
            text = body.decode("utf-8", "ignore")
            locs = [html_unescape(x) for x in re.findall(r"<loc>\s*(?:<!\[CDATA\[)?\s*([^<\]\s]+)", text)]
            if "<sitemapindex" in text:
                kids = [k for k in locs if us_locale_ok(k)]
                kids.sort(key=lambda k: ("product" not in k.lower(), not re.search(r"en[-_]?us|/us/", k.lower())))
                queue = kids + queue
            else:
                urls.extend(locs)
        self._sitemaps[domain] = urls
        return urls

    def site_search_links(self, domain: str, item: Item, ctx: AdapterContext) -> tuple:
        tpl = SITE_SEARCH.get(host_of("https://" + domain))
        if not tpl:
            return [], ""
        qs = [m for m in item.all_mpns[:1]] + [item.product]
        statuses = []
        for q in qs:
            r = ctx.fetcher.get(tpl.format(q=quote_plus(q)), retries=0)
            if r.ok:
                links = _extract_links(r.text, domain)
                if links:
                    return links, f"site search ({'model' if q != item.product else 'name'})"
                statuses.append("no links")
            else:
                statuses.append(r.outcome)
                if r.outcome == Outcome.BLOCKED:
                    break
        return [], "site search: " + ", ".join(statuses)

    def engine_links(self, domain: str, query: str, ctx: AdapterContext) -> tuple:
        statuses, sess = [], ctx.fetcher.session
        for name, method, url, kw in SEARCH_ENGINES:
            if ctx.fetcher.circuit_open(host_of(url)):
                statuses.append(f"{name} circuit open")
                continue
            try:
                fn = sess.post if method == "post" else sess.get
                from ..fetch import BROWSER_HEADERS
                r = fn(url, headers=BROWSER_HEADERS, timeout=15, **kw(f"site:{domain} {query}"))
            except Exception as e:
                statuses.append(f"{name} {type(e).__name__}")
                continue
            links = _extract_links(r.text, domain) if r.status_code == 200 else []
            if links:
                return links, f"search ({name})"
            statuses.append(f"{name} blocked (HTTP {r.status_code})" if r.status_code in _BLOCKED else f"{name} no results")
            ctx.fetcher._record(host_of(url), Outcome.BLOCKED if r.status_code in _BLOCKED else Outcome.NO_MATCH, name)
            ctx.fetcher.sleep(1)
        return [], "search engines: " + ", ".join(statuses)

    # ---- main -----------------------------------------------------------------------------------
    def search(self, item: Item, ctx: AdapterContext, vendor: Vendor = None, allow_engines: bool = False,
               **kw) -> SourceResult:
        dom = vendor.domain
        bare = host_of("https://" + dom)
        tag = vendor.name
        if bare in NO_CRAWL:
            return SourceResult(self.name, tag, Outcome.SKIPPED, "not crawled (API / Product URL / Google Shopping only)")
        cached = ctx.state.discovery_get(item, bare)
        if cached and not ctx.force_discovery:
            age = ctx.state.age_days(cached.get("ts"))
            if cached.get("ok") and cached.get("url") and age <= DISCOVERY_TTL_DAYS:
                r = self.page.fetch(cached["url"], item, ctx, vendor=vendor.name, evidence=VERIFIED_DISCOVERED)
                if r.listings and self._matches(item, r.listings):
                    r.source, r.target = self.name, tag
                    r.detail = f"cached page re-verified ({r.detail})"
                    ctx.state.discovery_put(item, bare, url=cached["url"], ok=True)
                    return r
                if r.outcome in (Outcome.BLOCKED, Outcome.NETWORK):
                    return SourceResult(self.name, tag, r.outcome, f"cached page: {r.detail}")
                ctx.state.discovery_put(item, bare, ok=False, detail=f"cached page no longer matches ({r.outcome})")
            elif not cached.get("ok") and age <= NEGATIVE_TTL_DAYS:
                return SourceResult(self.name, tag, Outcome.SKIPPED,
                                    f"no matching page found {age:.0f}d ago ({cached.get('detail', '')[:60]}); "
                                    f"re-check after {NEGATIVE_TTL_DAYS}d")
        if ctx.fetcher.circuit_open(bare):
            return SourceResult(self.name, tag, Outcome.BLOCKED, ctx.fetcher.circuit_open(bare))
        # Shopify store? (big-box retailers never are - don't spend requests asking)
        if bare not in SITEMAP_SKIP and self.shopify.is_shopify(dom, ctx):
            r = self.shopify.search(item, ctx, domain=dom, vendor=vendor.name, evidence=VERIFIED_DISCOVERED)
            good = [l for l in r.listings if self._matches(item, [l])]
            if good:
                ctx.state.discovery_put(item, bare, url=good[0].url, ok=True)
                return SourceResult(self.name, tag, Outcome.SUCCESS, r.detail, good)
            if r.outcome not in (Outcome.BLOCKED, Outcome.NETWORK):
                ctx.state.discovery_put(item, bare, ok=False, detail=r.detail)
            return SourceResult(self.name, tag, r.outcome if r.outcome != Outcome.SUCCESS else Outcome.NO_MATCH, r.detail)
        links, how = self.site_search_links(dom, item, ctx)
        ranked = rank_product_urls(links, item, vendor)
        if not ranked and bare not in SITEMAP_SKIP:
            ranked, how2 = rank_product_urls(self.sitemap_links(dom, ctx), item, vendor), "sitemap"
            how = f"{how}; {how2}" if how else how2
        if not ranked and allow_engines:
            found, how3 = self.engine_links(dom, item.product, ctx)
            ranked = rank_product_urls(found, item, vendor)
            how = f"{how}; {how3}"
        if not ranked:
            ctx.state.discovery_put(item, bare, ok=False, detail=f"no product page ({how})")
            return SourceResult(self.name, tag, Outcome.NO_MATCH, f"no product page ({how})")
        out, problems = [], []
        for u in ranked:
            r = self.page.fetch(u, item, ctx, vendor=vendor.name, evidence=VERIFIED_DISCOVERED)
            good = [l for l in r.listings if self._matches(item, [l])]
            if good:
                out += good
                ctx.state.discovery_put(item, bare, url=normalize_url(u), ok=True)
                break
            problems.append(f"{urlparse(u).path[:50]}: {r.outcome if not r.listings else 'identity_mismatch'}")
        if out:
            return SourceResult(self.name, tag, Outcome.SUCCESS, f"via {how}", out)
        blocked = all("blocked" in p for p in problems)
        if not blocked:
            ctx.state.discovery_put(item, bare, ok=False, detail="; ".join(problems)[:120])
        return SourceResult(self.name, tag, Outcome.BLOCKED if blocked else Outcome.IDENTITY_MISMATCH,
                            f"via {how}: " + "; ".join(problems)[:200])

    @staticmethod
    def _matches(item: Item, listings: list) -> bool:
        """Discovery is held to the normal match standard (not the Product-URL leniency)."""
        for l in listings:
            m = classify(item, l.title, gtins=l.gtins, mpns=l.mpns, pack_qty=l.pack_qty_hint or 1, color=l.color)
            if m.confidence in ("High", "Medium") or (m.confidence != "Low" and has_identifier_hit(item, l.title, l.gtins, l.mpns)):
                return True
        return False
