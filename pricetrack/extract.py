"""
Structured product-data extraction from a merchant page.

Preference order (first source that yields offers wins the price; identifiers are merged from all):
    retailer/API structured data (handled by adapters) -> JSON-LD Product/ProductGroup/Offer
    -> microdata -> OpenGraph/product meta -> scoped hydration JSON (__NEXT_DATA__, window state)
    -> visible DOM price markup
All offers/variants are returned (not the first plausible number), each with its own price,
regular/list price, availability, condition, identifiers and variant attributes.
"""
from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field
from html import unescape as html_unescape
from typing import Optional

from .identity import normalize_gtin, norm_code
from .text import extract_pack_qty, parse_price

_LD_RE = re.compile(r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', re.S | re.I)
_JSON_SCRIPT_RE = re.compile(r'<script[^>]+(?:id=["\']__NEXT_DATA__["\']|type=["\']application/json["\'])[^>]*>(.*?)</script>',
                             re.S | re.I)
_STATE_RE = re.compile(r'window\.(?:__INITIAL_STATE__|__PRELOADED_STATE__|__APOLLO_STATE__|__STATE__|__NUXT__|'
                       r'__PRODUCT__|productData|__data)\s*=\s*({.*?})\s*;?\s*</script>', re.S)

# Keys in hydration JSON that hold "the price you pay"; earlier keys win.
CURRENT_PRICE_KEYS = ("customerPrice", "currentPrice", "salePrice", "finalPrice", "dellPrice", "sellingPrice",
                      "offerPrice", "priceValue", "price", "amount")
REGULAR_PRICE_KEYS = ("regularPrice", "listPrice", "wasPrice", "originalPrice", "compareAtPrice", "compare_at_price",
                      "msrp", "MSRP", "strikePrice", "strikethroughPrice", "basePrice")
# Prices that require qualification - recorded separately, never used as the ordinary price.
CONDITIONAL_KEYS = {"memberPrice": "membership", "clubPrice": "membership", "plusPrice": "membership",
                    "primePrice": "membership", "couponPrice": "coupon", "priceWithCoupon": "coupon",
                    "subscriptionPrice": "subscription", "subscribeAndSavePrice": "subscription",
                    "snsPrice": "subscription", "tradeInPrice": "trade_in", "priceWithTradeIn": "trade_in",
                    "monthlyPrice": "financing", "financePrice": "financing", "installmentPrice": "financing",
                    "activationPrice": "carrier_activation"}


@dataclass
class PageOffer:
    price: float
    currency: str = "USD"
    regular_price: Optional[float] = None
    availability: str = ""            # in_stock | out_of_stock | preorder | backorder | limited | ''
    in_stock: Optional[bool] = None
    condition: str = "new"
    sku: str = ""
    gtin: str = ""
    mpn: str = ""
    name: str = ""
    variant: str = ""
    variant_id: str = ""
    url: str = ""
    shipping: Optional[float] = None
    color: str = ""
    pack_qty: Optional[int] = None
    conditional: str = ""
    conditional_detail: str = ""
    price_kind: str = ""              # '' | aggregate_low


@dataclass
class ProductPage:
    name: str = ""
    brand: str = ""
    gtins: set = field(default_factory=set)
    mpns: set = field(default_factory=set)
    skus: set = field(default_factory=set)
    offers: list = field(default_factory=list)
    method: str = ""
    is_product: bool = False          # page declares itself a product (schema Product / og:type product)
    title: str = ""
    canonical: str = ""
    color: str = ""

    def to_dict(self) -> dict:
        d = asdict(self)
        for k in ("gtins", "mpns", "skus"):
            d[k] = sorted(d[k])
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "ProductPage":
        d = dict(d)
        offers = [PageOffer(**o) for o in d.pop("offers", [])]
        p = cls(**{k: v for k, v in d.items() if k in cls.__dataclass_fields__})
        p.gtins, p.mpns, p.skus, p.offers = set(p.gtins), set(p.mpns), set(p.skus), offers
        return p


# =============================================================================
# helpers
# =============================================================================

def availability_of(avail) -> tuple:
    """schema.org / free-text availability -> (normalised label, in_stock flag)."""
    a = str(avail or "").lower().replace(" ", "").replace("_", "").replace("-", "")
    if not a:
        return "", None
    if "preorder" in a or "presale" in a:
        return "preorder", None
    if "backorder" in a:
        return "backorder", None
    if "limitedavailability" in a:
        return "limited", True
    if any(w in a for w in ("instock", "onlineonly", "instoreonly", "available", "addtocart", "true")) and \
            not any(w in a for w in ("unavailable", "notavailable")):
        return "in_stock", True
    if any(w in a for w in ("outofstock", "soldout", "discontinued", "unavailable", "notavailable", "false")):
        return "out_of_stock", False
    return "", None


def _condition_of(v) -> str:
    v = str(v or "").lower()
    if "refurb" in v:
        return "refurbished"
    if "used" in v or "damaged" in v:
        return "used"
    return "new"


def _first(v):
    if isinstance(v, list):
        return v[0] if v else None
    return v


def _text(v) -> str:
    v = _first(v)
    if isinstance(v, dict):
        v = v.get("name") or v.get("@id") or ""
    return html_unescape(str(v or "")).strip()


def _walk(node, depth=0):
    if depth > 8:
        return
    if isinstance(node, dict):
        yield node
        for v in node.values():
            if isinstance(v, (dict, list)):
                yield from _walk(v, depth + 1)
    elif isinstance(node, list):
        for v in node:
            yield from _walk(v, depth + 1)


def _types(node) -> set:
    t = node.get("@type")
    return {str(x).split("/")[-1] for x in (t if isinstance(t, list) else [t]) if x}


def _gtins_of(node) -> set:
    out = set()
    for k in ("gtin", "gtin8", "gtin12", "gtin13", "gtin14", "upc", "ean", "isbn"):
        for v in (node.get(k) if isinstance(node.get(k), list) else [node.get(k)]):
            g = normalize_gtin(v)
            if g:
                out.add(g)
    pid = str(node.get("productID") or "")
    if pid.lower().startswith(("gtin", "upc", "ean")):
        g = normalize_gtin(pid.split(":")[-1])
        if g:
            out.add(g)
    return out


def _variant_label(node) -> tuple:
    """(label, color, pack) from additionalProperty / color / size fields of a variant node."""
    parts, color, pack = [], _text(node.get("color")), None
    for k in ("color", "size", "pattern", "material"):
        v = _text(node.get(k))
        if v:
            parts.append(v)
    for prop in node.get("additionalProperty") or []:
        if isinstance(prop, dict):
            n, v = _text(prop.get("name")), _text(prop.get("value"))
            if v:
                parts.append(v)
                if re.search(r"pack|quantity|count|qty", n, re.I):
                    pack = extract_pack_qty(v + " pack") if re.fullmatch(r"\d+", v.strip()) else extract_pack_qty(v)
                if re.search(r"colou?r", n, re.I):
                    color = v
    return " / ".join(dict.fromkeys(parts)), color, pack


# =============================================================================
# JSON-LD
# =============================================================================

def _ld_offers(offers, product: dict, variant_node: Optional[dict] = None) -> list:
    out = []
    offers = offers if isinstance(offers, list) else [offers] if isinstance(offers, dict) else []
    vnode = variant_node or {}
    label, color, pack = _variant_label(vnode) if vnode else ("", "", None)
    for o in offers:
        if not isinstance(o, dict):
            continue
        types = _types(o)
        if "AggregateOffer" in types and o.get("offers"):
            out += _ld_offers(o.get("offers"), product, variant_node)
            continue
        specs = o.get("priceSpecification")
        specs = specs if isinstance(specs, list) else [specs] if isinstance(specs, dict) else []
        price, regular, conditional, cdetail = parse_price(o.get("price")), None, "", ""
        currency = o.get("priceCurrency") or ""
        for sp in specs:
            ptype = str(sp.get("priceType") or "").split("/")[-1].lower()
            p = parse_price(sp.get("price"))
            if p is None:
                continue
            currency = currency or sp.get("priceCurrency") or ""
            if sp.get("validForMemberTier") or sp.get("eligibleCustomerType") or "member" in ptype:
                conditional, cdetail = "membership", f"member price ${p:,.2f}"
                continue
            if ptype in ("listprice", "strikethroughprice", "msrp", "srp", "minimumadvertisedprice"):
                regular = p
            elif price is None:
                price = p
        kind = ""
        if price is None and "AggregateOffer" in types:
            price, kind = parse_price(o.get("lowPrice")), "aggregate_low"
        if price is None or price <= 0:
            continue
        if (currency or "USD").upper() != "USD":
            continue
        avail, stock = availability_of(o.get("availability"))
        ship = None
        sd = o.get("shippingDetails")
        sd = sd[0] if isinstance(sd, list) and sd else sd
        if isinstance(sd, dict):
            rate = sd.get("shippingRate")
            rate = rate[0] if isinstance(rate, list) and rate else rate
            if isinstance(rate, dict):
                ship = parse_price(rate.get("value"))
        g = _gtins_of(o) | _gtins_of(vnode)
        out.append(PageOffer(
            price=price, currency="USD", regular_price=regular if regular and regular > price else None,
            availability=avail, in_stock=stock, condition=_condition_of(o.get("itemCondition")),
            sku=_text(o.get("sku") or vnode.get("sku")), gtin=sorted(g)[0] if g else "",
            mpn=_text(o.get("mpn") or vnode.get("mpn")), name=_text(vnode.get("name")) or "",
            variant=label, variant_id=_text(vnode.get("sku") or vnode.get("productID") or ""),
            url=_text(o.get("url") or vnode.get("url")), shipping=ship, color=color, pack_qty=pack,
            price_kind=kind))
        if conditional:                       # member price published next to the ordinary price
            out[-1].conditional_detail = cdetail
    return out


def parse_jsonld(html: str) -> Optional[ProductPage]:
    page, consumed = None, set()
    for block in _LD_RE.findall(html or ""):
        try:
            data = json.loads(block.strip())
        except ValueError:
            try:                                           # tolerate trailing commas / control chars
                data = json.loads(re.sub(r",\s*([}\]])", r"\1", block.strip()).replace("\n", " "))
            except ValueError:
                continue
        for node in _walk(data):
            types = _types(node)
            if id(node) in consumed or not types & {"Product", "ProductGroup", "IndividualProduct", "ProductModel"}:
                continue
            page = page or ProductPage(method="JSON-LD", is_product=True)
            page.name = page.name or _text(node.get("name"))
            page.brand = page.brand or _text(node.get("brand") or node.get("manufacturer"))
            page.gtins |= _gtins_of(node)
            for k in ("mpn", "model"):
                v = _text(node.get(k))
                if v and len(norm_code(v)) >= 4:
                    page.mpns.add(v)
            if _text(node.get("sku")):
                page.skus.add(_text(node.get("sku")))
            page.color = page.color or _text(node.get("color"))
            variants = node.get("hasVariant") or []
            variants = variants if isinstance(variants, list) else [variants]
            got = False
            for v in variants:
                if isinstance(v, dict):
                    consumed.add(id(v))
                if isinstance(v, dict) and v.get("offers"):
                    page.offers += _ld_offers(v.get("offers"), node, v)
                    got = True
            if not got and node.get("offers"):
                page.offers += _ld_offers(node.get("offers"), node)
    if page:
        for o in page.offers:          # (offer GTINs stay on the offer: multipacks have their own GTIN)
            if o.mpn:
                page.mpns.add(o.mpn)
    return page


# =============================================================================
# microdata / meta / hydration / DOM
# =============================================================================

def parse_microdata(html: str) -> Optional[ProductPage]:
    """schema.org microdata (<span itemprop="price" content="219.00">) - Dell and many older stores."""
    h = html or ""
    prices = re.findall(r"""itemprop=["']price["'][^>]*?content=["']([^"']+)""", h, re.I) or \
        re.findall(r"""content=["']([\d.,]+)["'][^>]*?itemprop=["']price["']""", h, re.I)
    prices = [p for p in (parse_price(x) for x in prices) if p]
    if not prices:
        return None
    a = re.search(r"""itemprop=["']availability["'][^>]*?(?:href|content)=["']([^"']+)""", h, re.I)
    n = re.search(r"""itemprop=["']name["'][^>]*?content=["']([^"']+)""", h, re.I)
    avail, stock = availability_of(a.group(1)) if a else ("", None)
    page = ProductPage(method="microdata", is_product="schema.org/Product" in h, name=html_unescape(n.group(1)) if n else "")
    for k in ("gtin13", "gtin12", "gtin14", "gtin8", "gtin"):
        for v in re.findall(r"""itemprop=["']%s["'][^>]*?content=["']([^"']+)""" % k, h, re.I):
            g = normalize_gtin(v)
            if g:
                page.gtins.add(g)
    for v in re.findall(r"""itemprop=["']mpn["'][^>]*?content=["']([^"']+)""", h, re.I):
        page.mpns.add(v.strip())
    for p in dict.fromkeys(prices):
        page.offers.append(PageOffer(price=p, availability=avail, in_stock=stock))
    return page


def _meta(h: str, *names) -> str:
    for n in names:
        m = re.search(r"""<meta[^>]+(?:property|name|itemprop)=["']%s["'][^>]+content=["']([^"']*)""" % re.escape(n), h, re.I) or \
            re.search(r"""<meta[^>]+content=["']([^"']*)["'][^>]+(?:property|name|itemprop)=["']%s["']""" % re.escape(n), h, re.I)
        if m:
            return html_unescape(m.group(1)).strip()
    return ""


def parse_meta(html: str) -> Optional[ProductPage]:
    """OpenGraph / product meta tags."""
    h = html or ""
    price = parse_price(_meta(h, "product:price:amount", "og:price:amount", "product:sale_price:amount"))
    if not price:
        return None
    cur = _meta(h, "product:price:currency", "og:price:currency") or "USD"
    if cur.upper() != "USD":
        return None
    regular = parse_price(_meta(h, "product:original_price:amount", "og:original_price:amount"))
    avail, stock = availability_of(_meta(h, "product:availability", "og:availability"))
    page = ProductPage(method="meta tags", is_product=_meta(h, "og:type").lower().startswith("product"),
                       name=_meta(h, "og:title"), brand=_meta(h, "product:brand", "og:brand"))
    for v in (_meta(h, "product:upc"), _meta(h, "product:ean"), _meta(h, "product:gtin"), _meta(h, "og:upc")):
        g = normalize_gtin(v)
        if g:
            page.gtins.add(g)
    rid = _meta(h, "product:retailer_item_id")
    if rid:
        page.skus.add(rid)
    mpn = _meta(h, "product:mfr_part_no")
    if mpn:
        page.mpns.add(mpn)
    page.offers.append(PageOffer(price=price, regular_price=regular if regular and regular > price else None,
                                 availability=avail, in_stock=stock, condition=_condition_of(_meta(h, "product:condition"))))
    return page


def _hydration_docs(html: str) -> list:
    docs = []
    for raw in _JSON_SCRIPT_RE.findall(html or ""):
        try:
            docs.append(json.loads(raw.strip()))
        except ValueError:
            continue
    for raw in _STATE_RE.findall(html or ""):
        try:
            docs.append(json.loads(raw))
        except ValueError:
            continue
    return docs


def _num(v) -> Optional[float]:
    if isinstance(v, dict):
        for k in ("value", "amount", "price", "current", "raw"):
            if k in v:
                return _num(v[k])
        return None
    if isinstance(v, str) and re.search(r"/\s*mo|per month|apr", v, re.I):
        return None
    p = parse_price(v)
    return p if p and p > 0 else None


def parse_hydration(html: str, target: Optional[float] = None, name_hint: str = "") -> Optional[ProductPage]:
    """Scoped embedded-JSON extraction: only objects that look like a PRODUCT (they carry a name/sku/
    identifier next to the price) are considered - not every number on the page."""
    def plausible(p):
        return p and p > 0 and (not target or target * 0.25 <= p <= target * 4)

    best = None
    for doc in _hydration_docs(html):
        for node in _walk(doc):
            keys = set(node.keys())
            if not keys & {"name", "title", "productName", "sku", "skuId", "upc", "gtin", "modelNumber", "mpn"}:
                continue
            price = None
            for k in CURRENT_PRICE_KEYS:
                if k in node:
                    price = _num(node[k])
                    if price:
                        break
            if not plausible(price):
                continue
            regular = next((p for p in (_num(node.get(k)) for k in REGULAR_PRICE_KEYS if k in node) if p), None)
            cond = next(((kind, _num(node.get(k))) for k, kind in CONDITIONAL_KEYS.items() if _num(node.get(k))), None)
            name = _text(node.get("name") or node.get("title") or node.get("productName"))
            avail, stock = availability_of(node.get("availability") or node.get("stockStatus") or node.get("buttonState")
                                           or node.get("inStock") or node.get("available"))
            score = (2 if name_hint and name and set(name_hint.lower().split()) & set(name.lower().split()) else 0) + \
                    (1 if keys & {"sku", "skuId", "upc", "gtin", "modelNumber"} else 0)
            offer = PageOffer(price=price, regular_price=regular if regular and regular > price else None,
                              availability=avail, in_stock=stock, sku=_text(node.get("sku") or node.get("skuId")),
                              gtin=normalize_gtin(node.get("upc") or node.get("gtin")) or "",
                              mpn=_text(node.get("modelNumber") or node.get("mpn")), name=name)
            if cond:
                offer.conditional_detail = f"{cond[0]} price ${cond[1]:,.2f} (not used)"
            if best is None or score > best[0]:
                best = (score, offer)
    if not best:
        return None
    o = best[1]
    page = ProductPage(method="hydration JSON", name=o.name, offers=[o])
    if o.gtin:
        page.gtins.add(o.gtin)
    if o.mpn:
        page.mpns.add(o.mpn)
    return page


def parse_dom(html: str, target: Optional[float] = None) -> Optional[ProductPage]:
    """Visible price markup (last resort): '<span class="price">$219.00</span>'."""
    text = html or ""

    def plausible(p):
        return p and p > 0 and (not target or target * 0.35 <= p <= target * 4)

    for m in re.finditer(r"""(?:class|data-testid|id)=["'][^"']*(?<!was-)(?<!strike-)(?<!old-)price[^"']*["'][^>]*>\s*(?:<[^>]+>\s*){0,4}\$\s*([\d,]+\.\d{2})""",
                         text, re.I):
        after = text[m.end():m.end() + 25].lower()
        before = text[max(0, m.start() - 60):m.start()].lower()
        if re.search(r"^\s*(?:<[^>]+>\s*){0,2}(?:/\s*mo|per month|a month|mo\.)", after) or \
                re.search(r"(?:was|reg\.|compare at|strike|member|coupon)[^<>]{0,30}$", before):
            continue
        p = parse_price(m.group(1))
        if plausible(p):
            stock = False if re.search(r">\s*(?:Sold Out|Out of Stock|Currently unavailable)\s*<", text, re.I) else None
            return ProductPage(method="visible price markup", offers=[PageOffer(price=p, in_stock=stock,
                                                                              availability="out_of_stock" if stock is False else "")])
    return None


def page_title(html: str) -> str:
    t = re.search(r"<title[^>]*>([^<]+)</title>", html or "", re.I)
    return html_unescape(t.group(1)).strip() if t else ""


def looks_js_rendered(html: str) -> bool:
    """Heuristic: an app shell with little text and big bundles -> worth a browser render."""
    h = html or ""
    body = re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", h, flags=re.S | re.I)
    words = len(re.findall(r"[A-Za-z]{3,}", body))
    return words < 250 or bool(re.search(r'id=["\'](?:root|__next|app|__nuxt)["\']\s*>\s*</div>', h))


def looks_blocked(html: str) -> bool:
    h = (html or "")[:20000].lower()
    return any(s in h for s in ("captcha", "px-captcha", "are you a robot", "robot or human", "access denied",
                                "request unsuccessful. incapsula", "pardon our interruption", "attention required",
                                "verify you are human", "unusual traffic", "bot detection", "/_incapsula_resource"))


def extract_product(html: str, target: Optional[float] = None, name_hint: str = "") -> Optional[ProductPage]:
    """Run the extractor chain. Identifiers/brand/name are merged across every extractor that finds
    something; offers come from the highest-priority extractor that produced any."""
    found = []
    for fn in (parse_jsonld, parse_microdata, parse_meta):
        try:
            p = fn(html)
        except Exception:                 # one broken extractor never sinks the page
            p = None
        if p:
            found.append(p)
    if not any(p.offers for p in found):
        for fn in (lambda h: parse_hydration(h, target, name_hint), lambda h: parse_dom(h, target)):
            try:
                p = fn(html)
            except Exception:
                p = None
            if p and p.offers:
                found.append(p)
                break
    if not found:
        return None
    main = next((p for p in found if p.offers), found[0])
    page = ProductPage(name=main.name, brand=main.brand, offers=list(main.offers), method=main.method,
                       is_product=any(p.is_product for p in found), color=main.color)
    for p in found:
        page.name = page.name or p.name
        page.brand = page.brand or p.brand
        page.gtins |= p.gtins
        page.mpns |= p.mpns
        page.skus |= p.skus
    page.title = page_title(html)
    c = re.search(r"""<link[^>]+rel=["']canonical["'][^>]+href=["']([^"']+)""", html or "", re.I)
    page.canonical = html_unescape(c.group(1)) if c else ""
    # aggregate 'from' prices only when nothing better exists
    real = [o for o in page.offers if o.price_kind != "aggregate_low"]
    if real:
        page.offers = real
    return page
