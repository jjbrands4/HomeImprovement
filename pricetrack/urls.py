"""URL helpers: normalisation before persistence/dedup, merchant item ids, domains."""
from __future__ import annotations

import re
from typing import Optional
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

# Query parameters that only track a click / campaign / session. Everything else is kept, because
# it may select a variant (Shopify ?variant=, Best Buy ?skuId=, Amazon ?th=/psc=, color/size ...).
TRACKING_PARAMS = {
    "gclid", "gclsrc", "gbraid", "wbraid", "dclid", "fbclid", "msclkid", "yclid", "twclid", "ttclid",
    "srsltid", "gad_source", "gad_campaignid", "_ga", "_gl", "mc_cid", "mc_eid", "igshid", "si",
    "ref", "ref_", "refid", "referrer", "extstoreid", "loc", "irclickid", "irgwc", "clickid", "click_id",
    "affid", "aff_id", "affiliate", "tag", "linkcode", "creative", "creativeasin", "ascsubtag",
    "cmp", "intcmp", "icid", "cid", "scid", "spm", "ved", "ei", "_pos", "_sid", "_ss", "_psq", "_fid",
    "_v", "pf_rd_p", "pf_rd_r", "pd_rd_r", "pd_rd_w", "pd_rd_wg", "pd_rd_i", "content-id", "qid", "sr",
    "keywords", "crid", "sprefix", "dib", "dib_tag", "s_kwcid", "ef_id", "veh", "cjevent", "cjdata",
    "rrid", "mkcid", "mkrid", "campid", "toolid", "customid", "mkevt", "epik", "lgeo", "acqchannel",
    "athcpid", "athpgid", "athznid", "athieid", "athena", "adid", "adgroupid", "campaignid", "feeditemid",
    "targetid", "matchtype", "network", "device", "devicemodel", "placement", "adposition", "hsa_acc",
}
TRACKING_PREFIXES = ("utm_", "pk_", "mtm_", "hsa_", "_hs", "gad_", "ga_", "pd_rd", "pf_rd", "oly_", "vero_")


def normalize_url(url: str) -> str:
    """Canonical form used for persistence and de-duplication:
       lower-case scheme/host, no fragment, no tracking parameters, no trailing slash,
       remaining parameters sorted (so ?b=1&a=2 == ?a=2&b=1)."""
    if not url:
        return ""
    url = url.strip()
    try:
        p = urlparse(url)
    except ValueError:
        return url
    if not p.scheme or not p.netloc:
        return url
    q = [(k, v) for k, v in parse_qsl(p.query, keep_blank_values=False)
         if k.lower() not in TRACKING_PARAMS and not k.lower().startswith(TRACKING_PREFIXES)]
    q.sort()
    path = re.sub(r"/{2,}", "/", p.path or "/")
    if len(path) > 1:
        path = path.rstrip("/")
    return urlunparse((p.scheme.lower(), p.netloc.lower(), path, "", urlencode(q, doseq=True), ""))


def host_of(url: str) -> str:
    """'https://www.Sonos.com/x' -> 'sonos.com' (www./m. stripped)."""
    try:
        h = urlparse(url).netloc.lower().split(":")[0]
    except ValueError:
        return ""
    return re.sub(r"^(?:www\d?|m|shop|store)\.", "", h)


def same_site(host: str, domain: str) -> bool:
    host, domain = host_of("https://" + host), host_of("https://" + domain)
    return bool(host and domain) and (host == domain or host.endswith("." + domain))


# Merchant-specific item ids that survive URL rewrites (used for stable offer ids).
_ITEM_ID_PATTERNS = [
    ("bestbuy.com", re.compile(r"(?:skuId=|/sku/|/)(\d{7})(?:\.p\b|\b)")),
    ("amazon.", re.compile(r"/(?:dp|gp/product|gp/aw/d|product)/([A-Z0-9]{10})(?:[/?]|$)")),
    ("walmart.com", re.compile(r"/ip/(?:[^/]+/)?(\d{6,})")),
    ("target.com", re.compile(r"/A-(\d{6,})")),
    ("homedepot.com", re.compile(r"/p/(?:[^/]+/)?(\d{9})")),
    ("lowes.com", re.compile(r"/pd/[^/]+/(\d{6,})")),
    ("dell.com", re.compile(r"/apd/([a-z0-9-]{5,})", re.I)),
    ("bhphotovideo.com", re.compile(r"/c/product/(\d+-[A-Z]+)", re.I)),
    ("microcenter.com", re.compile(r"/product/(\d{6,})")),
    ("ebay.com", re.compile(r"/itm/(?:[^/]+/)?(\d{9,})")),
]


def merchant_item_id(url: str) -> str:
    """Best-effort merchant item id from a product URL ('' when unknown).
       Shopify variant ids and generic ?sku= / ?skuId= parameters are recognised everywhere."""
    if not url:
        return ""
    low = url.lower()
    for dom, rx in _ITEM_ID_PATTERNS:
        if dom in low:
            m = rx.search(url)
            if m:
                return m.group(1)
    m = re.search(r"[?&]variant=(\d+)", url)
    if m:
        return "v" + m.group(1)
    m = re.search(r"[?&](?:skuid|sku|pid|productid|itemid)=([A-Za-z0-9_-]{4,})", url, re.I)
    if m:
        return m.group(1)
    m = re.search(r"/products/([^/?#]+)", url)
    if m:
        return "h:" + m.group(1).lower()
    return ""


def shopify_handle(url: str) -> Optional[str]:
    m = re.search(r"/products/([^/?#.]+)", urlparse(url).path or "")
    return m.group(1) if m else None


def shopify_variant(url: str) -> Optional[str]:
    m = re.search(r"[?&]variant=(\d+)", url or "")
    return m.group(1) if m else None


_NON_PRODUCT_RX = re.compile(
    r"/(?:search|s|blog|blogs|support|help|compare|collections?|category|categories|c|cat|stories|press|news|"
    r"account|cart|login|signin|sign-in|deals|brands?|shop-all|all-products|sitemap|pages|faq|contact)(?:/|$)",
    re.I)


def looks_like_listing_page(url: str) -> bool:
    """True for category / search / home pages - a Product URL redirecting here is stale."""
    p = urlparse(url or "")
    path = (p.path or "/").lower()
    if path in ("", "/") or re.fullmatch(r"/[a-z]{2}[-_][a-z]{2}/?", path):
        return True
    if any(m in path for m in ("/products/", "/product/", "/p/", "/pd/", "/apd/", "/dp/", "/ip/", "/sku/", "/itm/")) \
            or re.search(r"/a-\d{6,}|\d{7}\.p$", path):
        return False
    return bool(_NON_PRODUCT_RX.search(path)) or bool(re.search(r"[?&](?:q|query|searchterm|ntt|k)=", p.query or "", re.I))


def us_locale_ok(url: str) -> bool:
    """Skip other-country pages (/en-gb/, /de-de/ ...) so prices are in USD for the US store."""
    for seg in urlparse(url).path.lower().split("/"):
        if re.fullmatch(r"[a-z]{2}[-_][a-z]{2}", seg) and seg.replace("_", "-") != "en-us":
            return False
    return True
