"""Data models shared by every module."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Optional

# ---- Evidence tiers ------------------------------------------------------------------------------
# How the PRICE was obtained (independent of how sure we are that it is the right product).
VERIFIED_DIRECT = "verified_direct"          # exact merchant page we were given, or an official merchant API
VERIFIED_DISCOVERED = "verified_discovered"  # merchant page we fetched ourselves after sitemap/search discovery
MARKET_SNAPSHOT = "market_snapshot"          # aggregator / search-engine row (Google Shopping via SerpApi ...)
VERIFIED = {VERIFIED_DIRECT, VERIFIED_DISCOVERED}
EVIDENCE_RANK = {VERIFIED_DIRECT: 0, VERIFIED_DISCOVERED: 1, MARKET_SNAPSHOT: 2}


class Outcome:
    """Retrieval outcome categories - never collapse failures into a generic 'N/A'."""
    SUCCESS = "success"
    NO_MATCH = "no_match"                    # source answered, but no listing for this product
    IDENTITY_MISMATCH = "identity_mismatch"  # page fetched, but it is a different product / stale redirect
    BLOCKED = "blocked"                      # bot wall, 401/403/429, captcha, circuit open
    NETWORK = "timeout_network"              # timeouts, DNS, connection resets, 5xx after retries
    PARSER = "parser_failure"                # page fetched but no usable price data could be extracted
    UNAVAILABLE = "unavailable"              # product found but out of stock / discontinued
    API = "api_unavailable"                  # API key missing/invalid, quota exhausted, auth failure
    SKIPPED = "skipped"                      # deliberately not run (flag, cache policy, not relevant)

    ALL = (SUCCESS, NO_MATCH, IDENTITY_MISMATCH, BLOCKED, NETWORK, PARSER, UNAVAILABLE, API, SKIPPED)


def utcnow() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None, microsecond=0)


@dataclass
class Item:
    """One Master Sheet row."""
    wid: object
    product: str
    specs: list
    keywords: list
    location: str
    priority: int
    qty: object
    open_used: bool
    target: Optional[float]
    bulk: bool
    bulk_sizes: set
    skus: list = field(default_factory=list)            # model-number-like keywords (soft identifiers)
    spec_phrases: list = field(default_factory=list)    # specs + non-SKU keywords (must appear in title)
    urls: list = field(default_factory=list)            # optional "Product URLs" column (trusted candidates)
    # ---- optional explicit identifiers (Master Sheet columns, if present) ----
    gtins: set = field(default_factory=set)             # GTIN-14 normalised
    mpns: list = field(default_factory=list)
    brand: str = ""
    # ---- identifiers learned from validated pages / APIs (state file) ----
    learned_gtins: set = field(default_factory=set)
    learned_mpns: set = field(default_factory=set)
    learned_brand: str = ""
    # ---- Master Sheet switches ----
    exclude: list = field(default_factory=list)         # '!phrase' entries from specs / keywords: never searched, never matched
    only_links: bool = False                            # 'Only Check Primary Links' = Yes: Product URLs (+ SerpApi search when < 3 vendor links work)
    links_only_fallback: bool = False                   # asked for Yes but no Product URLs are listed -> normal search is used

    @property
    def qty_needed(self) -> int:
        try:
            return max(1, int(float(self.qty)))
        except (TypeError, ValueError):
            return 1

    @property
    def all_gtins(self) -> set:
        return set(self.gtins) | set(self.learned_gtins)

    @property
    def all_mpns(self) -> list:
        seen, out = set(), []
        for m in list(self.skus) + list(self.mpns) + sorted(self.learned_mpns):
            k = m.upper().replace("-", "").replace(" ", "")
            if k and k not in seen:
                seen.add(k)
                out.append(m)
        return out


@dataclass
class Vendor:
    name: str
    key: str
    is_dtc: bool = False
    domain: str = ""


@dataclass
class Listing:
    """One normalised offer found by any adapter."""
    title: str
    url: str
    price: float                       # current listed price for this listing (may be a multi-pack)
    vendor: str
    source: str                        # adapter: page | shopify | bestbuy_api | discovered | browser | serpapi | ebay
    condition: str = "new"
    in_stock: Optional[bool] = None    # None = unknown
    seller_comment: str = ""
    seller_ok: bool = True
    from_url: bool = False             # came from one of YOUR Master Sheet Product URLs
    # ---- normalised price fields ----
    evidence: str = MARKET_SNAPSHOT
    regular_price: Optional[float] = None    # list / compare-at / MSRP published with the offer
    shipping: Optional[float] = None         # None = unknown, 0.0 = free
    availability: str = ""                   # in_stock | out_of_stock | preorder | backorder | limited | unknown
    pack_qty_hint: Optional[int] = None      # pack size from structured data (variant options) - beats title parsing
    variant: str = ""                        # variant label (e.g. '2 Pack / Black')
    variant_id: str = ""
    merchant_item_id: str = ""
    gtins: set = field(default_factory=set)
    mpns: set = field(default_factory=set)
    brand: str = ""
    color: str = ""
    conditional: str = ""                    # coupon | membership | subscription | financing | trade_in | ...
    conditional_detail: str = ""
    method: str = ""                         # extraction method (JSON-LD, Shopify .js, Best Buy API ...)
    retrieved_at: Optional[datetime] = None
    alt_url: str = ""                        # e.g. the Google Shopping link a snapshot came from
    page_slug: str = ""                      # words from the final page URL path (identity validation)
    # ---- filled in by scoring / reconciliation ----
    pack_qty: int = 1
    effective_price: float = 0.0             # price + known shipping (pre-tax acquisition price)
    unit_price: float = 0.0                  # effective price per single item
    vkey: str = ""
    is_primary: bool = False
    is_resale: bool = False
    confidence: str = "Low"
    conf_reason: str = ""
    match_evidence: str = ""                 # gtin | mpn | page_metadata | title | corroborated | product_url
    eligible: bool = False                   # counts toward market averages
    reportable: bool = False                 # may be reported as a deal / target hit
    trusted: bool = False                    # may feed verified baselines / history
    cond_ok: bool = False                    # would be eligible, but the price is conditional (coupon/member/...)
    deal_rule: str = ""
    ref_price: Optional[float] = None
    ref_type: str = ""
    corroborated_by: list = field(default_factory=list)
    offer_id: str = ""
    fingerprint: str = ""
    price_event: str = ""

    @property
    def is_verified(self) -> bool:
        return self.evidence in VERIFIED


@dataclass
class SourceResult:
    """What one adapter call produced - listings plus an auditable outcome."""
    source: str
    target: str
    outcome: str
    detail: str = ""
    listings: list = field(default_factory=list)

    def note(self) -> str:
        n = f" {len(self.listings)} listing(s)" if self.listings else ""
        return f"{self.source}[{self.target}]: {self.outcome}{n}" + (f" - {self.detail}" if self.detail else "")
