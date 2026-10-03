"""
Product / offer reconciliation.

  offer id     merchant + merchant item id (or canonical URL) + condition + variant + pack
               -> stable across runs; two listings are NEVER merged just because the price is equal
  duplicates   the same offer seen through several sources (Product URL page, Google
               Shopping row ...) collapses to the most authoritative copy, and the others are kept as
               corroboration metadata ("corroborated_by")
  promotion    a Medium listing becomes High only when independent evidence agrees on identity and
               price, and it has no hard identifier/spec conflict
"""
from __future__ import annotations

import difflib
import hashlib

from .identity import fingerprint, item_brand, item_colors, norm_code
from .models import EVIDENCE_RANK, Item, MARKET_SNAPSHOT, VERIFIED
from .text import keys_match, norm_text, strip_pack_phrases, vendor_key
from .urls import merchant_item_id, normalize_url

PRICE_AGREE = 0.15          # corroborating prices must be within +/-15%


def offer_identity(l) -> str:
    vk = vendor_key(l.vendor) or "unknown"
    if l.source == "ebay":
        mid = l.merchant_item_id or normalize_url(l.url)
    elif l.evidence == MARKET_SNAPSHOT:
        mid = "snap:" + (l.merchant_item_id or norm_text(strip_pack_phrases(l.title))[:60])
    else:
        mid = l.merchant_item_id or merchant_item_id(l.url) or normalize_url(l.url)
    variant = l.variant_id or norm_text(l.variant)[:30]
    raw = "|".join([vk, mid, l.condition or "new", variant, str(l.pack_qty)])
    return hashlib.sha1(raw.encode()).hexdigest()[:12]


def assign_ids(item: Item, listings: list) -> None:
    brand = item_brand(item)
    model = item.all_mpns[0] if item.all_mpns else item.product
    gtin = sorted(item.all_gtins)[0] if item.all_gtins else ""
    colors = " ".join(sorted(item_colors(item)))
    for l in listings:
        l.offer_id = offer_identity(l)
        # Listings that matched the item share the item's canonical fingerprint (per pack size);
        # non-matching ones get their own so they never pool with the product.
        if l.confidence in ("High", "Medium"):
            l.fingerprint = fingerprint(brand, model, gtin, colors, l.pack_qty)
        else:
            l.fingerprint = fingerprint(l.brand or "", (sorted(l.mpns)[0] if l.mpns else l.title[:60]),
                                        sorted(l.gtins)[0] if l.gtins else "", l.color, l.pack_qty)


def _authority(l) -> tuple:
    return (0 if l.from_url else 1, EVIDENCE_RANK.get(l.evidence, 9),
            {"High": 0, "Medium": 1}.get(l.confidence, 2), 0 if l.in_stock is not None else 1)


def _corr(l) -> str:
    return f"{l.source}/{l.evidence}: ${l.unit_price:,.2f}" + (f" ({l.vendor})" if l.vendor else "")


def merge_duplicates(listings: list) -> list:
    """Collapse the same offer seen through several sources; keep corroboration metadata."""
    by_id: dict = {}
    for l in sorted(listings, key=_authority):
        k = l.offer_id
        if k in by_id:
            keep = by_id[k]
            if l.source != keep.source or l.evidence != keep.evidence:
                keep.corroborated_by.append(_corr(l))
            continue
        by_id[k] = l
    out = list(by_id.values())
    # Cross-source duplicates: a Google Shopping row for a merchant whose own page/API we priced is
    # the same offer seen through a (possibly stale) feed -> corroboration, not a second listing.
    verified = [l for l in out if l.evidence in VERIFIED and l.confidence != "Low"]
    keep_list = []
    for l in out:
        if l.evidence == MARKET_SNAPSHOT and l.confidence != "Low":
            twin = next((v for v in verified if keys_match(vendor_key(v.vendor), vendor_key(l.vendor))
                         and v.condition == l.condition and v.pack_qty == l.pack_qty
                         and v.fingerprint == l.fingerprint), None)
            if twin:
                diff = abs(l.unit_price - twin.unit_price) / twin.unit_price if twin.unit_price else 0
                twin.corroborated_by.append(_corr(l) + (f" - differs {diff:.0%} (stale feed?)" if diff > 0.02 else " - agrees"))
                continue
        keep_list.append(l)
    return keep_list


def promote(item: Item, listings: list) -> int:
    """Medium -> High only on INDEPENDENT agreement (different merchant or different source type):
         (a) identifier evidence (GTIN/MPN/page metadata) + price agrees with a High listing, or
         (b) title mirrors a VERIFIED High listing of another merchant + price agrees, or
         (c) a verified page whose identity was unconfirmed shares a GTIN/MPN with an independent High listing.
       Listings carrying any hard conflict are Low already and are never promoted."""
    highs = [l for l in listings if l.confidence == "High" and l.condition == "new" and not l.conditional]
    n = 0
    for m in listings:
        if m.confidence != "Medium" or m.conditional:
            continue
        indep = [h for h in highs if h is not m and (not keys_match(vendor_key(h.vendor), vendor_key(m.vendor))
                                                   or h.evidence != m.evidence)]
        agree = [h for h in indep if h.unit_price and m.unit_price and h.pack_qty == m.pack_qty
                 and abs(m.unit_price - h.unit_price) / h.unit_price <= PRICE_AGREE]
        why = ""
        if m.match_evidence in ("gtin", "mpn", "page_metadata") and agree:
            why = f"identifier + price agrees with {agree[0].vendor}"
        elif agree:
            mt = norm_text(strip_pack_phrases(m.title))[:90]
            mirror = next((h for h in agree if h.evidence in VERIFIED and
                           difflib.SequenceMatcher(None, norm_text(strip_pack_phrases(h.title))[:90], mt).ratio() >= 0.85), None)
            if mirror:
                why = f"title mirrors verified listing at {mirror.vendor} + price agrees"
        if not why and m.evidence in VERIFIED and (m.gtins or m.mpns):
            mcodes = {norm_code(x) for x in m.mpns}
            twin = next((h for h in indep if (m.gtins & h.gtins) or (mcodes & {norm_code(x) for x in h.mpns})), None)
            if twin:
                why = f"shares identifier with independent High listing at {twin.vendor}"
        if why:
            m.confidence, m.match_evidence = "High", "corroborated"
            m.conf_reason += f"; promoted: {why}"
            n += 1
    return n
