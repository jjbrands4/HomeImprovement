"""
Price correctness: normalised price fields, eligibility, verified baseline, reference-price
hierarchy, market statistics, quantity-aware pack optimisation and deal rules.

  current_price   Listing.price (what the listing costs now, for its whole pack)
  shipping        Listing.shipping (None = unknown, 0 = free)
  effective_price current + known shipping  (pre-tax acquisition price)
  unit_price      effective / pack_qty
  regular_price   list / compare-at / MSRP published with the offer (reference only)

Reference price (what "normal" costs) is kept separate from current market statistics:
  1. official MSRP / regular / compare-at price from a verified merchant (manufacturer first)
  2. trusted regular-price consensus (>= 2 High listings publishing a regular price that agree)
  3. trusted historical reference (trailing verified median / EWMA)
"""
from __future__ import annotations

import statistics
from collections import Counter
from typing import Optional

from .identity import anchor_from, classify, has_identifier_hit, validate_page
from .models import Item, Listing, VERIFIED, VERIFIED_DIRECT
from .text import extract_pack_qty, keys_match, vendor_key
from .urls import host_of

# ---- Deal / sanity settings (price_tracker.py may override these module attributes) ----------------
DEAL_DISCOUNT = 0.15
TREND_MIN_DISCOUNT = 0.0
MARKET_MIN_SAMPLE = 3
MIN_PRICE_RATIO_OF_TARGET = 0.35
OUTLIER_LOW, OUTLIER_HIGH = 0.4, 2.5
MAX_DEALS_PER_ITEM = 5
ANCHOR_BAND_NEW = (0.60, 1.40)
ANCHOR_BAND_RESALE = (0.30, 1.05)
DEAL_MIN_CONFIDENCE = "High"
OOS_COUNTS_FOR_BASELINE = True
RESALE_VS_NEW_DISCOUNT = 0.35
USE_TARGET_RULE = False
REQUIRE_AT_OR_BELOW_TARGET = False
SNAPSHOT_DEALS = True            # market-snapshot rows may be deals, but are labelled and ranked last
REF_LABELS = {"msrp_regular": "regular/MSRP", "regular_consensus": "regular-price consensus",
              "historical_verified": "verified history", "verified_current": "verified price",
              "verified_pack": "verified pack price", "market_average": "market avg (unverified)"}
REGULAR_CONSENSUS_TOL = 0.05


# =============================================================================
# Scoring
# =============================================================================

def normalize_prices(l: Listing) -> None:
    l.pack_qty = max(1, int(l.pack_qty_hint or extract_pack_qty(l.title) or 1))
    l.effective_price = round(l.price + (l.shipping or 0.0), 2)
    l.unit_price = round(l.effective_price / l.pack_qty, 2)
    if not l.availability:
        l.availability = {True: "in_stock", False: "out_of_stock"}.get(l.in_stock, "unknown")


def allowed_packs(item: Item) -> Optional[set]:
    """None = any pack size allowed (bulk item without explicit sizes)."""
    if not item.bulk:
        return {1}
    return ({1} | item.bulk_sizes) if item.bulk_sizes else None


def score_listings(item: Item, listings: list, primary: list, secondary_keys: list, dtc_names: set = frozenset()) -> None:
    """Fill pack / prices / vendor class / identity confidence / eligibility for every listing."""
    for l in listings:
        normalize_prices(l)
        l.vkey = vendor_key(l.vendor)
        l.is_primary = any(keys_match(l.vkey, v.key) for v in primary)
        on_secondary = any(keys_match(l.vkey, k) for k in secondary_keys) or l.source == "ebay"
        l.is_resale = on_secondary or l.condition in ("used", "refurbished")

    anchor = None
    hits = [l for l in listings if has_identifier_hit(item, l.title, l.gtins, l.mpns)
            and classify(item, l.title, gtins=l.gtins, mpns=l.mpns, pack_qty=l.pack_qty, color=l.color).confidence != "Low"]
    if hits:
        best = sorted(hits, key=lambda l: (l.evidence not in VERIFIED, l.is_resale, not l.is_primary))[0]
        anchor = anchor_from(item, best.title)

    packs = allowed_packs(item)
    for l in listings:
        if l.from_url and l.evidence in VERIFIED:
            m = validate_page(item, l.title, l.gtins, l.mpns, brand_hint=l.brand,
                              domain_brand=f"{l.vendor} {host_of(l.url)}", slug=l.page_slug,
                              pack_qty=l.pack_qty, color=l.color)
        else:
            m = classify(item, l.title, anchor, gtins=l.gtins, mpns=l.mpns, brand=l.brand, pack_qty=l.pack_qty,
                         color=l.color)
        l.confidence, l.conf_reason, l.match_evidence = m.confidence, m.reason, m.evidence
        ok = (l.condition != "parts" and (packs is None or l.pack_qty in packs) and l.confidence in ("High", "Medium"))
        if l.in_stock is False and not (OOS_COUNTS_FOR_BASELINE and l.evidence in VERIFIED and not l.is_resale):
            ok = False
        if item.target and not l.from_url:
            ok = ok and l.unit_price >= item.target * MIN_PRICE_RATIO_OF_TARGET
        if l.is_resale and not item.open_used:
            ok = False
        l.cond_ok = ok and bool(l.conditional)
        l.eligible = ok and not l.conditional
        l.reportable = (l.eligible and l.in_stock is not False
                        and (item.open_used or (l.is_primary and not l.is_resale)))


def mark_trusted(listings: list) -> None:
    """Trusted = may feed verified baselines, trends and history: verified evidence + High identity +
    new condition + ordinary (non-conditional) price + still eligible after sanity bands."""
    for l in listings:
        l.trusted = (l.eligible and l.evidence in VERIFIED and l.confidence == "High" and l.condition == "new"
                     and not l.is_resale and not l.conditional)


# =============================================================================
# Baseline / reference / market stats
# =============================================================================

def verified_baseline(listings: list) -> tuple:
    """Current verified market price per single unit (median of verified High listings).
    Returns (price | None, 'verified: ...' source text, {pack: per-unit median})."""
    pool = [l for l in listings if l.eligible and l.evidence in VERIFIED and l.confidence == "High"
            and l.condition == "new" and not l.is_resale and not l.conditional]
    if not pool:
        return None, "none (no verified merchant page/API priced the product)", {}
    in_stock = [l for l in pool if l.in_stock is not False]
    singles_in = [l for l in in_stock if l.pack_qty == 1]
    singles_all = [l for l in pool if l.pack_qty == 1]
    use = singles_in or singles_all or in_stock or pool
    price = round(statistics.median(l.unit_price for l in use), 2)
    tiers = sorted({f"{l.vendor} ({'API' if l.source in ('bestbuy_api', 'ebay') else 'page' if l.evidence == VERIFIED_DIRECT else 'discovered'})"
                    for l in use})
    src = f"verified: {', '.join(tiers)} (median of {len(use)})"
    if not singles_in and singles_all:
        src += "; single unit sold out - list price used"
    base_pool = in_stock or pool
    packs = {q: round(statistics.median(l.unit_price for l in base_pool if l.pack_qty == q), 2)
             for q in {l.pack_qty for l in base_pool}}
    return price, src, packs


def reference_price(item: Item, listings: list, hist: dict, dtc_keys: set = frozenset()) -> tuple:
    """(price | None, type, detail) - see module docstring for the hierarchy."""
    def single_regular(l):
        reg = l.regular_price
        return round(reg / l.pack_qty, 2) if reg else None

    ver = [l for l in listings if l.evidence in VERIFIED and l.confidence == "High" and l.condition == "new"
           and not l.is_resale and (l.eligible or l.cond_ok)]
    ver.sort(key=lambda l: (l.pack_qty != 1, not any(keys_match(l.vkey, k) for k in dtc_keys), not l.from_url))
    # 1a. explicit MSRP / regular / compare-at from a verified merchant (manufacturer first)
    for l in ver:
        r = single_regular(l)
        if r:
            kind = "manufacturer regular/compare-at" if any(keys_match(l.vkey, k) for k in dtc_keys) else "retailer regular/list price"
            return r, "msrp_regular", f"{kind} at {l.vendor}"
    # 1b. manufacturer (DTC) current price, which is normally MSRP
    for l in ver:
        if any(keys_match(l.vkey, k) for k in dtc_keys) and l.pack_qty == 1:
            return round(l.price / l.pack_qty, 2), "msrp_regular", f"manufacturer price at {l.vendor}"
    # 2. regular-price consensus across High listings (incl. Google Shopping 'was' prices)
    regs = [single_regular(l) for l in listings if l.confidence == "High" and l.condition == "new"
            and not l.is_resale and single_regular(l)]
    if len(regs) >= 2:
        med = statistics.median(regs)
        agree = [r for r in regs if abs(r - med) / med <= REGULAR_CONSENSUS_TOL]
        if len(agree) >= 2:
            return round(statistics.median(agree), 2), "regular_consensus", f"{len(agree)} listings publish a regular price"
    # 3. verified history
    if hist.get("median"):
        return hist["median"], "historical_verified", f"30-day verified median ({hist.get('median_n')} runs)"
    if hist.get("ewma"):
        return hist["ewma"], "historical_verified", "verified EWMA"
    # 1c. verified current price as a last resort (not a 'regular' price, labelled as such)
    single = [l for l in ver if l.pack_qty == 1 and l.eligible]
    if single:
        return round(statistics.median(l.unit_price for l in single), 2), "verified_current", "verified current price (no regular price published)"
    return None, "none", "no reference available"


def apply_band(item: Item, listings: list, anchor_price: Optional[float]) -> None:
    """Listings far from the verified anchor price (bundles, wrong models that slipped through,
    marketplace gouging) stop counting toward averages and deals."""
    if not anchor_price:
        return
    for l in listings:
        if l.from_url and l.evidence in VERIFIED:
            continue
        lo, hi = ANCHOR_BAND_RESALE if l.is_resale else ANCHOR_BAND_NEW
        if not (anchor_price * lo <= l.unit_price <= anchor_price * hi):
            if l.eligible or l.cond_ok:
                l.conf_reason += f"; price outside {lo:.0%}-{hi:.0%} of verified price"
            l.eligible = l.reportable = l.cond_ok = False


def pool_stats(prices: list) -> Optional[dict]:
    if not prices:
        return None
    med = statistics.median(prices)
    lo, hi = med * OUTLIER_LOW, med * OUTLIER_HIGH
    kept = [p for p in prices if lo <= p <= hi]
    return {"avg": round(statistics.mean(kept), 2), "n": len(kept), "lo": lo} if kept else None


def market_stats(listings: list) -> dict:
    """Current market averages (High identity only; Medium needs corroboration first)."""
    def pool(resale):
        return [l.unit_price for l in listings if l.eligible and l.is_resale == resale and l.confidence == "High"]
    return {"new": pool_stats(pool(False)), "res": pool_stats(pool(True))}


# =============================================================================
# Quantity-aware pack optimisation
# =============================================================================

def quantity_plan(item: Item, listings: list) -> Optional[dict]:
    """Least-cost combination of purchasable single/multi-pack offers covering Quantity Needed.
    New-condition retail offers may be bought repeatedly; resale offers (eBay / used / refurbished) are
    one-of-a-kind and used at most once. Shipping is counted per purchase (conservative). Ties prefer
    verified evidence, then fewer surplus units."""
    need = item.qty_needed
    opts = [l for l in listings if l.reportable and l.confidence == "High" and l.seller_ok and not l.conditional
            and l.in_stock is not False and l.effective_price > 0]
    if not opts:
        return None
    retail, once = {}, []
    for l in opts:
        if l.is_resale or l.condition != "new":
            once.append(l)
            continue
        key = (l.effective_price, l.evidence not in VERIFIED, not l.is_primary)
        cur = retail.get(l.pack_qty)
        if cur is None or key < (cur.effective_price, cur.evidence not in VERIFIED, not cur.is_primary):
            retail[l.pack_qty] = l
    once = sorted(once, key=lambda l: l.unit_price)[:12]
    INF = float("inf")
    # state: dp[x] = (cost, unverified purchases, surplus, plan tuple) to cover >= x units
    dp = [(0.0, 0, 0, ())] + [(INF, 0, 0, ())] * need

    def relax(table, x, l, src):
        if src[0] == INF:
            return
        surplus = max(0, l.pack_qty - x) + src[2] if x < l.pack_qty else src[2]
        cand = (round(src[0] + l.effective_price, 2), src[1] + (l.evidence not in VERIFIED), surplus, src[3] + (id(l),))
        if cand[:3] < table[x][:3]:
            table[x] = cand

    for l in once:                                     # 0/1 items: iterate x downward on a copy
        new = list(dp)
        for x in range(need, 0, -1):
            relax(new, x, l, dp[max(0, x - l.pack_qty)])
        dp = new
    for l in retail.values():                          # unbounded items: iterate x upward in place
        for x in range(1, need + 1):
            relax(dp, x, l, dp[max(0, x - l.pack_qty)])
    if dp[need][0] == INF:
        return None
    by_id = {id(l): l for l in opts}
    counts = Counter(dp[need][3])
    lines = sorted(((by_id[i], n) for i, n in counts.items()), key=lambda t: (-t[0].pack_qty, t[0].unit_price))
    units = sum(l.pack_qty * n for l, n in lines)
    total = round(sum(l.effective_price * n for l, n in lines), 2)
    text = " + ".join(f"{n} x {'single' if l.pack_qty == 1 else f'{l.pack_qty}-pack'} @ {l.vendor} "
                      f"(${l.effective_price:,.2f}{'' if l.evidence in VERIFIED else ', snapshot'}"
                      f"{', ' + l.condition if l.condition != 'new' else ''})" for l, n in lines)
    singles = [l for l in retail.values() if l.pack_qty == 1]
    naive = round(singles[0].effective_price * need, 2) if singles else None
    return {"text": f"{text} = ${total:,.2f} for {units} unit(s) (${total / units:,.2f}/unit)"
                    + (f"; saves ${naive - total:,.2f} vs {need} new singles" if naive and naive - total >= 0.01 else ""),
            "total": total, "units": units, "unit": round(total / units, 2), "lines": lines}


# =============================================================================
# Deals
# =============================================================================

def find_deals(item: Item, listings: list, stats: dict, hist: dict, reference: tuple,
               baseline: Optional[float], pack_baselines: Optional[dict], resale_trend: Optional[float] = None) -> list:
    """Only High-confidence listings can be deals. Rules (per pool):
      New    A: >= DEAL_DISCOUNT below the reference price (single units) or the verified per-pack price
                (multi-packs; normal bulk pricing is not a deal); falls back to the market average only
                when no verified/reference price exists (needs MARKET_MIN_SAMPLE High listings)
             B: below the trailing 30-day VERIFIED median, or at/below the historical verified low
      Resale A: >= DEAL_DISCOUNT below the resale average (MARKET_MIN_SAMPLE+), or >= 35% below the new reference
             B: below the resale 30-day trend
    Conditional prices (coupon/member/...) are reported separately and labelled."""
    ref_price, ref_type, _ = reference
    deals = []
    for l in listings:
        if not (l.confidence == DEAL_MIN_CONFIDENCE and l.seller_ok and l.in_stock is not False):
            continue
        if not (l.reportable or (l.cond_ok and (item.open_used or (l.is_primary and not l.is_resale)))):
            continue
        if l.evidence not in VERIFIED and not SNAPSHOT_DEALS:
            continue
        if REQUIRE_AT_OR_BELOW_TARGET and item.target and l.unit_price > item.target:
            continue
        pool = "res" if l.is_resale else "new"
        st = stats.get(pool)
        market_ok = bool(st and st["n"] >= MARKET_MIN_SAMPLE)
        reasons, ref, label, rtype = [], None, "", ""
        if pool == "new":
            if l.pack_qty > 1 and l.pack_qty in (pack_baselines or {}):
                ref, label, rtype = pack_baselines[l.pack_qty], f"verified {l.pack_qty}-pack price", "verified_pack"
            elif ref_price:
                ref, label, rtype = ref_price, REF_LABELS.get(ref_type, ref_type), ref_type
            elif baseline:
                ref, label, rtype = baseline, "verified price", "verified_current"
            elif market_ok:
                ref, label, rtype = st["avg"], "market avg (unverified)", "market_average"
        else:
            if market_ok:
                ref, label, rtype = st["avg"], "resale avg", "market_average"
            elif ref_price and l.unit_price <= ref_price * (1 - RESALE_VS_NEW_DISCOUNT):
                reasons.append(f">={RESALE_VS_NEW_DISCOUNT:.0%} below new-condition {REF_LABELS.get(ref_type, ref_type)} ${ref_price:.2f}")
                l.ref_price, l.ref_type = ref_price, ref_type
        if ref and l.unit_price >= ref * OUTLIER_LOW:
            if l.unit_price <= ref * (1 - DEAL_DISCOUNT):
                reasons.append(f">={DEAL_DISCOUNT:.0%} below {label} ${ref:.2f}")
            l.ref_price, l.ref_type = ref, rtype
        trend = hist.get("median") if pool == "new" else resale_trend
        sane = ref is None or l.unit_price >= ref * OUTLIER_LOW
        if trend and sane and l.unit_price < trend * (1 - TREND_MIN_DISCOUNT):
            reasons.append(f"below 30d {'verified median' if pool == 'new' else 'trend'} ${trend:.2f}")
            l.ref_price = l.ref_price or trend
            l.ref_type = l.ref_type or ("historical_verified" if pool == "new" else "resale_trend")
        if pool == "new" and hist.get("low") and hist.get("median_n", 0) >= 3 and sane and l.unit_price <= hist["low"] - 0.01:
            reasons.append(f"new verified low (prev ${hist['low']:.2f})")
        if USE_TARGET_RULE and item.target and l.unit_price <= item.target * (1 - DEAL_DISCOUNT):
            reasons.append(f">=15% below target ${item.target:.2f}")
        if reasons:
            rule = "; ".join(reasons) + f" [{'resale' if l.is_resale else 'new'} pool]"
            if l.conditional:
                rule = f"CONDITIONAL ({l.conditional}) " + rule
            if l.evidence not in VERIFIED:
                rule += " [market snapshot - confirm on merchant page]"
            l.deal_rule = rule
            deals.append(l)
    deals.sort(key=lambda l: (bool(l.conditional), l.evidence not in VERIFIED,
                              0 if (l.is_primary and not l.is_resale) else 1 if not l.is_resale else 2, l.unit_price))
    return deals[:MAX_DEALS_PER_ITEM]
