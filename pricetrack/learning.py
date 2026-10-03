"""Self-improvement: what each run teaches the next one.

Everything here is derived from data the run already has (retrieval outcomes, scored listings, verified pages) and is
kept in tracker_state/state.json, so it needs no extra network calls or SerpApi credits.

  vendor yield        per retailer site: how often a site search / sitemap crawl actually produced a listing.
                      A retailer that never delivered after VENDOR_SKIP_MIN_ATTEMPTS attempts across
                      VENDOR_SKIP_MIN_ITEMS different items is skipped (retried every VENDOR_RETRY_DAYS);
                      Google Shopping (SerpApi) still covers it.
  suggested vendors   merchants that keep showing up with High-confidence new-condition listings but are on neither
                      vendor list. Reported at the end of every run with their stats and the domain to paste into the
                      'Primary Vendor List' tab - never added automatically.
  sheet suggestions   per item: identifiers the verified pages report but the Master Sheet lacks (GTIN/MPN/Brand),
                      name / spec words that no verified page title contains, size-unit slips (16.4" vs 16.4 ft),
                      links that failed, vendors worth a Product URL.
  learned identity    GTIN / MPN / brand read from VERIFIED, High-confidence, single-unit pages are remembered
                      (State.learn) so the next run matches those pages by identifier instead of by title.
"""
from __future__ import annotations

import re
from datetime import timedelta
from typing import Optional

from .models import MARKET_SNAPSHOT, Outcome, VERIFIED
from .text import keys_match, norm_text, vendor_key
from .urls import host_of

# ---- settings (price_tracker.py pushes its own values over these at start-up) ------------------------------
VENDOR_SKIP_ENABLED = True
VENDOR_SKIP_MIN_ATTEMPTS = 8       # attempts (item x run) with no listing ever found ...
VENDOR_SKIP_MIN_ITEMS = 3          # ... across at least this many different items
VENDOR_RETRY_DAYS = 30             # a skipped retailer is probed again this long after its last attempt
SUGGEST_MIN_HIGH = 3               # High-confidence, new-condition listings seen (counted once per run x item)
SUGGEST_MIN_ITEMS = 2              # ... for at least this many different wishlist items
SUGGEST_MIN_RUNS = 2               # ... in at least this many different runs
SUGGEST_MAX_PRICE_RATIO = 1.05     # ... and its best price is within this of the verified price (when one exists)
SUGGEST_IGNORE: set = {"bh", "bhphotovideo", "lowes"}   # vendor keys never suggested (removed on purpose)
SUGGEST_MAX_WATCHING = 6
LEARN_MAX_MPNS = 6                 # cap on MPNs remembered per item

_LOG_KEEP = 40
_OBS_KEEP = 80


def _iso(dt) -> str:
    return dt.replace(microsecond=0).isoformat()


# =============================================================================
# 1. Vendor yield
# =============================================================================

def _stats(state) -> dict:
    return state.data.setdefault("vendor_stats", {})


def record_vendor_attempt(state, domain: str, wid, run_id: str, outcome: str, n_listings: int) -> None:
    """One site-search / crawl attempt for (retailer, item, run). Idempotent per (run, item)."""
    if not domain or outcome in (Outcome.SKIPPED, Outcome.NETWORK, Outcome.PARSER, Outcome.API):
        return                                                 # not evidence about whether the retailer carries things
    kind = ("found" if n_listings and outcome in (Outcome.SUCCESS, Outcome.UNAVAILABLE) else
            "blocked" if outcome == Outcome.BLOCKED else "empty")
    with state.lock:
        st = _stats(state).setdefault(domain, {"log": {}, "found_total": 0, "first": _iso(state.now)})
        key = f"{run_id}|{str(wid).strip()}"
        if st["log"].get(key, {}).get("o") == "found":
            st["found_total"] = max(0, st.get("found_total", 0) - 1)
        st["log"][key] = {"o": kind, "ts": _iso(state.now)}
        if kind == "found":
            st["found_total"] = st.get("found_total", 0) + 1
            st["last_found"] = _iso(state.now)
        st["last_try"] = _iso(state.now)
        while len(st["log"]) > _LOG_KEEP:
            st["log"].pop(next(iter(st["log"])))


def vendor_skip_reason(state, domain: str) -> str:
    """'' = search it; otherwise why a retailer is skipped this run."""
    if not VENDOR_SKIP_ENABLED:
        return ""
    st = _stats(state).get(domain)
    if not st or st.get("found_total", 0) > 0:
        return ""                                              # a retailer that has ever delivered is never skipped
    log = st.get("log", {})
    items = {k.split("|", 1)[1] for k in log if "|" in k}
    if len(log) < VENDOR_SKIP_MIN_ATTEMPTS or len(items) < VENDOR_SKIP_MIN_ITEMS:
        return ""
    if state.age_days(st.get("last_try")) >= VENDOR_RETRY_DAYS:
        return ""                                              # due for a probe
    blocked = sum(1 for v in log.values() if v["o"] == "blocked")
    retry = (state.now + timedelta(days=max(0, VENDOR_RETRY_DAYS - state.age_days(st.get("last_try"))))).strftime("%Y-%m-%d")
    return (f"no yield: 0 listings in {len(log)} attempts across {len(items)} items"
            + (f" ({blocked} blocked)" if blocked else "") + f"; skipped until ~{retry} (Google Shopping still covers it)")


# =============================================================================
# 2. Suggested vendors
# =============================================================================

def vendor_display(raw: str) -> tuple:
    """'Newegg.com - Walts TV' -> ('Newegg.com', True): marketplace sellers are reported as the marketplace."""
    name = re.sub(r"^from\s+", "", raw or "", flags=re.I).strip()
    base = name.split(" - ")[0].strip() or name
    return base, base != name


def domain_guess(raw: str, hosts: dict) -> tuple:
    """(domain, how). From a merchant link when Google gave one, else from the vendor name, else a labelled guess."""
    if hosts:
        return max(hosts.items(), key=lambda kv: kv[1])[0], "merchant link"
    base, _ = vendor_display(raw)
    m = re.search(r"([a-z0-9][a-z0-9-]*(?:\.[a-z0-9-]+)*\.(?:com|net|org|co|us|shop|store))\b", base.lower())
    if m:
        return m.group(1), "vendor name"
    key = vendor_key(base)
    return (f"{key}.com", "guess - confirm") if key else ("", "unknown")


def record_vendor_candidates(state, item, listings: list, run_id: str, secondary_keys: list, reference: Optional[float]) -> None:
    """Remember merchants (not on either vendor list) that carry this product at High confidence."""
    wid = str(item.wid).strip()
    best = {}
    for l in listings:
        if (not l.vkey or l.is_primary or l.is_resale or l.condition != "new" or l.confidence not in ("High", "Medium")
                or l.vendor.startswith("Google Shopping (") or not l.eligible
                or any(keys_match(l.vkey, k) for k in secondary_keys) or l.vkey == "ebay"):
            continue
        cur = best.get(l.vkey)
        rank = (l.confidence == "High", l.evidence in VERIFIED, -l.unit_price)
        if cur is None or rank > cur[0]:
            best[l.vkey] = (rank, l)
    if not best:
        return
    with state.lock:
        cands = state.data.setdefault("vendor_candidates", {})
        for vk, (_, l) in best.items():
            c = cands.setdefault(vk, {"names": {}, "hosts": {}, "obs": {}, "first": _iso(state.now)})
            c["names"][l.vendor] = c["names"].get(l.vendor, 0) + 1
            for u in (l.url, l.alt_url):
                h = host_of(u) if u else ""
                if h and "google." not in h:
                    c["hosts"][h] = c["hosts"].get(h, 0) + 1
                    break
            key = f"{run_id}|{wid}"
            c["obs"][key] = {"conf": l.confidence, "unit": round(l.unit_price, 2), "verified": l.evidence in VERIFIED,
                             "ref": round(reference, 2) if reference else None, "ts": _iso(state.now)}
            c["last"] = _iso(state.now)
            while len(c["obs"]) > _OBS_KEEP:
                c["obs"].pop(next(iter(c["obs"])))


def candidate_report(state, primary: list, secondary_keys: list) -> tuple:
    """(qualified, watching): lists of dicts with the stats shown at the end of a run."""
    out = []
    for vk, c in state.data.get("vendor_candidates", {}).items():
        if vk in SUGGEST_IGNORE or any(keys_match(vk, v.key) for v in primary) or any(keys_match(vk, k) for k in secondary_keys):
            continue
        obs = c.get("obs", {})
        high = {k: o for k, o in obs.items() if o.get("conf") == "High"}
        if not high:
            continue
        items = {k.split("|", 1)[1] for k in high}
        runs = {k.split("|", 1)[0] for k in high}
        ratios = [o["unit"] / o["ref"] for o in high.values() if o.get("ref")]
        raw = max(c.get("names", {"": 0}).items(), key=lambda kv: kv[1])[0]
        name, marketplace = vendor_display(raw)
        domain, how = domain_guess(raw, c.get("hosts", {}))
        crit = {
            f"{SUGGEST_MIN_HIGH}+ High-confidence listings": len(high) >= SUGGEST_MIN_HIGH,
            f"{SUGGEST_MIN_ITEMS}+ different items": len(items) >= SUGGEST_MIN_ITEMS,
            f"{SUGGEST_MIN_RUNS}+ runs": len(runs) >= SUGGEST_MIN_RUNS,
            f"best price within {SUGGEST_MAX_PRICE_RATIO - 1:.0%} of the verified price":
                (not ratios) or min(ratios) <= SUGGEST_MAX_PRICE_RATIO,
        }
        out.append({"key": vk, "name": name, "seen_as": sorted(c.get("names", {}))[:3], "marketplace": marketplace,
                    "domain": domain, "domain_how": how, "high": len(high), "items": len(items), "runs": len(runs),
                    "verified": sum(1 for o in high.values() if o.get("verified")),
                    "best": min(o["unit"] for o in high.values()),
                    "best_ratio": min(ratios) if ratios else None,
                    "avg_ratio": (sum(ratios) / len(ratios)) if ratios else None,
                    "criteria": crit, "ok": all(crit.values())})
    qualified = sorted((c for c in out if c["ok"]), key=lambda c: (-c["high"], -c["items"], c["name"]))
    watching = sorted((c for c in out if not c["ok"]), key=lambda c: (-sum(c["criteria"].values()), -c["high"]))
    return qualified, watching[:SUGGEST_MAX_WATCHING]


def _fmt_ratio(r) -> str:
    return f"{r:.2f}x" if r else "n/a"


def _cand_lines(c: dict, qualified: bool) -> list:
    stats = (f"{c['high']} High-confidence listings across {c['items']} item(s) in {c['runs']} run(s); best ${c['best']:,.2f} "
             f"({_fmt_ratio(c['best_ratio'])} the verified price, avg {_fmt_ratio(c['avg_ratio'])}); "
             f"{c['verified']} verified on the merchant's own page")
    lines = [f"  * {c['name']}" + (f"  [marketplace seller; seen as {', '.join(c['seen_as'])}]" if c["marketplace"] else ""),
             f"      stats : {stats}"]
    if qualified:
        lines.append(f"      sheet : Primary Vendor List -> Vendor = {c['name']} | isDirectToConsumer = 0 | "
                     f"Domain = {c['domain']}   ({c['domain_how']})")
    else:
        miss = [k for k, v in c["criteria"].items() if not v]
        lines.append(f"      needs : {'; '.join(miss)}   (domain would be {c['domain']}, {c['domain_how']})")
    return lines


# =============================================================================
# 3. Master Sheet suggestions (per item)
# =============================================================================

_UNIT_FT = re.compile(r"(\d+(?:\.\d+)?)\s*(?:ft|feet|foot)\b", re.I)
_NAME_INCH = re.compile(r"(\d+(?:\.\d+)?)\s*(?:\"|”|″|in\b|inch(?:es)?\b)", re.I)


def _upc(g: str) -> str:
    g = g.lstrip("0")
    return g.zfill(12) if len(g) <= 12 else g


def sheet_suggestions(item, listings: list, failed: list) -> list:
    """Human-readable fixes for the Master Sheet row, from what the verified pages actually say."""
    out = []
    ver = [l for l in listings if l.evidence in VERIFIED and l.confidence in ("High", "Medium") and l.condition == "new"]
    high1 = [l for l in ver if l.confidence == "High" and l.pack_qty == 1]
    # identifiers the pages report but the sheet lacks
    page_g = {g for l in high1 for g in l.gtins}
    page_m = {m for l in high1 for m in l.mpns if len(str(m)) >= 5 and any(c.isdigit() for c in str(m))}
    page_b = next((l.brand for l in high1 if l.brand), "")
    if page_g and not item.gtins:
        out.append(f"GTIN/UPC is blank - verified pages report {', '.join(sorted(_upc(g) for g in page_g)[:2])}")
    sheet_m = {norm_text(m).replace(" ", "") for m in item.mpns} | {norm_text(m).replace(" ", "") for m in item.skus}
    new_m = sorted(m for m in page_m if norm_text(m).replace(" ", "") not in sheet_m)
    if new_m and not (item.mpns or item.skus):
        out.append(f"MPN/Model is blank - verified pages report {', '.join(new_m[:3])}")
    elif new_m:
        out.append(f"verified pages also report model code(s) {', '.join(new_m[:3])} (not in your MPN/Model cell)")
    if page_b and not item.brand:
        out.append(f"Brand is blank - verified pages report '{page_b}'")
    # words / specs / units that no verified page title contains
    if ver:
        titles = [norm_text(l.title) for l in ver]
        tokens = set(" ".join(titles).split())
        compact = " ".join(titles).replace(" ", "")
        brand = norm_text(item.brand or item.learned_brand)
        gone = [t for t in norm_text(item.product).split()
                if len(t) >= 3 and t not in tokens and t not in compact and t not in brand.split()]
        if gone:
            out.append(f"Product name word(s) {', '.join(repr(t) for t in gone[:4])} appear in none of the {len(ver)} verified page "
                       f"title(s) - they can hold a page at Medium; consider dropping them from the name")
        for s in item.spec_phrases:
            toks = norm_text(s).split()
            if toks and not all(t in tokens or (len(t) >= 5 and t in compact) for t in toks):
                out.append(f"spec/keyword '{s}' is in none of the verified page titles - it can hold a page at Medium")
        ft = {m.group(1) for l in ver for m in _UNIT_FT.finditer(l.title)}
        for m in _NAME_INCH.finditer(item.product):
            if m.group(1) in ft:
                out.append(f"Product name says {m.group(0).strip()} (inches) but verified pages say {m.group(1)} ft - "
                           f"write '{m.group(1)}ft' in the name")
    # pages stuck below High / failed links
    stuck = [l for l in listings if l.from_url and l.evidence in VERIFIED and l.confidence == "Medium"]
    for l in stuck[:2]:
        out.append(f"Product URL {host_of(l.url)} is only Medium confidence ({l.conf_reason[:90]})")
    for u, host, vname, outcome in failed:
        if outcome in (Outcome.IDENTITY_MISMATCH, Outcome.PARSER, Outcome.NO_MATCH):
            out.append(f"Product URL at {host} failed ({outcome}) - check the link still shows this exact product")
    # vendors that show a High price only via Google: a Product URL would verify them
    have = {vendor_key(host_of(u)) for u in item.urls}
    snap = {}
    for l in listings:
        if l.evidence == MARKET_SNAPSHOT and l.confidence == "High" and l.condition == "new" and l.eligible and l.vkey \
                and l.vkey not in have and not l.vendor.startswith("Google Shopping ("):
            if l.vkey not in snap or l.unit_price < snap[l.vkey].unit_price:
                snap[l.vkey] = l
    pri = sorted((l for l in snap.values() if l.is_primary), key=lambda l: l.unit_price)[:2]
    for l in pri:
        out.append(f"{l.vendor} shows ${l.unit_price:,.2f} only via Google Shopping - add its product page to Product URLs to verify it")
    if not (item.gtins or item.learned_gtins or item.all_mpns):
        out.append("no GTIN or MPN known yet - add either so pages are matched by identifier, not by title")
    return list(dict.fromkeys(out))


def store_suggestions(state, item, lines: list) -> None:
    with state.lock:
        b = state._item_bucket("suggestions", item)
        b["lines"], b["ts"] = lines, _iso(state.now)


# =============================================================================
# 4. Learned identity from verified pages
# =============================================================================

def learn_from_listings(state, item, listings: list) -> list:
    """Remember GTIN / MPN / brand of VERIFIED, High-confidence, single-unit, ordinary-priced pages."""
    new = []
    for l in listings:
        if not (l.evidence in VERIFIED and l.confidence == "High" and l.pack_qty == 1 and not l.conditional
                and l.condition == "new" and (l.from_url or l.source in ("page", "discovered", "shopify", "browser"))):
            continue
        mpns = {m for m in l.mpns if m}
        known = {m.upper() for m in list(item.all_mpns)}
        mpns = {m for m in mpns if m.upper() not in known}
        if len(item.learned_mpns) + len(mpns) > LEARN_MAX_MPNS:
            mpns = set()
        new += state.learn(item, l.gtins, mpns, l.brand, l.url)
    return list(dict.fromkeys(new))


# =============================================================================
# 5. End-of-run report
# =============================================================================

def build_report(state, primary: list, secondary_keys: list, items: list, run_id: str) -> list:
    """Text lines printed at the end of the run (and saved as tracker_state/suggestions.md)."""
    lines = []
    qualified, watching = candidate_report(state, primary, secondary_keys)
    lines.append("SUGGESTED VENDORS - confirm before adding (never added automatically)")
    lines.append(f"  Qualifies when: {SUGGEST_MIN_HIGH}+ High-confidence new listings, {SUGGEST_MIN_ITEMS}+ different items, "
                 f"{SUGGEST_MIN_RUNS}+ runs, best price within {SUGGEST_MAX_PRICE_RATIO - 1:.0%} of the verified price; "
                 f"vendor on neither vendor list.")
    if qualified:
        for c in qualified:
            lines += _cand_lines(c, True)
    else:
        lines.append("  (none qualify yet)")
    if watching:
        lines.append("  Watching (not yet qualified):")
        for c in watching:
            lines += _cand_lines(c, False)
    lines.append("")
    lines.append("MASTER SHEET SUGGESTIONS")
    shown = False
    for it in items:
        b = state.data.get("suggestions", {}).get(str(it.wid).strip())
        if b and b.get("product") == norm_text(it.product) and b.get("lines"):
            shown = True
            lines.append(f"  [{it.wid}] {it.product}")
            lines += [f"      - {x}" for x in b["lines"]]
    if not shown:
        lines.append("  (nothing to suggest)")
    lines.append("")
    lines.append("RETAILER SEARCH YIELD (site search / sitemap crawl)")
    rows = []
    for dom, st in sorted(_stats(state).items()):
        log = st.get("log", {})
        n = len(log)
        found = sum(1 for v in log.values() if v["o"] == "found")
        blocked = sum(1 for v in log.values() if v["o"] == "blocked")
        why = vendor_skip_reason(state, dom)
        rows.append(f"  {dom:<24} attempts {n:>3} | found {found:>3} | blocked {blocked:>3} | "
                    + ("SKIPPED - " + why if why else "active"))
    lines += rows or ["  (no data yet)"]
    return lines
