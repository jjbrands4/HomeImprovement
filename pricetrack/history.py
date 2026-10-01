"""
Persistent tracker state + verified-only historical statistics.

tracker_state/state.json        (small, rewritten atomically each run)
    identity    learned GTIN/MPN/brand per wishlist item (from validated pages / official APIs)
    discovery   validated merchant URL per (item, retailer) + negative results with timestamps
    shopify     per-domain "is this a Shopify store" detection
    serp_cache  SerpApi responses (<= SERP_CACHE_HOURS old) so retries / re-runs spend no credits
    serp_last   when each item last used SerpApi discovery
    http_cache  ETag / Last-Modified validators + parsed results for conditional GETs
    strategies  per-domain client/extractor that worked last time
    circuits    per-domain cool-downs after repeated blocks
    offers      last known state of every stable offer id (price, stock, since) -> change events
tracker_state/observations.jsonl  one row per (run, item, offer) for trusted + candidate offers

Idempotency: observation rows and offer states are keyed by run id, so re-running the same run id
(e.g. a GitHub Actions re-run) replaces its own rows instead of duplicating them, and offers compare
against the state *before* that run - so an unchanged offer is never reported as a new price event.
"""
from __future__ import annotations

import json
import os
import statistics
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from .models import Item, Listing, utcnow
from .text import norm_text

STATE_DIR = "tracker_state"
STATE_FILE = "state.json"
OBS_FILE = "observations.jsonl"
SHOPIFY_TTL_DAYS = 30
SERP_KEEP_HOURS = 48
OBS_KEEP_DAYS = 400
OFFER_KEEP_DAYS = 180
PRICE_EPSILON = 0.01             # an offer whose unit price moved less than 1 cent is 'unchanged'
MATERIAL_CHANGE_PCT = 0.02       # Run Data 'Change vs Prior' is flagged material at >= 2%
EWMA_ALPHA = 0.3


def _iso(dt: datetime) -> str:
    return dt.replace(microsecond=0).isoformat()


def _dt(s) -> Optional[datetime]:
    if isinstance(s, datetime):
        return s
    try:
        return datetime.fromisoformat(str(s))
    except (TypeError, ValueError):
        return None


class State:
    def __init__(self, folder: Optional[Path] = None, now: Optional[datetime] = None):
        self.folder = Path(folder) if folder else None
        self.now = now or utcnow()
        self.data: dict = {}
        self._obs: Optional[list] = None
        self.lock = threading.RLock()          # adapters run in a small thread pool
        self.load()

    # ---- persistence ----------------------------------------------------------------------------
    def load(self) -> None:
        d = {}
        if self.folder and (self.folder / STATE_FILE).exists():
            try:
                d = json.loads((self.folder / STATE_FILE).read_text(encoding="utf-8"))
            except (OSError, ValueError):
                d = {}
        for k in ("identity", "discovery", "shopify", "serp_cache", "serp_last", "http_cache", "strategies",
                  "circuits", "offers"):
            d.setdefault(k, {})
        d["version"] = 1
        self.data = d

    def observations(self) -> list:
        if self._obs is None:
            self._obs = []
            p = self.folder / OBS_FILE if self.folder else None
            if p and p.exists():
                for line in p.read_text(encoding="utf-8").splitlines():
                    try:
                        self._obs.append(json.loads(line))
                    except ValueError:
                        continue
        return self._obs

    def save(self) -> None:
        if not self.folder:
            return
        self.folder.mkdir(parents=True, exist_ok=True)
        cut = _iso(self.now - timedelta(hours=SERP_KEEP_HOURS))
        self.data["serp_cache"] = {k: v for k, v in self.data["serp_cache"].items() if v.get("ts", "") >= cut}
        stale = _iso(self.now - timedelta(days=OFFER_KEEP_DAYS))
        self.data["offers"] = {k: v for k, v in self.data["offers"].items() if v.get("ts", "") >= stale}
        self._atomic(self.folder / STATE_FILE, json.dumps(self.data, indent=1, sort_keys=True, default=str))
        if self._obs is not None:
            keep = _iso(self.now - timedelta(days=OBS_KEEP_DAYS))
            rows = [o for o in self._obs if o.get("ts", "") >= keep]
            self._atomic(self.folder / OBS_FILE, "".join(json.dumps(o, sort_keys=True) + "\n" for o in rows))

    @staticmethod
    def _atomic(path: Path, text: str) -> None:
        tmp = path.with_name(path.name + ".tmp")
        tmp.write_text(text, encoding="utf-8")
        os.replace(tmp, path)

    def age_days(self, ts) -> float:
        d = _dt(ts)
        return (self.now - d).total_seconds() / 86400 if d else 1e9

    # ---- item keys ------------------------------------------------------------------------------
    @staticmethod
    def item_key(item: Item) -> str:
        return str(item.wid).strip().replace(".0", "")

    def _item_bucket(self, section: str, item: Item) -> dict:
        """Per-item bucket; reset automatically when the Master Sheet product name changes."""
        with self.lock:
            k = self.item_key(item)
            b = self.data[section].get(k)
            if not b or b.get("product") != norm_text(item.product):
                b = {"product": norm_text(item.product)}
                self.data[section][k] = b
            return b

    # ---- learned identity -----------------------------------------------------------------------
    def apply_learned(self, item: Item) -> None:
        b = self._item_bucket("identity", item)
        item.learned_gtins = set(b.get("gtins", []))
        item.learned_mpns = set(b.get("mpns", []))
        item.learned_brand = b.get("brand", "")

    def learn(self, item: Item, gtins=(), mpns=(), brand: str = "", source: str = "") -> list:
        """Remember identifiers from a VALIDATED page/API record. Returns the newly learned ones."""
        with self.lock:
            b = self._item_bucket("identity", item)
            new = []
            g0, m0 = set(b.get("gtins", [])), set(b.get("mpns", []))
            for g in gtins:
                if g and g not in g0:
                    g0.add(g)
                    new.append(f"GTIN {g.lstrip('0')}")
            for m in mpns:
                m = str(m).strip()
                if m and len(m) >= 4 and m.upper() not in {x.upper() for x in m0} and any(c.isdigit() for c in m):
                    m0.add(m)
                    new.append(f"MPN {m}")
            if new or (brand and not b.get("brand")):
                b.update({"gtins": sorted(g0), "mpns": sorted(m0), "brand": b.get("brand") or brand,
                          "updated": _iso(self.now)})
                src = b.setdefault("sources", [])
                if source and source not in src:
                    src.append(source)
                    del src[:-5]
            item.learned_gtins, item.learned_mpns = set(g0), set(m0)
            item.learned_brand = b.get("brand", "")
            return new

    # ---- discovery cache ------------------------------------------------------------------------
    def discovery_get(self, item: Item, domain: str) -> Optional[dict]:
        return self._item_bucket("discovery", item).get(domain)

    def discovery_put(self, item: Item, domain: str, url: str = "", ok: bool = True, detail: str = "",
                      item_id: str = "") -> None:
        with self.lock:
            b = self._item_bucket("discovery", item)
            prev = b.get(domain) or {}
            b[domain] = {"url": url or (prev.get("url") if ok else ""), "ok": ok, "detail": detail[:160],
                         "ts": _iso(self.now), "item_id": item_id or (prev.get("item_id") if ok else "")}

    def known_urls(self, item: Item) -> dict:
        return {d: v["url"] for d, v in self._item_bucket("discovery", item).items()
                if isinstance(v, dict) and v.get("ok") and v.get("url")}

    # ---- shopify detection ----------------------------------------------------------------------
    def shopify_get(self, domain: str) -> Optional[bool]:
        v = self.data["shopify"].get(domain)
        if v and self.age_days(v.get("ts")) <= SHOPIFY_TTL_DAYS:
            return bool(v.get("is"))
        return None

    def shopify_put(self, domain: str, is_shopify: bool) -> None:
        with self.lock:
            self.data["shopify"][domain] = {"is": bool(is_shopify), "ts": _iso(self.now)}

    # ---- SerpApi cache --------------------------------------------------------------------------
    def serp_get(self, query: str, max_hours: float) -> Optional[list]:
        v = self.data["serp_cache"].get(query)
        if v and self.age_days(v.get("ts")) * 24 <= max_hours:
            return v.get("rows")
        return None

    def serp_put(self, query: str, rows: list) -> None:
        with self.lock:
            self.data["serp_cache"][query] = {"ts": _iso(self.now), "rows": rows}

    def serp_mark(self, item: Item) -> None:
        with self.lock:
            self.data["serp_last"][self.item_key(item)] = _iso(self.now)

    def serp_age_days(self, item: Item) -> float:
        return self.age_days(self.data["serp_last"].get(self.item_key(item)))

    # ---- offers: change events (idempotent) -----------------------------------------------------
    def offer_event(self, l: Listing, item: Item, run_id: str) -> str:
        with self.lock:
            offers = self.data["offers"]
            st = offers.get(l.offer_id)
            if st and st.get("run") == run_id:
                prev = st.get("prev")                       # re-run of the same run: compare with pre-run state
            else:
                prev = {k: st.get(k) for k in ("unit", "ts", "since", "in_stock")} if st else None
            unit = round(l.unit_price, 2)
            if not prev or prev.get("unit") is None:
                event, since = "new offer", _iso(self.now)
            else:
                d = unit - float(prev["unit"])
                if abs(d) < PRICE_EPSILON:
                    event, since = f"unchanged since {str(prev.get('since') or prev.get('ts'))[:10]}", prev.get("since") or prev.get("ts")
                else:
                    event, since = (f"price drop from ${prev['unit']:,.2f}" if d < 0 else f"price increase from ${prev['unit']:,.2f}"), _iso(self.now)
                if prev.get("in_stock") is False and l.in_stock is not False:
                    event += "; back in stock"
            offers[l.offer_id] = {"wid": self.item_key(item), "unit": unit, "price": l.price, "ts": _iso(self.now),
                                  "since": since, "in_stock": l.in_stock, "evidence": l.evidence, "vendor": l.vendor,
                                  "url": l.url, "run": run_id, "prev": prev}
            return event

    # ---- observations (idempotent) --------------------------------------------------------------
    def record_observations(self, run_id: str, item: Item, listings: list) -> int:
        obs = self.observations()
        wid = self.item_key(item)
        obs[:] = [o for o in obs if not (o.get("run") == run_id and o.get("wid") == wid)]
        n = 0
        for l in listings:
            if not (l.trusted or l.eligible or l.from_url):
                continue
            obs.append({"run": run_id, "ts": _iso(self.now), "wid": wid, "offer": l.offer_id, "fp": l.fingerprint,
                        "vendor": l.vendor, "source": l.source, "evidence": l.evidence, "conf": l.confidence,
                        "match": l.match_evidence, "trusted": bool(l.trusted), "condition": l.condition,
                        "availability": l.availability or ({True: "in_stock", False: "out_of_stock"}.get(l.in_stock, "")),
                        "pack": l.pack_qty, "price": l.price, "shipping": l.shipping, "effective": l.effective_price,
                        "unit": l.unit_price, "regular": l.regular_price, "conditional": l.conditional,
                        "url": l.url, "event": l.price_event})
            n += 1
        return n

    def trusted_units(self, item: Item, exclude_run: str = "", days: int = 365) -> list:
        """Verified, High-confidence, single-unit new observations: [(ts, unit)]."""
        wid, cut = self.item_key(item), _iso(self.now - timedelta(days=days))
        return [(o["ts"], o["unit"]) for o in self.observations()
                if o.get("wid") == wid and o.get("trusted") and o.get("run") != exclude_run and o.get("ts", "") >= cut
                and o.get("condition", "new") == "new" and o.get("unit")]


# =============================================================================
# Historical statistics (verified only)
# =============================================================================

def verified_series(history_rows: list, wid, exclude_run: str = "", trust_legacy: bool = False) -> list:
    """Per-run verified market price for one item, oldest first, from Run Data rows whose
    'Baseline Source' says the baseline came from verified listings."""
    out = []
    for h in history_rows:
        if str(h.get("wid")).strip().replace(".0", "") != str(wid).strip().replace(".0", ""):
            continue
        if exclude_run and h.get("run") == exclude_run:
            continue
        src = str(h.get("baseline_src") or "")
        if h.get("baseline") and (src.startswith("verified") or (trust_legacy and src.startswith("Product URLs"))):
            out.append((h["dt"], float(h["baseline"])))
    return sorted(out)


def history_stats(series: list, trusted_units: list, now: datetime, window_days: int = 30,
                  min_points: int = 3) -> dict:
    """prior verified price, trailing verified median, historical verified low, EWMA."""
    st = {"prior": None, "prior_dt": None, "median": None, "median_n": 0, "low": None, "ewma": None}
    if series:
        st["prior_dt"], st["prior"] = series[-1]
        recent = [v for d, v in series if d >= now - timedelta(days=window_days)]
        st["median_n"] = len(recent)
        if len(recent) >= min_points:
            st["median"] = round(statistics.median(recent), 2)
        e = None
        for _, v in series:
            e = v if e is None else EWMA_ALPHA * v + (1 - EWMA_ALPHA) * e
        st["ewma"] = round(e, 2) if e is not None else None
    lows = [v for _, v in series] + [u for _, u in trusted_units]
    if lows:
        st["low"] = round(min(lows), 2)
    return st
