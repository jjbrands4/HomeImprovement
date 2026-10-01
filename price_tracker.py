#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
price_tracker.py - Home wishlist price tracker
==============================================

WHAT IT DOES
  1. Reads the wishlist workbook (Master Sheet + Primary/Secondary Vendor List tabs).
  2. For each selected wishlist item it gathers live prices from:
       - Direct-to-consumer vendor sites (only when the vendor name loosely matches the
         product name, e.g. "Sonos Ray" -> "Sonos Direct")              [free, no key]
       - Pipeline A: Google Shopping via SerpApi (thousands of retailers)  [SERPAPI_KEY]
       - Pipeline B: eBay Browse API (used / refurbished + seller ratings) [EBAY_* keys]
         (only for items where "Open to used or high quality refurbished?" = Yes)
  3. Scores how confident we are that each listing is the *exact* product.
  4. Flags deals, records everything in the "Run Data" and "Deals Data" tabs.
     NOTHING else in the workbook is modified.

DEAL DEFINITION (per your instructions)
  A listing is a deal when its per-item price is
      (A) >= 15% below this run's average market price   OR
      (B) below the trailing-30-day average of previous runs,
  where the averages are computed SEPARATELY for two price pools:
      - "New"    : new-condition listings (primary vendors + any unlisted new retailer)
      - "Resale" : secondary-list vendors (eBay, woot!, ...) and any used/refurbished listing
  so resale prices never skew the new-price baseline.

NEVER CRASHES ON A SOURCE FAILURE
  Every source is wrapped in try/except. Missing keys, exhausted SerpApi credits, eBay
  errors or blocked sites just produce "N/A" in the Source Notes column of Run Data.

USAGE EXAMPLES
  python price_tracker.py --dry-run --force          # run everything, write nothing
  python price_tracker.py --min-priority 4           # only priority 4-5 items
  python price_tracker.py --locations "Master Bedroom"
  python price_tracker.py --items 1,3 --force        # specific WishlistItem ids

All timestamps written to the workbook are UTC.
"""

from __future__ import annotations

import argparse
import base64
import difflib
import json
import os
import re
import statistics
import sys
import time
import uuid
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional
from urllib.parse import unquote, urljoin, urlparse

import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# =============================================================================
# 1. SETTINGS  (safe to tweak; everything else reads from here)
# =============================================================================

WORKBOOK_GLOB = "Home*Wishlist*.xlsx"      # finds "Home Wishlist.xlsx" or "Home_Wishlist.xlsx"

# ---- Deal rules -------------------------------------------------------------
DEAL_DISCOUNT = 0.15          # rule A: price <= (1 - 0.15) * average market price of its pool
TREND_MIN_DISCOUNT = 0.0      # rule B: price < 30d trend * (1 - this). 0.0 = literal "below the trend".
                              #   Raise to 0.05 if too many listings qualify as deals.
TREND_WINDOW_DAYS = 30        # look-back window for the trend (uses RowDateTime in Run Data)
TREND_MIN_POINTS = 3          # need at least this many past runs before rule B is used
MARKET_MIN_SAMPLE = 3         # need at least this many listings in a pool before rule A is used
USE_TARGET_RULE = False       # True = ALSO treat "15% below your Target Price" as a deal
REQUIRE_AT_OR_BELOW_TARGET = False   # True = never report a "deal" priced above your Target Price

# ---- Noise / sanity filters -------------------------------------------------
MIN_PRICE_RATIO_OF_TARGET = 0.25   # listings priced under 25% of target are almost surely accessories
OUTLIER_LOW, OUTLIER_HIGH = 0.4, 2.5   # prices outside 0.4x..2.5x the pool median are excluded from averages
MAX_DEALS_PER_ITEM = 5             # cap Deals Data rows per item per run

# ---- Cost control -----------------------------------------------------------
# Minimum days between checks per Priority. 0 = check on every scheduled run.
# Use --force to ignore. (Workflow cron is every 3 days, so 0 and 3 behave the same.)
CADENCE_DAYS = {5: 0, 4: 0, 3: 0, 2: 6, 1: 9}

# ---- Matching ---------------------------------------------------------------
HEAD_MIN_CHARS = 70           # only the first N characters of a listing title are compared to the
HEAD_EXTRA_CHARS = 40         #   product name: max(70, len(name) + 40). Long SEO tails are ignored.
QUERY_INCLUDE_SPECS = True    # add "Product specifications" text to the search query

# ---- Networking -------------------------------------------------------------
HTTP_TIMEOUT = 15
USER_AGENT = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/124.0 Safari/537.36")
SERPAPI_SEARCH_URL = "https://serpapi.com/search.json"
SERPAPI_ACCOUNT_URL = "https://serpapi.com/account.json"
SERPAPI_RESERVE = 3           # stop using SerpApi when this many credits (or fewer) remain
EBAY_TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
EBAY_SEARCH_URL = "https://api.ebay.com/buy/browse/v1/item_summary/search"
EBAY_LIMIT = 40               # results per eBay search (your minimum was 20)
EBAY_MIN_FEEDBACK_PCT = 95.0  # sellers below this are kept for averages but never reported as deals
EBAY_MIN_FEEDBACK_COUNT = 10
# eBay condition IDs: New, New other, Cert/Excellent/Very good/Good refurbished, Seller refurb, Like new, Used, Very good
EBAY_CONDITION_IDS = "1000|1500|2000|2010|2020|2030|2500|2750|3000|4000"

# Domains for direct-to-consumer vendors (only used when isDirectToConsumer = 1).
# You can override/add by putting a "Domain" column in the Primary Vendor List tab.
KNOWN_DOMAINS = {
    "philipshue": "www.philips-hue.com",
    "sonos": "www.sonos.com",
    "thirdreality": "thirdreality.com",
}

# ---- Workbook layout --------------------------------------------------------
SHEET_MASTER = "Master Sheet"
SHEET_PRIMARY = "Primary Vendor List"
SHEET_SECONDARY = "Secondary Vendor List"
SHEET_RUN = "Run Data"
SHEET_DEALS = "Deals Data"

# Columns the script needs in the two output tabs. Existing columns are matched by name
# (case/punctuation-insensitive); missing ones are appended to the right of the header row.
RUN_COLS = ["RowDateTime", "runID", "WishlistItem", "Product", "Listings Searched",
            "Unique Websites Searched", "Target Price or better found",
            "Primary Vendor Price or Better", "Deals found", "Primary Vendor Deals",
            # --- added by this script (needed for the 30-day trend) ---
            "Avg Price (New)", "Avg Price (Resale)", "30d Trend (New)", "30d Trend (Resale)",
            "Lowest Price Found", "Source Notes"]
DEALS_COLS = ["runID", "WishlistItem", "Product", "URL", "isPrimaryVendor", "Target Price",
              "Price", "RightProductConfidence",
              # --- added by this script ---
              "Vendor", "Source", "Condition", "Listed Price", "Pack Qty", "% Below Target",
              "Deal Rule", "Secondary Vendor?", "Secondary Vendor Comments",
              "In Stock Verified", "Listing Title"]
MONEY_COLS = {"Target Price", "Price", "Listed Price", "Avg Price (New)", "Avg Price (Resale)",
              "30d Trend (New)", "30d Trend (Resale)", "Lowest Price Found"}
COL_WIDTHS = {"URL": 50, "Source Notes": 60, "Listing Title": 55, "Deal Rule": 42,
              "Secondary Vendor Comments": 55, "Vendor": 20, "RowDateTime": 17}


def log(msg: str) -> None:
    """Print immediately so GitHub Actions streams the log in real time."""
    print(msg, flush=True)


# =============================================================================
# 2. TEXT HELPERS  (normalisation, pack sizes, vendor names, product matching)
# =============================================================================

def norm_text(s: str) -> str:
    """Lower-case and normalise text so '4-INCH', '4"' and '4 inch' all compare equal."""
    s = (s or "").lower()
    s = s.replace("\u2019", "'").replace("\u2033", '"').replace("\u201d", '"').replace("\u201c", '"')
    s = re.sub(r'(\d+(?:\.\d+)?)\s*(?:"|-?\s*inch(?:es)?\b)', r"\1 inch ", s)   # 4" / 4-inch -> "4 inch"
    s = re.sub(r"\bgen\s*-?\s*(\d+)\b", r"gen\1", s)                            # "Gen 3" -> "gen3"
    s = re.sub(r"\bwi[\s-]?fi\b", "wifi", s)
    s = re.sub(r"(?<=\d)(?=(?:ghz|mhz|ft)\b)", " ", s)                          # "2.4ghz" -> "2.4 ghz"
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_NUMWORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
             "eight": 8, "nine": 9, "ten": 10, "twelve": 12}
_NUM = r"(\d{1,3}|" + "|".join(_NUMWORDS) + r")"
# Every phrase style we treat as "multi-pack": "Pack of 2", "2 Pack", "4-pack", "2pk", "Set of 3",
# "6 count", "twin pack" ...
_PACK_RES = [
    re.compile(r"\bpack\s+of\s+" + _NUM + r"\b", re.I),
    re.compile(r"\b(?:set|lot|bundle|case|box)\s+of\s+" + _NUM + r"\b", re.I),
    re.compile(r"\b" + _NUM + r"\s*[- ]?\s*(?:pack|pk|pcs|pc|pieces|piece|count|ct)\b", re.I),
    re.compile(r"\b(?:twin|double)\s*[- ]?\s*pack\b", re.I),
]


def _to_int(tok: str) -> int:
    return int(tok) if tok.isdigit() else _NUMWORDS.get(tok.lower(), 1)


def extract_pack_qty(title: str) -> int:
    """How many units are in this listing? (1 if no pack phrase is found.)"""
    best_pos, best_qty = None, 1
    for rx in _PACK_RES:
        m = rx.search(title or "")
        if m and (best_pos is None or m.start() < best_pos):
            best_pos = m.start()
            best_qty = 2 if not m.groups() else max(1, _to_int(m.group(1)))
    return best_qty


def strip_pack_phrases(title: str) -> str:
    """Remove pack phrases so '2 Pack' does not count against the name similarity."""
    for rx in _PACK_RES:
        title = rx.sub(" ", title or "")
    return title


def parse_bulk_sizes(text: str) -> set:
    """'2Pack; 4 Pack' -> {2, 4}. Sizes are what the Master Sheet's Bulk Keywords allow."""
    sizes = set()
    for part in re.split(r"[;,]", text or ""):
        part = part.strip()
        if not part:
            continue
        q = extract_pack_qty(part)
        if q > 1:
            sizes.add(q)
            continue
        m = re.search(r"\d+", part)
        if m and int(m.group()) > 1:
            sizes.add(int(m.group()))
    return sizes


def parse_price(v) -> Optional[float]:
    """'$1,299.00' -> 1299.0 ; 14.99 -> 14.99 ; junk -> None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"(\d[\d,]*\.?\d*)", str(v))
    return float(m.group(1).replace(",", "")) if m else None


# ---- Vendor-name matching ---------------------------------------------------
_VENDOR_STOP = {"the", "inc", "llc", "ltd", "co", "corp", "com", "net", "us", "usa", "official",
                "store", "shop", "direct", "online", "seller", "from"}
_VENDOR_ALIASES = {"bhphotovideo": {"bh", "bhphoto"}}


def vendor_key(name: str) -> str:
    """Reduce a vendor/source name to a comparable key:
       'Amazon.com - Seller' -> 'amazon', 'The Home Depot' -> 'homedepot', "Lowe's" -> 'lowes'."""
    s = (name or "").lower().replace("&", "").replace("'", "").replace("\u2019", "")
    s = re.sub(r"^from\s+", "", s)
    s = re.sub(r"\.(com|net|org|co)\b", "", s)
    toks = [t for t in re.split(r"[^a-z0-9]+", s) if t and t not in _VENDOR_STOP]
    return "".join(toks)


def keys_match(a: str, b: str) -> bool:
    """Loose vendor match: equal, alias, or one is a >=4-char prefix of the other."""
    if not a or not b:
        return False
    if a == b or a in _VENDOR_ALIASES.get(b, ()) or b in _VENDOR_ALIASES.get(a, ()):
        return True
    short, long_ = sorted((a, b), key=len)
    return len(short) >= 4 and long_.startswith(short)


# ---- Product matching -------------------------------------------------------
ACCESSORY_WORDS = {"case", "cover", "skin", "stand", "bracket", "mount", "mounting", "holder",
                   "protector", "sticker", "decal", "replacement"}
VARIANT_WORDS = {"pro", "plus", "max", "mini", "lite", "ultra", "se", "xl", "ambiance"}


def looks_like_sku(tok: str) -> bool:
    """True for model-number-like keywords ('RAYG1US1BLK', '609404') but not 'Gen3' / '4 inch'."""
    t = tok.strip()
    if len(t) < 5 or " " in t or not re.search(r"\d", t):
        return False
    if re.fullmatch(r"(?i)gen\d+|v\d+|\d+(?:\.\d+)?(?:ghz|inch|ft|w|v|mm|cm|k)", t):
        return False
    return re.fullmatch(r"[A-Za-z0-9\-_/.]+", t) is not None


def _name_present(tok: str, head_set: set, head_compact: str) -> bool:
    """Is a name token in the listing head? Also catches 'Third Reality' vs 'thirdreality'."""
    return tok in head_set or (len(tok) >= 6 and tok in head_compact)


def _in_order(name_tokens: list, head_tokens: list, head_set: set, head_compact: str) -> bool:
    """Do the product-name tokens appear in the listing head in the same order?"""
    pos = 0
    for tok in name_tokens:
        if tok in head_tokens[pos:]:
            pos = head_tokens.index(tok, pos) + 1
        elif _name_present(tok, head_set, head_compact):
            continue                      # found only as part of a longer/joined word
        else:
            return False
    return True


def classify_confidence(item: "Item", title: str, anchor_head: Optional[str] = None) -> tuple:
    """
    Return (High|Medium|Low, reason) for "is this listing the exact product we want?".

    Your rules, as implemented:
      * SKU/model number found in title  -> High if every spec/keyword also matches, else Medium.
      * Name matches exactly (all name words, same order, in the first ~70+ chars) ->
            High if no specs are defined OR all specs match; Medium if some specs are missing.
      * Listing mirrors a SKU-verified listing (similarity >= 0.8) and specs match -> High.
      * Name is a strong-but-not-exact match (>=75% of words or >=0.85 similarity) -> Medium.
      * Otherwise Low.
    Pack phrases ('2 Pack') are removed before comparing.
    Hard Lows: accessories ('case', 'mount'), a different 'GenN', a different colour/model
    code variant of the SKU, or a missing name word plus a variant word (e.g. 'Pro', 'Ambiance').
    """
    t_norm = norm_text(strip_pack_phrases(title))
    t_set = set(t_norm.split())
    t_compact = t_norm.replace(" ", "")
    name_norm = norm_text(item.product)
    name_tokens = name_norm.split()
    head_len = max(HEAD_MIN_CHARS, len(name_norm) + HEAD_EXTRA_CHARS)
    head = t_norm[:head_len]
    head_tokens = head.split()
    head_set = set(head_tokens)
    head_compact = head.replace(" ", "")
    spec_tokens = {tok for s in item.spec_phrases for tok in norm_text(s).split()}

    # ---- hard disqualifiers -------------------------------------------------
    acc = (head_set & ACCESSORY_WORDS) - set(name_tokens) - spec_tokens
    if acc:
        return "Low", f"accessory word '{sorted(acc)[0]}'"
    name_gens = set(re.findall(r"gen\d+", name_norm))
    title_gens = set(re.findall(r"gen\d+", t_norm))
    if name_gens and title_gens and not (name_gens & title_gens):
        return "Low", "different generation"
    for sku in item.skus:                                   # e.g. ...BLK vs ...WHT colour variants
        if len(sku) >= 8:
            for tok in re.findall(r"[A-Za-z0-9]{6,}", title):
                if tok.lower() != sku.lower() and len(tok) == len(sku):
                    common = len(os.path.commonprefix([tok.lower(), sku.lower()]))
                    if common >= len(sku) - 3:
                        return "Low", "different variant of SKU"

    # ---- similarity numbers -------------------------------------------------
    present = [_name_present(tok, head_set, head_compact) for tok in name_tokens]
    coverage = (sum(present) / len(name_tokens)) if name_tokens else 0.0
    exact = coverage == 1.0 and _in_order(name_tokens, head_tokens, head_set, head_compact)
    ratio = difflib.SequenceMatcher(None, name_norm, head[:len(name_norm) + 10]).ratio()

    if coverage < 1.0 and ((head_set & VARIANT_WORDS) - set(name_tokens)):
        return "Low", "variant word present and name word missing"

    # ---- spec check (whole title, not just head) ----------------------------
    specs_ok = True
    for spec in item.spec_phrases:
        toks = norm_text(spec).split()
        if toks and not all(t in t_set or (len(t) >= 5 and t in t_compact) for t in toks):
            specs_ok = False
            break

    sku_hit = any(sku.lower().replace("-", "") in t_compact for sku in item.skus)
    if sku_hit:
        return ("High", "SKU + specs match") if specs_ok else ("Medium", "SKU match, some specs not in title")
    if exact:
        if specs_ok:
            return "High", "exact name, specs match" if item.spec_phrases else "exact name, no specs defined"
        return "Medium", "exact name, specs not in title"
    if anchor_head and specs_ok and coverage >= 0.75:
        if difflib.SequenceMatcher(None, anchor_head[:head_len], head).ratio() >= 0.8:
            return "High", "mirrors SKU-verified listing"
    if coverage >= 0.75 or ratio >= 0.85:
        return "Medium", "strong name match"
    return "Low", "weak name match"


def detect_condition(title: str, hint: str = "") -> str:
    """Return new | used | refurbished | parts from a title and/or the source's condition field."""
    h, t = (hint or "").lower(), (title or "").lower()
    if re.search(r"for parts|not working|broken|damaged|as[- ]is|defective", t) or "parts" in h:
        return "parts"
    if "refurb" in h or re.search(r"refurb|renewed|reconditioned|recertified", t):
        return "refurbished"
    if re.search(r"open[- ]box|pre-?owned|\bused\b|like new|new other|very good|acceptable|\bgood\b|excellent", h):
        return "used"
    if re.search(r"open[- ]box|pre-?owned|\bused\b|like[- ]new", t):
        return "used"
    return "new"


# =============================================================================
# 3. DATA MODELS + WORKBOOK READERS
# =============================================================================

def truthy(v) -> bool:
    return str(v).strip().lower() in {"yes", "y", "true", "1", "1.0"}


def split_multi(v) -> list:
    """Cells can hold several values separated by ';'."""
    return [p.strip() for p in str(v or "").split(";") if p.strip()]


def norm_header(s) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s or "").lower())


def header_map(ws) -> dict:
    """{normalised header -> 1-based column index} for row 1."""
    return {norm_header(c.value): c.column for c in ws[1] if c.value not in (None, "")}


def find_col(hm: dict, *names: str) -> Optional[int]:
    """Exact normalised match first, then 'starts with' (so 'Priority (1 low - 5 high)' ~ 'priority')."""
    for n in names:
        if norm_header(n) in hm:
            return hm[norm_header(n)]
    for n in names:
        for k, v in hm.items():
            if k.startswith(norm_header(n)):
                return v
    return None


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
    skus: list = field(default_factory=list)            # model-number-like keywords
    spec_phrases: list = field(default_factory=list)    # specs + non-SKU keywords (must appear in title)


@dataclass
class Vendor:
    name: str
    key: str
    is_dtc: bool = False
    domain: str = ""


@dataclass
class Listing:
    """One price found by any pipeline."""
    title: str
    url: str
    price: float                      # price as listed (may be for a multi-pack)
    vendor: str
    source: str                       # 'direct' | 'serpapi' | 'ebay'
    condition: str = "new"
    in_stock: Optional[bool] = None   # None = unknown
    seller_comment: str = ""
    seller_ok: bool = True            # False = eBay seller below feedback thresholds
    # ---- filled in by score_listings() ----
    pack_qty: int = 1
    unit_price: float = 0.0           # price per single item
    vkey: str = ""
    is_primary: bool = False
    is_resale: bool = False
    confidence: str = "Low"
    conf_reason: str = ""
    eligible: bool = False            # counts toward averages
    reportable: bool = False          # may be reported as a deal / target hit
    deal_rule: str = ""


def load_items(ws) -> list:
    """Read Master Sheet rows into Item objects (skips blank rows)."""
    hm = header_map(ws)
    c = {k: find_col(hm, *v) for k, v in {
        "wid": ("WishlistItem",), "product": ("Product",), "specs": ("Product specifications",),
        "kw": ("Search Keywords",), "loc": ("Location",), "pri": ("Priority",),
        "qty": ("Quantity Needed",), "used": ("Open to used",), "target": ("Target Price",),
        "bulk": ("Is Bulk Option",), "bkw": ("Bulk Keywords",)}.items()}

    def get(r, k):
        return ws.cell(r, c[k]).value if c.get(k) else None

    items = []
    for r in range(2, ws.max_row + 1):
        product = get(r, "product")
        if not product or not str(product).strip():
            continue
        kws = split_multi(get(r, "kw"))
        specs = split_multi(get(r, "specs"))
        skus = [k for k in kws if looks_like_sku(k)]
        spec_kw = [k for k in kws if k not in skus]
        try:
            pri = int(float(get(r, "pri")))
        except (TypeError, ValueError):
            pri = 3
        items.append(Item(
            wid=get(r, "wid"), product=str(product).strip(), specs=specs, keywords=kws,
            location=str(get(r, "loc") or "").strip(), priority=pri, qty=get(r, "qty"),
            open_used=truthy(get(r, "used")), target=parse_price(get(r, "target")),
            bulk=truthy(get(r, "bulk")), bulk_sizes=parse_bulk_sizes(get(r, "bkw")),
            skus=skus, spec_phrases=specs + spec_kw))
    return items


def load_primary_vendors(ws) -> list:
    hm = header_map(ws)
    cv, cd, cdom = find_col(hm, "Vendor"), find_col(hm, "isDirectToConsumer"), find_col(hm, "Domain")
    out = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, cv).value if cv else None
        if not name:
            continue
        key = vendor_key(str(name))
        domain = str(ws.cell(r, cdom).value or "").strip() if cdom else ""
        out.append(Vendor(str(name).strip(), key, truthy(ws.cell(r, cd).value) if cd else False,
                          domain or KNOWN_DOMAINS.get(key, "")))
    return out


def load_secondary_keys(ws) -> list:
    hm = header_map(ws)
    cv = find_col(hm, "Vendor")
    return [vendor_key(str(ws.cell(r, cv).value)) for r in range(2, ws.max_row + 1)
            if cv and ws.cell(r, cv).value]


def _to_dt(v) -> Optional[datetime]:
    if isinstance(v, datetime):
        return v
    try:
        return datetime.fromisoformat(str(v))
    except (TypeError, ValueError):
        return None


def load_history(ws) -> list:
    """Past Run Data rows -> [{'wid','dt','avg_new','avg_res'}] (used for the 30-day trend + cadence)."""
    hm = header_map(ws)
    c_dt, c_w = find_col(hm, "RowDateTime"), find_col(hm, "WishlistItem")
    c_n, c_r = find_col(hm, "Avg Price (New)"), find_col(hm, "Avg Price (Resale)")
    out = []
    for r in range(2, ws.max_row + 1):
        dt = _to_dt(ws.cell(r, c_dt).value) if c_dt else None
        if dt is None:
            continue
        out.append({"wid": str(ws.cell(r, c_w).value).strip() if c_w else "", "dt": dt,
                    "avg_new": parse_price(ws.cell(r, c_n).value) if c_n else None,
                    "avg_res": parse_price(ws.cell(r, c_r).value) if c_r else None})
    return out


def trend_for(history: list, wid, pool: str, now: datetime) -> Optional[float]:
    """Mean of this item's stored per-run averages over the last TREND_WINDOW_DAYS (needs >= 3 points)."""
    cutoff = now - timedelta(days=TREND_WINDOW_DAYS)
    key = "avg_new" if pool == "new" else "avg_res"
    vals = [h[key] for h in history
            if h["wid"] == str(wid).strip() and h["dt"] >= cutoff and h[key]]
    return round(statistics.mean(vals), 2) if len(vals) >= TREND_MIN_POINTS else None


# =============================================================================
# 4. DATA SOURCES
# =============================================================================

def build_query(item: Item, include_specs: bool = True) -> str:
    """Product name + SKU (+ specs), skipping words already in the name."""
    parts, have = [item.product], set(norm_text(item.product).split())
    extras = item.skus + (item.spec_phrases if include_specs else [])
    for e in extras:
        toks = norm_text(e).split()
        if toks and not all(t in have for t in toks):
            parts.append(e)
            have.update(toks)
    return " ".join(parts)[:150]


# ---- 4a. Pipeline A: Google Shopping via SerpApi ------------------------------
class SerpApiClient:
    """Thin SerpApi wrapper that degrades gracefully: once disabled (no key, invalid key,
    out of credits, rate limited) every further call returns None instead of raising."""

    def __init__(self, key: Optional[str]):
        self.key = key
        self.disabled = not key
        self.reason = "SERPAPI_KEY not set" if not key else ""
        self.credits_left: Optional[float] = None
        self.calls = 0
        self.session = requests.Session()

    def _disable(self, why: str) -> None:
        if not self.disabled:
            log(f"  [SerpApi] disabled for rest of run: {why}")
        self.disabled, self.reason = True, why

    def check_credits(self) -> None:
        """Free account call (does not use a search credit) - bail out early if credits are nearly gone."""
        if self.disabled:
            return
        try:
            r = self.session.get(SERPAPI_ACCOUNT_URL, params={"api_key": self.key}, timeout=HTTP_TIMEOUT)
            if r.status_code in (401, 403):
                self._disable("invalid API key")
            elif r.ok:
                d = r.json()
                left = d.get("total_searches_left", d.get("plan_searches_left"))
                if isinstance(left, (int, float)):
                    self.credits_left = left
                    log(f"  [SerpApi] credits left this month: {int(left)}")
                    if left <= SERPAPI_RESERVE:
                        self._disable(f"only {int(left)} credits left (reserve={SERPAPI_RESERVE})")
        except Exception as e:                                   # network hiccup: keep going
            log(f"  [SerpApi] account check skipped ({type(e).__name__})")

    def shopping(self, query: str) -> Optional[list]:
        """Return the raw shopping_results list ([] = no results, None = source unavailable)."""
        if self.disabled:
            return None
        try:
            r = self.session.get(SERPAPI_SEARCH_URL, timeout=HTTP_TIMEOUT * 2, params={
                "engine": "google_shopping", "q": query, "gl": "us", "hl": "en",
                "google_domain": "google.com", "api_key": self.key})
            self.calls += 1
            if r.status_code in (401, 403):
                self._disable("invalid API key")
                return None
            if r.status_code == 429:
                self._disable("rate limited / out of searches")
                return None
            data = r.json()
            err = str(data.get("error", "")).lower()
            if err:
                if "hasn't returned any results" in err or "no results" in err:
                    return []
                if any(w in err for w in ("run out", "out of searches", "limit", "exceeded", "plan", "invalid api key")):
                    self._disable(data.get("error", "quota error"))
                else:
                    log(f"  [SerpApi] error: {data.get('error')}")
                return None
            if self.credits_left is not None:
                self.credits_left -= 1
                if self.credits_left <= SERPAPI_RESERVE:
                    self._disable("credit reserve reached")
            return data.get("shopping_results") or []
        except Exception as e:
            log(f"  [SerpApi] request failed ({type(e).__name__}: {e})")
            return None


def parse_serpapi(results: list) -> list:
    """Google Shopping rows -> Listing objects. NOTE: SerpApi returns a Google product-page link
    (which lists merchants), not a direct merchant URL, so these can't be page-verified."""
    out = []
    for r in results or []:
        title = r.get("title") or ""
        price = r.get("extracted_price")
        if price is None:
            price = parse_price(r.get("price"))
        if not title or not price or price <= 0:
            continue
        vendor = re.sub(r"^from\s+", "", r.get("source") or "", flags=re.I).strip()
        comment = ""
        if r.get("rating"):
            comment = (f"Google Shopping product rating {r['rating']}/5"
                       + (f" ({r['reviews']:,} reviews)" if isinstance(r.get("reviews"), int) else "")
                       + "; product-level, not seller-level")
        out.append(Listing(title=title, url=r.get("product_link") or r.get("link") or "",
                           price=float(price), vendor=vendor, source="serpapi",
                           condition=detect_condition(title, r.get("second_hand_condition") or ""),
                           seller_comment=comment))
    return out


# ---- 4b. Pipeline B: eBay Browse API ------------------------------------------
class EbayClient:
    """eBay Browse API (application token). Needs a *Production* keyset - see setup notes."""

    def __init__(self, client_id: Optional[str], client_secret: Optional[str]):
        self.cid, self.secret = client_id, client_secret
        self.disabled = not (client_id and client_secret)
        self.reason = "EBAY_CLIENT_ID/EBAY_CLIENT_SECRET not set" if self.disabled else ""
        self.token: Optional[str] = None
        self.session = requests.Session()

    def _disable(self, why: str) -> None:
        if not self.disabled:
            log(f"  [eBay] disabled for rest of run: {why}")
        self.disabled, self.reason = True, why

    def _get_token(self) -> bool:
        try:
            basic = base64.b64encode(f"{self.cid}:{self.secret}".encode()).decode()
            r = self.session.post(EBAY_TOKEN_URL, timeout=HTTP_TIMEOUT, headers={
                "Authorization": f"Basic {basic}", "Content-Type": "application/x-www-form-urlencoded"},
                data={"grant_type": "client_credentials", "scope": "https://api.ebay.com/oauth/api_scope"})
            if r.status_code == 200:
                self.token = r.json().get("access_token")
                return bool(self.token)
            self._disable(f"token request failed HTTP {r.status_code} (production keyset active yet?)")
        except Exception as e:
            self._disable(f"token request error {type(e).__name__}")
        return False

    def search(self, query: str) -> Optional[list]:
        """Return raw itemSummaries ([] = none, None = source unavailable)."""
        if self.disabled:
            return None
        if not self.token and not self._get_token():
            return None
        try:
            r = self.session.get(EBAY_SEARCH_URL, timeout=HTTP_TIMEOUT, headers={
                "Authorization": f"Bearer {self.token}", "X-EBAY-C-MARKETPLACE-ID": "EBAY_US"},
                params={"q": query, "limit": EBAY_LIMIT,
                        # fixed-price only (auction "price" is just the current bid), sensible conditions
                        "filter": f"buyingOptions:{{FIXED_PRICE}},conditionIds:{{{EBAY_CONDITION_IDS}}},priceCurrency:USD"})
            if r.status_code == 401 and self._get_token():       # token expired mid-run: retry once
                return self.search(query)
            if r.status_code in (403, 429):
                self._disable(f"HTTP {r.status_code}")
                return None
            if r.status_code != 200:
                log(f"  [eBay] HTTP {r.status_code}")
                return None
            return r.json().get("itemSummaries") or []
        except Exception as e:
            log(f"  [eBay] request failed ({type(e).__name__}: {e})")
            return None


def parse_ebay(items: list) -> list:
    """eBay itemSummaries -> Listings. Item price only (shipping ignored, as requested)."""
    out = []
    for it in items or []:
        p = it.get("price") or {}
        price = parse_price(p.get("value"))
        if not price or p.get("currency", "USD") != "USD" or not it.get("title"):
            continue
        s = it.get("seller") or {}
        pct, cnt = parse_price(s.get("feedbackPercentage")), s.get("feedbackScore")
        if pct is not None and isinstance(cnt, int):
            comment = f"eBay seller '{s.get('username', '?')}': {pct:g}% positive, {cnt:,} total feedback ratings"
            ok = pct >= EBAY_MIN_FEEDBACK_PCT and cnt >= EBAY_MIN_FEEDBACK_COUNT
        else:
            comment, ok = "N/A (seller feedback not returned)", False
        out.append(Listing(title=it["title"], url=it.get("itemWebUrl", ""), price=price, vendor="eBay",
                           source="ebay", condition=detect_condition(it["title"], it.get("condition", "")),
                           in_stock=True, seller_comment=comment, seller_ok=ok))
    return out


# ---- 4c. Direct-to-consumer vendor sites ---------------------------------------
_LD_RE = re.compile(r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', re.S | re.I)


def _walk_ld(node, depth=0):
    """Yield every dict inside a JSON-LD document (handles lists and @graph)."""
    if depth > 6:
        return
    if isinstance(node, dict):
        yield node
        for v in node.values():
            if isinstance(v, (dict, list)):
                yield from _walk_ld(v, depth + 1)
    elif isinstance(node, list):
        for v in node:
            yield from _walk_ld(v, depth + 1)


def parse_jsonld_product(html: str) -> Optional[dict]:
    """Pull {name, price, in_stock} from schema.org Product JSON-LD (most retail pages embed it)."""
    for block in _LD_RE.findall(html or ""):
        try:
            data = json.loads(block.strip())
        except ValueError:
            continue
        for node in _walk_ld(data):
            t = node.get("@type")
            if "Product" not in (t if isinstance(t, list) else [t]):
                continue
            offers = node.get("offers")
            offers = offers if isinstance(offers, list) else [offers] if isinstance(offers, dict) else []
            best = None
            for o in offers:
                price = parse_price(o.get("price", o.get("lowPrice")))
                if not price or o.get("priceCurrency", "USD") != "USD":
                    continue
                avail = str(o.get("availability", ""))
                stock = True if "InStock" in avail else False if avail else None
                if best is None or (stock and not best["in_stock"]) or (stock == best["in_stock"] and price < best["price"]):
                    best = {"name": node.get("name", ""), "price": price, "in_stock": stock}
            if best:
                return best
    return None


def direct_vendor_listings(vendor: Vendor, query: str, session: requests.Session) -> tuple:
    """
    Look up a product on one direct-to-consumer site. Returns (listings, note).
      1) Shopify stores expose a free JSON search (/search/suggest.json).
      2) Otherwise: DuckDuckGo HTML 'site:' search, then read the product page's JSON-LD.
    Anything that fails or is blocked returns ([], 'N/A ...') - never raises.
    """
    if not vendor.domain:
        return [], f"Direct[{vendor.name}]: N/A (no domain known; add a 'Domain' column to the vendor tab)"
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}
    base = f"https://{vendor.domain}"
    # ---- 1) Shopify predictive search ---------------------------------------
    try:
        r = session.get(f"{base}/search/suggest.json", headers=headers, timeout=HTTP_TIMEOUT, params={
            "q": query, "resources[type]": "product", "resources[limit]": 8})
        if r.ok and "json" in r.headers.get("content-type", ""):
            prods = (r.json().get("resources", {}).get("results", {}) or {}).get("products", [])
            out = []
            for p in prods:
                price = parse_price(p.get("price"))
                if not price:
                    continue
                out.append(Listing(title=p.get("title", ""), url=urljoin(base, (p.get("url") or "").split("?")[0]),
                                   price=price, vendor=vendor.name, source="direct",
                                   in_stock=p.get("available") if isinstance(p.get("available"), bool) else None))
            if out:
                return out, f"Direct[{vendor.name}]: {len(out)} listings (Shopify search)"
    except Exception:
        pass
    # ---- 2) search-engine fallback + JSON-LD ---------------------------------
    try:
        r = session.post("https://html.duckduckgo.com/html/", headers=headers, timeout=HTTP_TIMEOUT,
                         data={"q": f"site:{vendor.domain} {query}"})
        if r.status_code != 200:
            return [], f"Direct[{vendor.name}]: N/A (search engine returned HTTP {r.status_code})"
        bare = vendor.domain.replace("www.", "")
        links, seen = [], set()
        for raw in re.findall(r"uddg=([^&\"']+)", r.text):
            u = unquote(raw)
            if urlparse(u).netloc.replace("www.", "").endswith(bare) and u not in seen:
                seen.add(u)
                links.append(u)
        out = []
        for u in links[:2]:
            time.sleep(1)
            page = session.get(u, headers=headers, timeout=HTTP_TIMEOUT)
            info = parse_jsonld_product(page.text) if page.ok else None
            if info:
                out.append(Listing(title=info["name"] or query, url=u, price=info["price"],
                                   vendor=vendor.name, source="direct", in_stock=info["in_stock"]))
        if out:
            return out, f"Direct[{vendor.name}]: {len(out)} listings (page JSON-LD)"
        return [], f"Direct[{vendor.name}]: N/A (no readable product page found)"
    except Exception as e:
        return [], f"Direct[{vendor.name}]: N/A ({type(e).__name__})"


def vendor_name_matches_product(vendor_name: str, item: Item) -> bool:
    """Loose test used to decide whether to check a direct-to-consumer site at all:
       'Philips Hue Direct' ~ 'Hue Color Slim Downlight'. Saves runtime by skipping unrelated DTC sites."""
    toks = [t for t in re.split(r"[^a-z0-9]+", vendor_name.lower()) if len(t) >= 3 and t not in _VENDOR_STOP]
    hay = norm_text(item.product + " " + " ".join(item.keywords))
    hay_set, hay_compact = set(hay.split()), hay.replace(" ", "")
    return any(t in hay_set or (len(t) >= 6 and t in hay_compact) for t in toks)


# =============================================================================
# 5. SCORING + DEAL LOGIC
# =============================================================================

@dataclass
class Context:
    """Everything shared across items in one run."""
    run_id: str
    now: datetime
    primary: list
    secondary_keys: list
    history: list
    serp: SerpApiClient
    ebay: EbayClient
    session: requests.Session
    args: argparse.Namespace


def dedupe(listings: list) -> list:
    """Drop repeats of the same (vendor, unit price, title start). Direct > eBay API > SerpApi on ties."""
    rank = {"direct": 0, "ebay": 1, "serpapi": 2}
    seen, out = set(), []
    for l in sorted(listings, key=lambda x: rank.get(x.source, 9)):
        k = (l.vkey, round(l.unit_price, 2), norm_text(l.title)[:40])
        if k not in seen:
            seen.add(k)
            out.append(l)
    return out


def score_listings(item: Item, listings: list, ctx: Context) -> list:
    """Fill in pack size, unit price, vendor class, confidence, and eligibility for every listing."""
    for l in listings:
        l.vkey = vendor_key(l.vendor)
        l.pack_qty = extract_pack_qty(l.title) if l.source != "direct" else 1
        l.unit_price = round(l.price / max(l.pack_qty, 1), 2)
        l.is_primary = any(keys_match(l.vkey, v.key) for v in ctx.primary)
        on_secondary_list = any(keys_match(l.vkey, k) for k in ctx.secondary_keys) or l.source == "ebay"
        l.is_resale = on_secondary_list or l.condition in ("used", "refurbished")

    # First pass: any listing carrying the exact SKU becomes the "anchor" other listings are compared to.
    anchor = None
    sku_hits = [l for l in listings if item.skus and
                any(s.lower().replace("-", "") in norm_text(l.title).replace(" ", "") for s in item.skus)]
    if sku_hits:
        best = sorted(sku_hits, key=lambda l: (l.is_resale, not l.is_primary))[0]
        anchor = norm_text(strip_pack_phrases(best.title))[:max(HEAD_MIN_CHARS, len(item.product) + HEAD_EXTRA_CHARS)]

    allowed_packs = {1} | (item.bulk_sizes if item.bulk else set())
    for l in listings:
        l.confidence, l.conf_reason = classify_confidence(item, l.title, anchor)
        ok = (l.condition != "parts" and l.pack_qty in allowed_packs
              and l.confidence in ("High", "Medium") and l.in_stock is not False)
        if item.target:
            ok = ok and l.unit_price >= item.target * MIN_PRICE_RATIO_OF_TARGET
        if l.is_resale and not item.open_used:
            ok = False                                   # used/resale only when the Master Sheet says Yes
        l.eligible = ok
        # Items NOT open to used/wider search: only report new listings from Primary Vendors.
        # (Other new retailers still feed the price baseline, which stabilises the averages.)
        l.reportable = ok and (item.open_used or (l.is_primary and not l.is_resale))
    return listings


def pool_stats(prices: list) -> Optional[dict]:
    """Average with outliers removed (prices beyond 0.4x-2.5x the median are likely wrong products)."""
    if not prices:
        return None
    med = statistics.median(prices)
    lo, hi = med * OUTLIER_LOW, med * OUTLIER_HIGH
    kept = [p for p in prices if lo <= p <= hi]
    return {"avg": round(statistics.mean(kept), 2), "n": len(kept), "lo": lo} if kept else None


def find_deals(item: Item, listings: list, stats: dict, trends: dict) -> list:
    """Apply rules A and B per pool and return deals sorted: primary vendors first, then by price."""
    deals = []
    for l in listings:
        if not (l.reportable and l.seller_ok):
            continue
        if REQUIRE_AT_OR_BELOW_TARGET and item.target and l.unit_price > item.target:
            continue
        pool = "res" if l.is_resale else "new"
        st, trend, reasons = stats.get(pool), trends.get(pool), []
        if st and st["n"] >= MARKET_MIN_SAMPLE and l.unit_price >= st["lo"]:   # 'too good to be true' guard
            if l.unit_price <= st["avg"] * (1 - DEAL_DISCOUNT):
                reasons.append(f">=15% below run avg ${st['avg']:.2f}")
            if trend and l.unit_price < trend * (1 - TREND_MIN_DISCOUNT):
                reasons.append(f"below 30d trend ${trend:.2f}")
        if USE_TARGET_RULE and item.target and l.unit_price <= item.target * (1 - DEAL_DISCOUNT):
            reasons.append(f">=15% below target ${item.target:.2f}")
        if reasons:
            l.deal_rule = "; ".join(reasons) + f" [{'resale' if l.is_resale else 'new'} pool]"
            deals.append(l)
    deals.sort(key=lambda l: (0 if (l.is_primary and not l.is_resale) else 1 if not l.is_resale else 2,
                              l.unit_price))
    return deals[:MAX_DEALS_PER_ITEM]


def stock_status(l: Listing) -> str:
    """'In Stock Verified' column. Only direct-site pages and eBay's own live listings can be
    confirmed; Google Shopping rows come from merchant feeds we can't independently check -> N/A."""
    if l.source == "ebay":
        return "Yes"
    if l.source == "direct":
        return "Yes" if l.in_stock else "N/A"
    return "N/A"


def process_item(item: Item, ctx: Context) -> tuple:
    """Run every pipeline for one item. Returns (run_row dict, deal_row dicts, deal Listings)."""
    a = ctx.args
    listings, notes = [], []

    # --- Direct-to-consumer sites whose name loosely matches the product -------
    if not a.no_direct:
        for v in ctx.primary:
            if v.is_dtc and vendor_name_matches_product(v.name, item):
                ls, note = direct_vendor_listings(v, build_query(item, include_specs=False), ctx.session)
                listings += ls
                notes.append(note)
                time.sleep(1)

    # --- Pipeline A: Google Shopping via SerpApi --------------------------------
    serp_ls = []
    if a.no_serpapi:
        notes.append("SerpApi: skipped (--no-serpapi)")
    elif ctx.serp.disabled:
        notes.append(f"SerpApi: N/A ({ctx.serp.reason})")
    else:
        q1 = build_query(item, QUERY_INCLUDE_SPECS)
        raw = ctx.serp.shopping(q1)
        if raw is None:
            notes.append(f"SerpApi: N/A ({ctx.serp.reason or 'request failed'})")
        else:
            serp_ls = parse_serpapi(raw)
            q2 = build_query(item, include_specs=False)
            if len(serp_ls) < 3 and q2 != q1 and not ctx.serp.disabled:   # specs may have over-narrowed
                raw2 = ctx.serp.shopping(q2)
                serp_ls += parse_serpapi(raw2 or [])
            notes.append(f"SerpApi: {len(serp_ls)} listings")

    # --- Pipeline B: eBay (only when open to used / wider search) -------------
    ebay_ls, ebay_ok = [], False
    if item.open_used:
        if a.no_ebay:
            notes.append("eBay: skipped (--no-ebay)")
        elif ctx.ebay.disabled:
            notes.append(f"eBay: N/A ({ctx.ebay.reason})")
        else:
            raw = ctx.ebay.search(build_query(item, include_specs=False))
            if raw is None:
                notes.append(f"eBay: N/A ({ctx.ebay.reason or 'request failed'})")
            else:
                ebay_ok, ebay_ls = True, parse_ebay(raw)
                notes.append(f"eBay: {len(ebay_ls)} listings")
        notes.append("woot!/Mercari/Poshmark: only if Google Shopping surfaces them (no free API)")
    if ebay_ok:   # the API copy has seller ratings, so drop eBay rows that Google also returned
        serp_ls = [l for l in serp_ls if vendor_key(l.vendor) != "ebay"]

    listings = listings + serp_ls + ebay_ls
    score_listings(item, listings, ctx)
    listings = dedupe(listings)

    # --- Averages (current run) and trends (Run Data history) ------------------
    stats = {"new": pool_stats([l.unit_price for l in listings if l.eligible and not l.is_resale]),
             "res": pool_stats([l.unit_price for l in listings if l.eligible and l.is_resale])}
    trends = {"new": trend_for(ctx.history, item.wid, "new", ctx.now),
              "res": trend_for(ctx.history, item.wid, "res", ctx.now)}
    deals = find_deals(item, listings, stats, trends)

    # --- Build output rows -----------------------------------------------------
    rep = [l for l in listings if l.reportable]
    tgt = item.target
    at_target = [l for l in rep if tgt and l.unit_price <= tgt]
    lowest = min((l.unit_price for l in rep), default=None)

    def store(pool):      # store an average only when the sample is big enough to trust as trend data
        s = stats[pool]
        return s["avg"] if s and s["n"] >= MARKET_MIN_SAMPLE else None

    run_row = {
        "RowDateTime": ctx.now, "runID": ctx.run_id, "WishlistItem": item.wid, "Product": item.product,
        "Listings Searched": len(listings),
        "Unique Websites Searched": len({l.vkey for l in listings if l.vkey}),
        "Target Price or better found": len(at_target),
        "Primary Vendor Price or Better": sum(1 for l in at_target if l.is_primary),
        "Deals found": len(deals), "Primary Vendor Deals": sum(1 for l in deals if l.is_primary),
        "Avg Price (New)": store("new"), "Avg Price (Resale)": store("res"),
        "30d Trend (New)": trends["new"], "30d Trend (Resale)": trends["res"],
        "Lowest Price Found": lowest, "Source Notes": " | ".join(notes)[:900]}

    deal_rows = []
    for l in deals:
        deal_rows.append({
            "runID": ctx.run_id, "WishlistItem": item.wid, "Product": item.product, "URL": l.url,
            "isPrimaryVendor": 1 if l.is_primary else 0, "Target Price": tgt, "Price": l.unit_price,
            "RightProductConfidence": l.confidence, "Vendor": l.vendor, "Source": l.source,
            "Condition": l.condition, "Listed Price": l.price, "Pack Qty": l.pack_qty,
            "% Below Target": round((tgt - l.unit_price) / tgt, 4) if tgt else None,
            "Deal Rule": l.deal_rule, "Secondary Vendor?": "Yes" if l.is_resale else "No",
            "Secondary Vendor Comments": (l.seller_comment or "N/A (no seller rating available)") if l.is_resale else "",
            "In Stock Verified": stock_status(l), "Listing Title": l.title[:200]})
    return run_row, deal_rows, deals


# =============================================================================
# 6. WORKBOOK WRITER  (touches ONLY Run Data and Deals Data)
# =============================================================================

def ensure_columns(ws, wanted: list) -> dict:
    """Make sure every header in `wanted` exists in row 1 (appending missing ones, copying the
    style of the last header cell). Returns {normalised header -> column index}."""
    hm = header_map(ws)
    last = max(hm.values()) if hm else 0
    for name in wanted:
        if norm_header(name) not in hm:
            last += 1
            cell = ws.cell(1, last, name)
            if last > 1:
                src = ws.cell(1, last - 1)
                cell.font, cell.fill = copy(src.font), copy(src.fill)
                cell.alignment, cell.border = copy(src.alignment), copy(src.border)
            hm[norm_header(name)] = last
        if name in COL_WIDTHS:
            letter = get_column_letter(hm[norm_header(name)])
            cur = ws.column_dimensions[letter].width
            if not cur or cur < COL_WIDTHS[name]:
                ws.column_dimensions[letter].width = COL_WIDTHS[name]
    return hm


def next_empty_row(ws) -> int:
    r = ws.max_row
    while r > 1 and all(c.value is None for c in ws[r]):
        r -= 1
    return r + 1


def append_rows(ws, wanted: list, rows: list) -> None:
    hm = ensure_columns(ws, wanted)
    for row in rows:
        r = next_empty_row(ws)
        for name, val in row.items():
            col = hm.get(norm_header(name))
            if col is None or val is None:
                continue
            cell = ws.cell(r, col, val)
            if name == "RowDateTime":
                cell.number_format = "yyyy-mm-dd hh:mm"
            elif name in MONEY_COLS:
                cell.number_format = "$#,##0.00"
            elif name == "% Below Target":
                cell.number_format = "0.0%"


def write_github_summary(outcomes: list, run_id: str) -> None:
    """If running in GitHub Actions, show a deals table on the run's summary page."""
    path = os.getenv("GITHUB_STEP_SUMMARY")
    if not path:
        return
    lines = [f"### Price tracker run `{run_id}`", "",
             "| Item | Deals | Best price | Notes |", "|---|---|---|---|"]
    for item, run_row, _, deals in outcomes:
        best = f"${min(d.unit_price for d in deals):,.2f}" if deals else "-"
        lines.append(f"| {item.product} | {len(deals)} | {best} | {run_row['Source Notes'][:140]} |")
    try:
        with open(path, "a", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")
    except OSError:
        pass


# =============================================================================
# 7. MAIN
# =============================================================================

def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Home wishlist price tracker")
    p.add_argument("--workbook", help="path to the .xlsx (default: auto-detect Home*Wishlist*.xlsx)")
    p.add_argument("--min-priority", type=int, default=1, help="only items with Priority >= N")
    p.add_argument("--locations", default="", help="';'-separated location filter, e.g. 'Master Bedroom; Overall'")
    p.add_argument("--items", default="", help="comma-separated WishlistItem ids, e.g. 1,3")
    p.add_argument("--max-items", type=int, default=0, help="cap how many items run (highest priority first)")
    p.add_argument("--force", action="store_true", help="ignore the per-priority cadence (CADENCE_DAYS)")
    p.add_argument("--no-serpapi", action="store_true")
    p.add_argument("--no-ebay", action="store_true")
    p.add_argument("--no-direct", action="store_true")
    p.add_argument("--dry-run", action="store_true", help="do everything but do NOT write the workbook")
    return p.parse_args(argv)


def find_workbook(arg: Optional[str]) -> Path:
    if arg:
        p = Path(arg)
        if p.exists():
            return p
        raise FileNotFoundError(f"Workbook not found: {arg}")
    matches = sorted(Path(".").glob(WORKBOOK_GLOB))
    if matches:
        return matches[0]
    raise FileNotFoundError(f"No workbook matching '{WORKBOOK_GLOB}' in {Path('.').resolve()}")


def select_items(items: list, args: argparse.Namespace, history: list, now: datetime) -> list:
    """Apply priority / location / id filters and the per-priority cadence."""
    ids = {s.strip() for s in args.items.split(",") if s.strip()}
    locs = [s.lower() for s in split_multi(args.locations)]
    chosen = []
    for it in items:
        if it.priority < args.min_priority:
            continue
        if ids and str(it.wid).strip() not in ids:
            continue
        if locs and not any(l in it.location.lower() for l in locs):
            continue
        if not args.force:
            gap = CADENCE_DAYS.get(it.priority, 0)
            last = max((h["dt"] for h in history if h["wid"] == str(it.wid).strip()), default=None)
            if gap and last and (now - last) < timedelta(days=gap) - timedelta(hours=12):
                log(f"  skip '{it.product}' (priority {it.priority}: last checked {last:%Y-%m-%d}, cadence {gap}d)")
                continue
        chosen.append(it)
    chosen.sort(key=lambda i: -i.priority)
    return chosen[:args.max_items] if args.max_items else chosen


def main(argv=None) -> int:
    args = parse_args(argv)
    try:
        path = find_workbook(args.workbook)
        wb = load_workbook(path)
        for name in (SHEET_MASTER, SHEET_PRIMARY, SHEET_SECONDARY, SHEET_RUN, SHEET_DEALS):
            if name not in wb.sheetnames:
                raise KeyError(f"Missing worksheet '{name}'")
    except Exception as e:
        log(f"FATAL: cannot open workbook: {e}")
        return 2

    now = datetime.now(timezone.utc).replace(tzinfo=None, microsecond=0)
    run_id = now.strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
    log(f"Run {run_id} | workbook: {path} | dry-run: {args.dry_run}")

    items = load_items(wb[SHEET_MASTER])
    history = load_history(wb[SHEET_RUN])
    todo = select_items(items, args, history, now)
    log(f"{len(items)} wishlist items, {len(todo)} selected for this run")
    if not todo:
        return 0

    ctx = Context(run_id=run_id, now=now, primary=load_primary_vendors(wb[SHEET_PRIMARY]),
                  secondary_keys=load_secondary_keys(wb[SHEET_SECONDARY]), history=history,
                  serp=SerpApiClient(os.getenv("SERPAPI_KEY")),
                  ebay=EbayClient(os.getenv("EBAY_CLIENT_ID"), os.getenv("EBAY_CLIENT_SECRET")),
                  session=requests.Session(), args=args)
    if not args.no_serpapi:
        ctx.serp.check_credits()

    outcomes = []
    for it in todo:
        log(f"\n> [{it.wid}] {it.product} (priority {it.priority}, target ${it.target}, "
            f"{'wide+used' if it.open_used else 'primary vendors only'})")
        try:
            run_row, deal_rows, deals = process_item(it, ctx)
        except Exception as e:     # one bad item must never sink the whole run
            log(f"  ERROR processing item: {type(e).__name__}: {e}")
            run_row = {"RowDateTime": now, "runID": run_id, "WishlistItem": it.wid, "Product": it.product,
                       "Listings Searched": 0, "Unique Websites Searched": 0,
                       "Target Price or better found": 0, "Primary Vendor Price or Better": 0,
                       "Deals found": 0, "Primary Vendor Deals": 0,
                       "Source Notes": f"ERROR: {type(e).__name__}: {e}"[:300]}
            deal_rows, deals = [], []
        log(f"  listings={run_row['Listings Searched']} sites={run_row['Unique Websites Searched']} "
            f"deals={run_row['Deals found']} | {run_row['Source Notes']}")
        for d in deals:
            log(f"  DEAL ${d.unit_price:,.2f} @ {d.vendor} [{d.confidence}] {d.deal_rule} -> {d.url}")
        outcomes.append((it, run_row, deal_rows, deals))
        time.sleep(1)

    if args.dry_run:
        log("\nDry run: workbook NOT modified.")
    else:
        try:
            append_rows(wb[SHEET_RUN], RUN_COLS, [o[1] for o in outcomes])
            append_rows(wb[SHEET_DEALS], DEALS_COLS, [r for o in outcomes for r in o[2]])
            tmp = path.with_name(path.name + ".tmp")
            wb.save(tmp)
            os.replace(tmp, path)       # atomic swap so a crash can't leave a half-written workbook
            log(f"\nSaved {len(outcomes)} Run Data rows and {sum(len(o[2]) for o in outcomes)} Deals Data rows.")
        except Exception as e:
            log(f"FATAL: could not save workbook: {type(e).__name__}: {e}")
            return 2
    write_github_summary(outcomes, run_id)
    return 0


if __name__ == "__main__":
    sys.exit(main())
