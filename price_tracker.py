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

VENDOR BASELINE (ground truth)
  The pages in the Master Sheet's "Product URLs" column are priced first (retries, Best Buy API,
  Shopify variants, JSON-LD / microdata / meta / embedded-JSON extractors, curl_cffi Chrome
  impersonation, and finally the same store's Google Shopping row). Their median single-unit price -
  in stock OR sold out - is the "Vendor Baseline (New)". New listings outside 60-140% of it (and resale
  listings above it) are excluded from averages and deals, so bundles / wrong models can't skew them.

DEAL DEFINITION
  Only HIGH-confidence, in-stock (or unknown) listings can be deals. A listing is a deal when its
  per-item price is
      (A) >= 15% below the vendor baseline (or, with no Product URL priced, the run average
          of 3+ listings)   OR
      (B) below the trailing-30-day average of previous runs,
  computed SEPARATELY for two price pools:
      - "New"    : new-condition listings (primary vendors + any unlisted new retailer)
      - "Resale" : secondary-list vendors (eBay, woot!, ...) and any used/refurbished listing;
                   with fewer than 3 resale listings, >= 35% below the new vendor baseline counts.
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
import gzip
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
from html import unescape as html_unescape
from urllib.parse import unquote, urljoin, urlparse

import requests
# Real-Chrome TLS fingerprint client (gets past Akamai bot checks at Best Buy / Dell). Imported under its own
# name so it doesn't replace the standard `requests` used for SerpApi / eBay / sitemaps.
from curl_cffi import requests as cffi_requests      # pip install curl_cffi
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
MIN_PRICE_RATIO_OF_TARGET = 0.35   # listings priced under 35% of target are almost surely accessories
OUTLIER_LOW, OUTLIER_HIGH = 0.4, 2.5   # prices outside 0.4x..2.5x the pool median are excluded from averages
MAX_DEALS_PER_ITEM = 5             # cap Deals Data rows per item per run

# ---- Vendor baseline (ground truth from your Master Sheet "Product URLs") -------
# When at least one Product URL is priced, its (single-unit) price becomes the "Vendor Baseline".
# The New average is then built only from listings within this band around the baseline, and
# deal rule A compares against the baseline instead of a run average that junk listings can skew.
ANCHOR_BAND_NEW = (0.60, 1.40)     # new listings outside 60%..140% of the vendor baseline are ignored
ANCHOR_BAND_RESALE = (0.30, 1.05)  # used/refurb listings above the new baseline are almost always bundles/junk
DEAL_MIN_CONFIDENCE = "High"       # only High-confidence listings can become deals / target hits
# Out-of-stock vendor pages still publish the vendor's list price; count it toward the baseline/average
# (never reported as a deal, since you can't buy it).
OOS_COUNTS_FOR_BASELINE = True

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
PAGE_TIMEOUT = (10, 30)       # (connect, read) seconds for product pages - big retail pages are slow
PAGE_RETRIES = 2              # extra attempts per Product URL (with backoff) before giving up
BESTBUY_API_URL = "https://api.bestbuy.com/v1/products(sku={sku})"   # optional: BESTBUY_API_KEY env/secret
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

# Direct-site discovery (used only for isDirectToConsumer vendors whose name matches the product)
SITEMAP_MAX_FILES = 6          # max sitemap files fetched per domain per run (product sitemaps are tried first)
DIRECT_MAX_PAGES = 2           # max product pages fetched per vendor per item
SHOPIFY_MAX_PRODUCTS = 3       # max matching Shopify products whose variants (1/2/4-pack...) are read

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
            "Lowest Price Found", "Source Notes", "Matching Listings",
            "Avg Sample (New)", "Avg Sample (Resale)",    # how many listings each average is based on
            "Vendor Baseline (New)", "Baseline Source", "Product URL Status"]
DEALS_COLS = ["runID", "WishlistItem", "Product", "URL", "isPrimaryVendor", "Target Price",
              "Price", "RightProductConfidence",
              # --- added by this script ---
              "Vendor", "Source", "Condition", "Listed Price", "Pack Qty", "% Below Target",
              "Deal Rule", "Secondary Vendor?", "Secondary Vendor Comments",
              "In Stock Verified", "Listing Title", "Baseline Price", "% Below Baseline"]
MONEY_COLS = {"Target Price", "Price", "Listed Price", "Avg Price (New)", "Avg Price (Resale)",
              "30d Trend (New)", "30d Trend (Resale)", "Lowest Price Found", "Vendor Baseline (New)",
              "Baseline Price"}
PCT_COLS = {"% Below Target", "% Below Baseline"}
COL_WIDTHS = {"URL": 50, "Source Notes": 60, "Product URL Status": 50, "Listing Title": 55, "Deal Rule": 42,
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
# Sibling models sold under the same brand. If a listing names a sibling that is NOT in your product
# name (e.g. 'Sonos Beam' when you want 'Sonos Ray'), it is a different product -> Low.
# Add a brand here when you add a product from a new product family.
MODEL_FAMILIES = {
    # (words that commonly appear in genuine titles - 'bridge', 'hub', 'one', 'recessed' - are left out)
    "sonos": {"ray", "beam", "arc", "era", "five", "move", "roam", "sub", "ace", "amp", "port",
              "playbar", "playbase", "symfonisk"},
    "hue": {"lightstrip", "bloom", "iris", "signe", "gradient", "centris", "festavia", "datura",
            "filament", "candle", "dimmer"},
    "thirdreality": {"nightlight", "button", "e2", "zp1", "zp2"},
}
# Phrases that mean "works with X", i.e. an accessory for the product rather than the product itself.
_FOR_PRODUCT_RE = re.compile(r"\b(?:compatible with|for use with|designed for|fits|works with sonos|"
                             r"replacement for|for)\s+(?:the\s+)?(?:philips\s+)?(?:sonos|hue|thirdreality|third reality)\b",
                             re.I)


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
    A model/version word from your product name ('Gen3', 'E2') missing from the title is always Low.
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
    name_set = set(name_tokens)
    for brand, models in MODEL_FAMILIES.items():            # 'Sonos Beam' is not 'Sonos Ray'
        if brand in name_set or brand in name_norm.replace(" ", ""):
            other = (head_set & models) - name_set - spec_tokens
            if other:
                return "Low", f"different {brand} model '{sorted(other)[0]}'"
    if _FOR_PRODUCT_RE.search(title or "") and not (set(t_norm.split()[:3]) & (name_set | {"philips", "sonos"})):
        return "Low", "accessory 'for/compatible with' listing"
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
    # Model/version words in YOUR product name (letters+digits, e.g. 'gen3', 'e2', 'zp1') must appear in the
    # listing, otherwise 'ThirdReality Smart Plug E2' would pass for 'Smart Plug Gen3'. Exact SKU overrides.
    sku_in_title = any(sku.lower().replace("-", "") in t_compact for sku in item.skus)
    missing_model = [t for t in name_tokens if re.search(r"[a-z]", t) and re.search(r"\d", t)
                     and t not in t_set and t not in t_compact]
    if missing_model and not sku_in_title:
        return "Low", f"model '{missing_model[0]}' not in title"

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
    urls: list = field(default_factory=list)            # optional "Product URLs" column: exact pages to price-check


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
    from_url: bool = False            # True = priced from one of YOUR Master Sheet Product URLs (ground truth)
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
    ref_price: Optional[float] = None  # the baseline/average this deal was measured against


def load_items(ws) -> list:
    """Read Master Sheet rows into Item objects (skips blank rows)."""
    hm = header_map(ws)
    c = {k: find_col(hm, *v) for k, v in {
        "wid": ("WishlistItem",), "product": ("Product",), "specs": ("Product specifications",),
        "kw": ("Search Keywords",), "loc": ("Location",), "pri": ("Priority",),
        "qty": ("Quantity Needed",), "used": ("Open to used",), "target": ("Target Price",),
        "bulk": ("Is Bulk Option",), "bkw": ("Bulk Keywords",),
        "urls": ("Product URLs", "Product URL", "Product Links", "Links")}.items()}

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
            skus=skus, spec_phrases=specs + spec_kw,
            urls=extract_urls(ws.cell(r, c["urls"])) if c.get("urls") else []))
    return items


_URL_RE = re.compile(r"https?://[^\s;,<>\"']+", re.I)


def extract_urls(cell) -> list:
    """Every http(s) URL in a cell - separated by ';', ',', spaces or line breaks - plus the cell's
    clickable hyperlink (Excel stores that separately from the visible text, e.g. text 'Sonos.com')."""
    found = _URL_RE.findall(str(cell.value or ""))
    link = getattr(cell, "hyperlink", None)
    target = getattr(link, "target", None) if link else None
    if target and target.lower().startswith("http"):
        found.append(target)
    return list(dict.fromkeys(u.rstrip(").") for u in found))


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
    c_sn, c_sr = find_col(hm, "Avg Sample (New)"), find_col(hm, "Avg Sample (Resale)")
    c_b = find_col(hm, "Vendor Baseline (New)")
    out = []
    for r in range(2, ws.max_row + 1):
        dt = _to_dt(ws.cell(r, c_dt).value) if c_dt else None
        if dt is None:
            continue
        out.append({"wid": str(ws.cell(r, c_w).value).strip() if c_w else "", "dt": dt,
                    "avg_new": parse_price(ws.cell(r, c_n).value) if c_n else None,
                    "avg_res": parse_price(ws.cell(r, c_r).value) if c_r else None,
                    "n_new": parse_price(ws.cell(r, c_sn).value) if c_sn else None,
                    "n_res": parse_price(ws.cell(r, c_sr).value) if c_sr else None,
                    "baseline": parse_price(ws.cell(r, c_b).value) if c_b else None})
    return out


def trend_for(history: list, wid, pool: str, now: datetime) -> Optional[float]:
    """Mean of this item's stored per-run averages over the last TREND_WINDOW_DAYS (needs >= 3 points).
    Runs whose average came from fewer than MARKET_MIN_SAMPLE listings are left out of the trend,
    unless that run's New average was anchored to a vendor baseline (your Product URLs)."""
    cutoff = now - timedelta(days=TREND_WINDOW_DAYS)
    key, nkey = ("avg_new", "n_new") if pool == "new" else ("avg_res", "n_res")
    vals = [h[key] for h in history
            if h["wid"] == str(wid).strip() and h["dt"] >= cutoff and h[key]
            and (h.get(nkey) is None or h[nkey] >= MARKET_MIN_SAMPLE
                 or (pool == "new" and h.get("baseline")))]      # vendor-anchored averages are trustworthy
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


def availability_flag(avail) -> Optional[bool]:
    """schema.org availability -> True (buyable) / False (sold out) / None (unknown or pre/back-order)."""
    a = str(avail or "").lower().replace(" ", "").replace("_", "")
    if not a:
        return None
    if any(w in a for w in ("instock", "limitedavailability", "onlineonly", "instoreonly")):
        return True
    if any(w in a for w in ("outofstock", "soldout", "discontinued", "unavailable")):
        return False
    return None


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
                spec = o.get("priceSpecification")
                spec = spec[0] if isinstance(spec, list) and spec else spec if isinstance(spec, dict) else {}
                price = parse_price(o.get("price", o.get("lowPrice", spec.get("price"))))
                if not price or (o.get("priceCurrency") or spec.get("priceCurrency") or "USD") != "USD":
                    continue
                stock = availability_flag(o.get("availability", ""))
                if best is None or (stock and not best["in_stock"]) or (stock == best["in_stock"] and price < best["price"]):
                    best = {"name": node.get("name", ""), "price": price, "in_stock": stock}
            if best:
                return best
    return None


BROWSER_HEADERS = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9",
                   "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                   "Accept-Encoding": "gzip, deflate", "Connection": "keep-alive",
                   "Upgrade-Insecure-Requests": "1", "Cache-Control": "no-cache",
                   "Sec-Fetch-Dest": "document", "Sec-Fetch-Mode": "navigate", "Sec-Fetch-Site": "none",
                   "Sec-Fetch-User": "?1",
                   "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
                   "sec-ch-ua-mobile": "?0", "sec-ch-ua-platform": '"Windows"'}


_SITEMAP_CACHE: dict = {}          # domain -> list of URLs (shared by all items in one run)
_BLOCKED_STATUSES = {202, 401, 403, 429, 503}
_NON_PRODUCT_PATHS = ("/search", "/blog", "/support", "/help", "/compare", "/collections", "/category",
                      "/categories", "/stories", "/press", "/news", "/account", "/cart", "/login")


def _get(session, url, **kw):
    """GET that never raises (returns None on network errors)."""
    try:
        return session.get(url, headers=BROWSER_HEADERS, timeout=HTTP_TIMEOUT, **kw)
    except Exception:
        return None


def _us_locale_ok(url: str) -> bool:
    """Skip other-country pages (/en-gb/, /de-de/ ...) so prices are in USD for the US store."""
    for seg in urlparse(url).path.lower().split("/"):
        if re.fullmatch(r"[a-z]{2}[-_][a-z]{2}", seg) and seg.replace("_", "-") != "en-us":
            return False
    return True


def rank_product_urls(urls: list, item: Item, vendor: Vendor, limit: int = DIRECT_MAX_PAGES) -> list:
    """Pick the URLs whose path best matches the product name (brand words already implied by the
    domain are ignored, e.g. 'sonos' on sonos.com). '/en-us/shop/ray' scores well for 'Sonos Ray Soundbar'."""
    brand = set(re.split(r"[^a-z0-9]+", (vendor.domain + " " + vendor.name).lower()))
    name_tokens = [t for t in norm_text(item.product).split() if t not in brand] or norm_text(item.product).split()
    skus = [k.lower().replace("-", "") for k in item.skus]
    scored = []
    for u in dict.fromkeys(urls):                                   # de-duplicate, keep order
        path = unquote(urlparse(u).path).lower()
        if not _us_locale_ok(u) or path in ("", "/") or any(x in path for x in _NON_PRODUCT_PATHS):
            continue
        toks = set(norm_text(path).split())
        comp = norm_text(path).replace(" ", "")
        if (toks & ACCESSORY_WORDS) - set(name_tokens):
            continue                                                # '/ray-wall-mount' is an accessory
        cov = sum(1 for t in name_tokens if t in toks or (len(t) >= 4 and t in comp)) / len(name_tokens)
        sku_hit = any(k and k in comp for k in skus)
        if cov >= 0.5 or sku_hit:
            scored.append((cov + (1 if sku_hit else 0), -len(path), u))
    scored.sort(reverse=True)
    return [u for _, _, u in scored[:limit]]


def sitemap_links(domain: str, session) -> list:
    """All page URLs from the site's sitemaps (robots.txt -> sitemap index -> product sitemaps first).
    Sitemaps are published for search engines, so they are rarely bot-blocked. Cached per run."""
    if domain in _SITEMAP_CACHE:
        return _SITEMAP_CACHE[domain]
    urls, fetched = [], 0
    r = _get(session, f"https://{domain}/robots.txt")
    queue = re.findall(r"(?im)^\s*sitemap:\s*(\S+)", r.text) if (r is not None and r.ok) else []
    queue = queue or [f"https://{domain}/sitemap.xml"]
    while queue and fetched < SITEMAP_MAX_FILES:
        sm = queue.pop(0)
        fetched += 1
        r = _get(session, sm)
        if r is None or not r.ok:
            continue
        body = r.content
        if sm.endswith(".gz") or body[:2] == b"\x1f\x8b":
            try:
                body = gzip.decompress(body)
            except OSError:
                continue
        text = body.decode("utf-8", "ignore")
        locs = [html_unescape(x) for x in re.findall(r"<loc>\s*(?:<!\[CDATA\[)?\s*([^<\]\s]+)", text)]
        if "<sitemapindex" in text:
            kids = [k for k in locs if _us_locale_ok(k)]
            kids.sort(key=lambda k: ("product" not in k.lower(), not re.search(r"en[-_]?us|/us/", k.lower())))
            queue = kids + queue
        else:
            urls.extend(locs)
    _SITEMAP_CACHE[domain] = urls
    return urls


def _extract_result_links(page_html: str, domain: str) -> list:
    """Pull result URLs for `domain` out of a search-results page (handles DuckDuckGo and Bing redirect links)."""
    bare, out = domain.replace("www.", ""), []
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
        if urlparse(u).netloc.lower().replace("www.", "").endswith(bare) and u not in out:
            out.append(u)
    return out


SEARCH_ENGINES = [   # tried in order; each is free and key-less but may rate-limit automated traffic
    ("DuckDuckGo", "post", "https://html.duckduckgo.com/html/", lambda q: {"data": {"q": q}}),
    ("DuckDuckGo Lite", "post", "https://lite.duckduckgo.com/lite/", lambda q: {"data": {"q": q}}),
    ("Bing", "get", "https://www.bing.com/search", lambda q: {"params": {"q": q, "setlang": "en-US", "cc": "US"}}),
]


def search_engine_links(domain: str, query: str, session) -> tuple:
    """Last-resort 'site:' search across several engines. Returns (links, note)."""
    statuses = []
    for name, method, url, kw in SEARCH_ENGINES:
        try:
            fn = session.post if method == "post" else session.get
            r = fn(url, headers=BROWSER_HEADERS, timeout=HTTP_TIMEOUT, **kw(f"site:{domain} {query}"))
        except Exception as e:
            statuses.append(f"{name} {type(e).__name__}")
            continue
        links = _extract_result_links(r.text, domain) if r.status_code == 200 else []
        if links:
            return links, f"search ({name})"
        statuses.append(f"{name} blocked (HTTP {r.status_code})" if r.status_code in _BLOCKED_STATUSES
                        else f"{name} no results")
        time.sleep(1)
    return [], "search engines: " + ", ".join(statuses)


def parse_meta_price(page_html: str) -> Optional[dict]:
    """Fallback when a page has no JSON-LD: Open Graph / product meta tags."""
    m = re.search(r"""<meta[^>]+(?:property|name)=["'](?:product:price:amount|og:price:amount)["'][^>]+content=["']([^"']+)""",
                  page_html or "", re.I)
    if not m:
        return None
    a = re.search(r"""<meta[^>]+(?:property|name)=["'](?:product:availability|og:availability)["'][^>]+content=["']([^"']+)""",
                  page_html, re.I)
    t = re.search(r"""<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)""", page_html, re.I)
    avail = (a.group(1) if a else "").lower().replace(" ", "").replace("_", "")
    stock = None if not avail else False if "out" in avail else True if "instock" in avail else None
    return {"name": html_unescape(t.group(1)) if t else "", "price": parse_price(m.group(1)), "in_stock": stock}


def shopify_variant_listings(base: str, handle: str, title: str, vendor_name: str, session) -> list:
    """Every purchasable variant of a Shopify product (e.g. '1 Pack', '2 Pack', '4 Pack') with its own
    price and stock flag. Shopify's /products/<handle>.js returns prices in CENTS."""
    r = _get(session, f"{base}/products/{handle}.js")
    if r is None or not r.ok:
        return []
    try:
        data = r.json()
        variants = data.get("variants") or []
        title = data.get("title") or title          # the store's own product name beats ours
    except ValueError:
        return []
    out = []
    for v in variants:
        cents = v.get("price")
        price = cents / 100 if isinstance(cents, (int, float)) else parse_price(cents)
        if not price:
            continue
        vt = str(v.get("title") or "").strip()
        full = title if vt.lower() in ("", "default title") else f"{title} - {vt}"
        url = f"{base}/products/{handle}" + (f"?variant={v['id']}" if v.get("id") and len(variants) > 1 else "")
        out.append(Listing(title=full, url=url, price=price, vendor=vendor_name, source="direct",
                           in_stock=v["available"] if isinstance(v.get("available"), bool) else None))
    return out


def shopify_listings(vendor: Vendor, item: Item, session) -> Optional[tuple]:
    """Shopify stores: site search -> keep only results that actually match the product -> read each
    match's variants (pack sizes). Returns None when the site is not a Shopify store."""
    base = f"https://{vendor.domain}"
    r = _get(session, f"{base}/search/suggest.json", params={
        "q": build_query(item, include_specs=False), "resources[type]": "product", "resources[limit]": 10})
    if r is None or not r.ok or "json" not in r.headers.get("content-type", ""):
        return None
    try:
        prods = (r.json().get("resources", {}).get("results", {}) or {}).get("products", []) or []
    except ValueError:
        return None
    matched = [p for p in prods if classify_confidence(item, p.get("title", ""))[0] in ("High", "Medium")]
    out = []
    for p in matched[:SHOPIFY_MAX_PRODUCTS]:
        handle = p.get("handle") or (re.search(r"/products/([^/?#]+)", p.get("url") or "") or [None, None])[1]
        variants = shopify_variant_listings(base, handle, p.get("title", ""), vendor.name, session) if handle else []
        if variants:
            out += variants
        else:                                          # fall back to the search result's own price
            price = parse_price(p.get("price"))
            if price:
                out.append(Listing(title=p.get("title", ""), url=urljoin(base, (p.get("url") or "").split("?")[0]),
                                   price=price, vendor=vendor.name, source="direct",
                                   in_stock=p.get("available") if isinstance(p.get("available"), bool) else None))
    note = (f"Direct[{vendor.name}]: Shopify search {len(prods)} results, {len(matched)} matching product(s), "
            f"{len(out)} priced option(s)")
    return out, note


def parse_microdata_price(page_html: str) -> Optional[dict]:
    """schema.org microdata (<span itemprop="price" content="219.00">) - used by Dell and many older stores."""
    m = re.search(r"""itemprop=["']price["'][^>]*?content=["']([^"']+)""", page_html or "", re.I) or \
        re.search(r"""content=["']([\d.,]+)["'][^>]*?itemprop=["']price["']""", page_html or "", re.I)
    if not m or not parse_price(m.group(1)):
        return None
    a = re.search(r"""itemprop=["']availability["'][^>]*?(?:href|content)=["']([^"']+)""", page_html, re.I)
    n = re.search(r"""itemprop=["']name["'][^>]*?content=["']([^"']+)""", page_html, re.I)
    return {"name": html_unescape(n.group(1)) if n else "", "price": parse_price(m.group(1)),
            "in_stock": availability_flag(a.group(1)) if a else None}


# Keys that hold "the price you pay" inside a page's embedded JSON state (Next.js __NEXT_DATA__, Redux
# state, Best Buy / Dell data layers ...). Earlier keys are preferred; 'regularPrice' is the list price.
_EMBEDDED_PRICE_KEYS = ("customerPrice", "currentPrice", "salePrice", "finalPrice", "dellPrice",
                        "sellingPrice", "offerPrice", "priceValue", "regularPrice", "listPrice")


def parse_embedded_price(page_html: str, target: Optional[float] = None) -> Optional[dict]:
    """Last-resort extractor for JavaScript-heavy pages with no JSON-LD/meta price: scan embedded
    JSON for well-known price keys, then visible '<... class="...price...">$219.00' markup.
    Values implausibly far from your Target Price (<25% or >4x) are skipped - they are usually
    accessories, financing ('$18/mo') or bundle prices from elsewhere on the page."""
    text = page_html or ""

    def plausible(p):
        return p and p > 0 and (not target or target * MIN_PRICE_RATIO_OF_TARGET <= p <= target * 4)

    for key in _EMBEDDED_PRICE_KEYS:
        for m in re.finditer(r'\\?"%s\\?"\s*:\s*\\?"?\$?\s*([\d,]+(?:\.\d+)?)' % key, text):
            p = parse_price(m.group(1))
            if plausible(p):
                return {"name": "", "price": p, "in_stock": _embedded_stock(text), "how": f"embedded '{key}'"}
    for m in re.finditer(r"""(?:class|data-testid|id)=["'][^"']*price[^"']*["'][^>]*>\s*(?:<[^>]+>\s*){0,4}\$\s*([\d,]+\.\d{2})""",
                         text, re.I):
        p = parse_price(m.group(1))
        if plausible(p):
            return {"name": "", "price": p, "in_stock": _embedded_stock(text), "how": "visible price markup"}
    return None


def _embedded_stock(text: str) -> Optional[bool]:
    if re.search(r'"(?:buttonState|availability|stockStatus)"\s*:\s*"(?:SOLD_OUT|OUT_OF_STOCK|OutOfStock|SoldOut)"', text, re.I) \
            or re.search(r">\s*(?:Sold Out|Out of Stock|Currently unavailable)\s*<", text, re.I):
        return False
    if re.search(r'"(?:buttonState|availability|stockStatus)"\s*:\s*"(?:ADD_TO_CART|IN_STOCK|InStock)"', text, re.I):
        return True
    return None


def fetch_page(url: str, session) -> tuple:
    """GET a product page as robustly as we can. Returns (response|None, reason).
       1) requests with full browser headers, retried with backoff on network errors / 429 / 5xx;
       2) if still blocked or timing out, retry with curl_cffi impersonating real Chrome
          (Best Buy and Dell sit behind Akamai, which drops plain Python TLS handshakes - that is
          what showed up as 'network error')."""
    last = "network error"
    for attempt in range(PAGE_RETRIES + 1):
        try:
            r = session.get(url, headers=BROWSER_HEADERS, timeout=PAGE_TIMEOUT, allow_redirects=True)
        except Exception as e:
            last = f"network error ({type(e).__name__})"
            r = None
        if r is not None:
            if r.ok and r.status_code != 202:
                return r, "ok"
            last = (f"site blocked automated access (HTTP {r.status_code})" if r.status_code in _BLOCKED_STATUSES
                    else f"HTTP {r.status_code}")
            if r.status_code == 404:
                return None, last + " - check the URL in the Master Sheet"
            if r.status_code in (401, 403):
                break                                   # retrying the same client won't help
        time.sleep(1.5 * (attempt + 1))
    if cffi_requests is not None:              # (tests may switch it off)
        try:
            r = cffi_requests.get(url, impersonate="chrome", timeout=PAGE_TIMEOUT[1],
                                  headers={"Accept-Language": "en-US,en;q=0.9"})
            if r.status_code == 200:
                return r, "ok (browser impersonation)"
            last += f"; browser impersonation HTTP {r.status_code}"
        except Exception as e:
            last += f"; browser impersonation {type(e).__name__}"
    return None, last


def _bestbuy_sku(url: str) -> Optional[str]:
    m = re.search(r"(?:skuId=|/sku/|/)(\d{7})(?:\.p\b|\b)", url)
    return m.group(1) if m else None


def bestbuy_api_listing(url: str, vendor_name: str, session) -> tuple:
    """Best Buy's official Products API (free key at developer.bestbuy.com -> BESTBUY_API_KEY).
    Never blocked, returns salePrice + onlineAvailability. Returns (listings|None, reason)."""
    key, sku = os.getenv("BESTBUY_API_KEY"), _bestbuy_sku(url)
    if not key or not sku:
        return None, "no BESTBUY_API_KEY" if not key else "no SKU in URL"
    try:
        r = session.get(BESTBUY_API_URL.format(sku=sku), timeout=HTTP_TIMEOUT, params={
            "apiKey": key, "format": "json", "show": "sku,name,salePrice,regularPrice,onlineAvailability,url"})
        prods = (r.json().get("products") or []) if r.ok else []
    except Exception as e:
        return None, f"Best Buy API {type(e).__name__}"
    if not prods:
        return None, f"Best Buy API HTTP {r.status_code}, SKU {sku} not found"
    p = prods[0]
    price = parse_price(p.get("salePrice") or p.get("regularPrice"))
    if not price:
        return None, "Best Buy API returned no price"
    return [Listing(title=p.get("name") or "", url=url, price=price, vendor=vendor_name, source="direct",
                    in_stock=bool(p.get("onlineAvailability")))], "ok (Best Buy API)"


def page_listing(url: str, vendor_name: str, item: Item, session) -> tuple:
    """Price one product page. Returns (list of Listings | None, reason). Order of attempts:
       site API (Best Buy) -> Shopify variants (.js, every pack size) -> page JSON-LD -> microdata
       -> Open Graph meta -> embedded JSON / visible price markup."""
    parts = urlparse(url)
    host = parts.netloc.lower()
    if "bestbuy.com" in host:
        ls, why = bestbuy_api_listing(url, vendor_name, session)
        if ls:
            return ls, why
    shop = re.search(r"/products/([^/?#]+)", parts.path)
    if shop:                       # Shopify: the .js endpoint lists every variant (1/2/4-pack) + stock
        vs = shopify_variant_listings(f"{parts.scheme}://{parts.netloc}", shop.group(1), item.product, vendor_name, session)
        if vs:
            m = re.search(r"[?&]variant=(\d+)", url)          # you linked one specific variant -> keep only it
            pick = [v for v in vs if m and f"variant={m.group(1)}" in v.url]
            return (pick or vs), "ok (Shopify variants)"
    r, why = fetch_page(url, session)
    if r is None:
        return None, why
    html = r.text
    info, how = None, ""
    for fn, label in ((parse_jsonld_product, "JSON-LD"), (parse_microdata_price, "microdata"),
                      (parse_meta_price, "meta tags")):
        info = fn(html)
        if info and info.get("price"):
            how = label
            break
    if not (info and info.get("price")):
        info = parse_embedded_price(html, item.target)
        how = info.get("how", "") if info else ""
    if not (info and info.get("price")):
        return None, "no price data on page (JavaScript-rendered)"
    title = info.get("name") or ""
    if not title:
        t = re.search(r"<title[^>]*>([^<]+)</title>", html, re.I)
        title = html_unescape(t.group(1)).strip() if t else ""
    return [Listing(title=title or item.product, url=url, price=info["price"], vendor=vendor_name,
                    source="direct", in_stock=info.get("in_stock"))], f"ok ({how})"


def direct_vendor_listings(vendor: Vendor, item: Item, session) -> tuple:
    """
    Price a product on a direct-to-consumer site. Returns (listings, note); never raises.
      1) Shopify JSON search + variants (exact pack-size prices).
      2) The site's sitemap -> best-matching product URL -> page JSON-LD.  (rarely blocked)
      3) 'site:' search on DuckDuckGo, DuckDuckGo Lite, then Bing -> page JSON-LD.  (often blocked)
    """
    tag = f"Direct[{vendor.name}]"
    if not vendor.domain:
        return [], f"{tag}: N/A (no domain known; add a 'Domain' column to the vendor tab)"
    try:
        res = shopify_listings(vendor, item, session)
        if res is not None:
            return res
        links, how = rank_product_urls(sitemap_links(vendor.domain, session), item, vendor), "sitemap"
        if not links:
            found, how = search_engine_links(vendor.domain, build_query(item, include_specs=False), session)
            links = rank_product_urls(found, item, vendor)
            if not links:
                return [], f"{tag}: N/A (no product page in sitemap; {how})"
        out, problems = [], []
        for u in links:
            ls, why = page_listing(u, vendor.name, item, session)
            out += ls or []
            if not ls:
                problems.append(why)
            time.sleep(0.5)
        if out:
            return out, f"{tag}: {len(out)} listing(s) via {how}"
        return [], f"{tag}: N/A (found {urlparse(links[0]).path} via {how}, but {problems[0]})"
    except Exception as e:
        return [], f"{tag}: N/A ({type(e).__name__}: {e})"[:200]


def product_url_listings(item: Item, ctx: "Context") -> tuple:
    """Price the exact pages listed in the Master Sheet's optional 'Product URLs' column.
    Returns (listings, status_dict {url: reason}, set_of_domains_covered, failed [(url, host, vendor_name)]).
    These listings are YOUR verified product pages, so they are trusted as the exact product (High)."""
    out, status, covered, failed = [], {}, set(), []
    for u in item.urls:
        host = urlparse(u).netloc.lower().replace("www.", "")
        covered.add(host)
        vname = url_vendor_name(host, ctx.primary)
        ls, why = page_listing(u, vname, item, ctx.session)
        for l in ls or []:
            l.from_url = True
        out += ls or []
        if ls:
            stock = {True: "in stock", False: "OUT OF STOCK", None: "stock unknown"}[ls[0].in_stock]
            status[u] = f"{host}: ${ls[0].price:,.2f} {stock} - {why}" + (f" (+{len(ls) - 1} variants)" if len(ls) > 1 else "")
        else:
            status[u] = f"{host}: FAILED - {why}"
            failed.append((u, host, vname))
        time.sleep(0.5)
    return out, status, covered, failed


def url_vendor_name(host: str, primary: list) -> str:
    """'bestbuy.com' -> 'Best Buy' (your Primary Vendor name) so isPrimaryVendor is set correctly."""
    for v in primary:
        if v.domain and host.endswith(v.domain.replace("www.", "")):
            return v.name
    hkey = vendor_key(host.split(":")[0])
    return next((v.name for v in primary if keys_match(hkey, v.key)), host)


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
    """Drop repeats of the same offer. Your Product URLs > direct > eBay API > SerpApi on ties."""
    rank = {"direct": 0, "ebay": 1, "serpapi": 2}
    seen, out = set(), []
    for l in sorted(listings, key=lambda x: (not x.from_url, rank.get(x.source, 9))):
        # Same store + same unit price = same offer, even when Google's title differs from the page title.
        # (eBay is many independent sellers under one name, so its title/URL stay in the key.)
        k = (l.vkey, round(l.unit_price, 2)) + ((norm_text(l.title)[:40], l.url) if l.vkey == "ebay" else ())
        if k not in seen:
            seen.add(k)
            out.append(l)
    return out


def score_listings(item: Item, listings: list, ctx: Context) -> list:
    """Fill in pack size, unit price, vendor class, confidence, and eligibility for every listing."""
    for l in listings:
        l.vkey = vendor_key(l.vendor)
        l.pack_qty = extract_pack_qty(l.title)
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
        if l.from_url and l.source == "direct":     # you picked this exact page in the Master Sheet
            l.confidence, l.conf_reason = "High", "your Product URL"
        else:
            l.confidence, l.conf_reason = classify_confidence(item, l.title, anchor)
        ok = (l.condition != "parts" and l.pack_qty in allowed_packs and l.confidence in ("High", "Medium"))
        # Sold-out pages: a vendor's own sold-out page still shows its real list price, so it may feed the
        # baseline/average; a sold-out third-party listing is dropped. Neither can ever be a deal.
        if l.in_stock is False and not (OOS_COUNTS_FOR_BASELINE and l.source == "direct" and not l.is_resale):
            ok = False
        if item.target and not l.from_url:
            ok = ok and l.unit_price >= item.target * MIN_PRICE_RATIO_OF_TARGET
        if l.is_resale and not item.open_used:
            ok = False                                   # used/resale only when the Master Sheet says Yes
        l.eligible = ok
        # Items NOT open to used/wider search: only report new listings from Primary Vendors.
        # (Other new retailers still feed the price baseline, which stabilises the averages.)
        l.reportable = (ok and l.in_stock is not False
                        and (item.open_used or (l.is_primary and not l.is_resale)))
    return listings


def vendor_baseline(item: Item, listings: list) -> tuple:
    """Ground-truth 'New' price from YOUR Product URLs: median single-unit price of those pages
    (in stock or not). Falls back to High-confidence direct-vendor pages the script found itself.
    Returns (single-unit price | None, source description, {pack_qty: per-unit price})."""
    for label, pool in (("Product URLs", [l for l in listings if l.from_url]),
                        ("direct vendor site", [l for l in listings if l.source == "direct" and not l.from_url
                                                and l.confidence == "High"])):
        pool = [l for l in pool if l.eligible and not l.is_resale]
        if not pool:
            continue
        singles = [l for l in pool if l.pack_qty == 1] or pool
        price = round(statistics.median(l.unit_price for l in singles), 2)
        names = sorted({l.vendor for l in singles})
        # Per-pack-size baselines: a 4-pack is compared with the vendor's own 4-pack price, so normal
        # bulk pricing is not reported as a "deal" on every run.
        packs = {q: round(statistics.median(l.unit_price for l in pool if l.pack_qty == q), 2)
                 for q in {l.pack_qty for l in pool}}
        return price, f"{label}: {', '.join(names)} (median of {len(singles)})", packs
    return None, "none (no Product URL priced) - using market average", {}


def apply_baseline(item: Item, listings: list, baseline: Optional[float]) -> None:
    """With a vendor baseline in hand:
       * listings outside the plausible band around it stop counting toward averages/deals
         (this is what kept $454 'Sonos Ray' bundles and Beam/Arc listings out of the New average);
       * an exact-SKU Medium listing priced inside the band is promoted to High (SKU + price both agree)."""
    if not baseline:
        # No vendor price: SKU + agreement with the market median still corroborates a Medium listing.
        for resale in (False, True):
            prices = [l.unit_price for l in listings if l.eligible and l.is_resale == resale]
            if len(prices) < MARKET_MIN_SAMPLE:
                continue
            med = statistics.median(prices)
            for l in listings:
                if (l.eligible and l.is_resale == resale and l.confidence == "Medium" and item.skus
                        and med * ANCHOR_BAND_NEW[0] <= l.unit_price <= med * ANCHOR_BAND_NEW[1]
                        and any(s.lower().replace("-", "") in norm_text(l.title).replace(" ", "") for s in item.skus)):
                    l.confidence, l.conf_reason = "High", "SKU match + price agrees with market median"
        return
    for l in listings:
        if l.from_url:
            continue
        lo, hi = ANCHOR_BAND_RESALE if l.is_resale else ANCHOR_BAND_NEW
        inside = baseline * lo <= l.unit_price <= baseline * hi
        if not inside:
            if l.eligible:
                l.conf_reason += f"; price outside {lo:.0%}-{hi:.0%} of vendor baseline"
            l.eligible = l.reportable = False
        elif l.confidence == "Medium" and item.skus and \
                any(s.lower().replace("-", "") in norm_text(l.title).replace(" ", "") for s in item.skus):
            l.confidence, l.conf_reason = "High", "SKU match + price agrees with vendor baseline"


def pool_stats(prices: list) -> Optional[dict]:
    """Average with outliers removed (prices beyond 0.4x-2.5x the median are likely wrong products).
    When a vendor baseline exists, apply_baseline() has already narrowed the pool around it."""
    if not prices:
        return None
    med = statistics.median(prices)
    lo, hi = med * OUTLIER_LOW, med * OUTLIER_HIGH
    kept = [p for p in prices if lo <= p <= hi]
    return {"avg": round(statistics.mean(kept), 2), "n": len(kept), "lo": lo} if kept else None


RESALE_VS_NEW_DISCOUNT = 0.35   # resale deal when the resale pool is too small: >=35% below the new baseline


def find_deals(item: Item, listings: list, stats: dict, trends: dict, baseline: Optional[float] = None,
               pack_baselines: Optional[dict] = None) -> list:
    """Apply the deal rules and return deals sorted: primary vendors first, then by price.
    Only listings at DEAL_MIN_CONFIDENCE (High) with an in-stock/unknown status can become deals.
      New pool    - rule A: >=15% below the VENDOR BASELINE (or the run average when no Product URL priced,
                    which then needs MARKET_MIN_SAMPLE listings); rule B: below the 30-day trend.
      Resale pool - rule A: >=15% below the resale average (needs MARKET_MIN_SAMPLE listings), or
                    >=35% below the new vendor baseline; rule B: below the resale 30-day trend.
    Each deal records the reference price it beat in l.ref_price."""
    deals = []
    for l in listings:
        if not (l.reportable and l.seller_ok and l.confidence == DEAL_MIN_CONFIDENCE):
            continue
        if REQUIRE_AT_OR_BELOW_TARGET and item.target and l.unit_price > item.target:
            continue
        pool = "res" if l.is_resale else "new"
        st, trend, reasons, ref = stats.get(pool), trends.get(pool), [], None
        market_ok = bool(st and st["n"] >= MARKET_MIN_SAMPLE)
        if pool == "new" and baseline:
            ref = (pack_baselines or {}).get(l.pack_qty, baseline)
            ref_label = "vendor baseline" + (f" ({l.pack_qty}-pack)" if l.pack_qty in (pack_baselines or {}) and l.pack_qty > 1 else "")
        elif market_ok:
            ref, ref_label = st["avg"], "run avg"
        elif pool == "res" and baseline:
            ref, ref_label = None, ""
            if l.unit_price <= baseline * (1 - RESALE_VS_NEW_DISCOUNT):
                reasons.append(f">={RESALE_VS_NEW_DISCOUNT:.0%} below new vendor baseline ${baseline:.2f}")
                l.ref_price = baseline
        else:
            ref_label = ""
        if ref and l.unit_price >= ref * OUTLIER_LOW:                     # 'too good to be true' guard
            if l.unit_price <= ref * (1 - DEAL_DISCOUNT):
                reasons.append(f">={DEAL_DISCOUNT:.0%} below {ref_label} ${ref:.2f}")
            l.ref_price = ref
        if trend and (market_ok or baseline) and l.unit_price < trend * (1 - TREND_MIN_DISCOUNT):
            reasons.append(f"below 30d trend ${trend:.2f}")
            l.ref_price = l.ref_price or trend
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


def google_fill_failed_urls(item: Item, failed: list, serp_ls: list) -> tuple:
    """A Product URL whose page could not be priced (blocked, JavaScript-only) is filled from the SAME
    merchant's Google Shopping row, if Google has a matching one. The Google row is moved (not copied)
    so it is never counted twice. Returns (filled listings, remaining serp listings, {url: note})."""
    filled, notes, used = [], {}, set()
    for url, host, vname in failed:
        keys = {vendor_key(vname), vendor_key(host)}
        cands = [l for l in serp_ls if id(l) not in used and any(keys_match(vendor_key(l.vendor), k) for k in keys)
                 and l.condition == "new" and classify_confidence(item, l.title)[0] in ("High", "Medium")]
        if not cands:
            continue
        best = min(cands, key=lambda l: (classify_confidence(item, l.title)[0] != "High", extract_pack_qty(l.title)))
        used.add(id(best))
        best.from_url, best.vendor = True, vname
        best.seller_comment = f"price from Google Shopping (direct page failed); Google link: {best.url}"
        best.url = url
        filled.append(best)
        notes[url] = f"{host}: ${best.price:,.2f} via Google Shopping fallback"
    return filled, [l for l in serp_ls if id(l) not in used], notes


def process_item(item: Item, ctx: Context) -> tuple:
    """Run every pipeline for one item. Returns (run_row dict, deal_row dicts, deal Listings)."""
    a = ctx.args
    listings, notes = [], []

    # --- Exact pages you listed in the Master Sheet ("Product URLs" column) - the ground truth ---
    covered, url_status, failed = set(), {}, []
    if item.urls and not a.no_direct:
        ls, url_status, covered, failed = product_url_listings(item, ctx)
        listings += ls

    # --- Direct-to-consumer sites whose name loosely matches the product -------
    if not a.no_direct:
        for v in ctx.primary:
            if v.is_dtc and vendor_name_matches_product(v.name, item):
                if v.domain and v.domain.replace("www.", "") in covered:
                    continue                      # you already gave us this vendor's exact page
                ls, note = direct_vendor_listings(v, item, ctx.session)
                listings += ls
                notes.append(note)

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

    # --- Google Shopping <-> Product URLs: fill failures, never double count ------------
    if failed and serp_ls:
        filled, serp_ls, fill_notes = google_fill_failed_urls(item, failed, serp_ls)
        listings += filled
        url_status.update(fill_notes)
    ok_keys = {vendor_key(l.vendor) for l in listings if l.from_url and l.source == "direct"}
    if ok_keys:   # the store's own page was priced directly; its Google row is a stale duplicate
        before = len(serp_ls)
        serp_ls = [l for l in serp_ls if not any(keys_match(vendor_key(l.vendor), k) for k in ok_keys)]
        if before != len(serp_ls):
            notes.append(f"SerpApi: {before - len(serp_ls)} row(s) dropped - same store already priced from your Product URL")

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

    # --- Vendor baseline (ground truth), averages, trends --------------------------
    baseline, baseline_src, pack_baselines = vendor_baseline(item, listings)
    apply_baseline(item, listings, baseline)
    stats = {"new": pool_stats([l.unit_price for l in listings if l.eligible and not l.is_resale]),
             "res": pool_stats([l.unit_price for l in listings if l.eligible and l.is_resale])}
    trends = {"new": trend_for(ctx.history, item.wid, "new", ctx.now),
              "res": trend_for(ctx.history, item.wid, "res", ctx.now)}
    deals = find_deals(item, listings, stats, trends, baseline, pack_baselines)

    # --- Build output rows -----------------------------------------------------
    rep = [l for l in listings if l.reportable and l.confidence == DEAL_MIN_CONFIDENCE]
    tgt = item.target
    at_target = [l for l in rep if tgt and l.unit_price <= tgt]
    lowest = min((l.unit_price for l in rep), default=None)

    def store(pool):      # always record the average; its sample size is stored beside it
        s = stats[pool]
        return s["avg"] if s else None

    def sample(pool):
        s = stats[pool]
        return s["n"] if s else None

    for pool, label in (("new", "New"), ("res", "Resale")):
        s = stats[pool]
        if s and s["n"] < MARKET_MIN_SAMPLE and not (pool == "new" and baseline):
            notes.append(f"Avg ({label}) from only {s['n']} listing(s): recorded, but deal rules and the "
                         f"30-day trend need {MARKET_MIN_SAMPLE}+")
    if baseline:
        oos = [l.vendor for l in listings if l.from_url and l.in_stock is False]
        if oos:
            notes.append(f"Vendor baseline includes sold-out page(s) ({', '.join(sorted(set(oos)))}) - list price only")
    if item.urls and a.no_direct:
        url_status = {u: "skipped (--no-direct)" for u in item.urls}

    run_row = {
        "RowDateTime": ctx.now, "runID": ctx.run_id, "WishlistItem": item.wid, "Product": item.product,
        "Listings Searched": len(listings), "Matching Listings": sum(1 for l in listings if l.eligible),
        "Unique Websites Searched": len({l.vkey for l in listings if l.vkey}),
        "Target Price or better found": len(at_target),
        "Primary Vendor Price or Better": sum(1 for l in at_target if l.is_primary),
        "Deals found": len(deals), "Primary Vendor Deals": sum(1 for l in deals if l.is_primary),
        "Avg Price (New)": store("new"), "Avg Price (Resale)": store("res"),
        "Avg Sample (New)": sample("new"), "Avg Sample (Resale)": sample("res"),
        "30d Trend (New)": trends["new"], "30d Trend (Resale)": trends["res"],
        "Vendor Baseline (New)": baseline, "Baseline Source": baseline_src,
        "Product URL Status": " | ".join(url_status.values())[:900] if url_status else
                              ("none listed in Master Sheet" if not item.urls else None),
        "Lowest Price Found": lowest, "Source Notes": " | ".join(notes)[:900]}

    deal_rows = []
    for l in deals:
        ref = l.ref_price
        deal_rows.append({
            "runID": ctx.run_id, "WishlistItem": item.wid, "Product": item.product, "URL": l.url,
            "isPrimaryVendor": 1 if l.is_primary else 0, "Target Price": tgt, "Price": l.unit_price,
            "RightProductConfidence": l.confidence, "Vendor": l.vendor, "Source": l.source,
            "Condition": l.condition, "Listed Price": l.price, "Pack Qty": l.pack_qty,
            # per-UNIT price vs per-item target; negative = above your target
            "% Below Target": round((tgt - l.unit_price) / tgt, 4) if tgt else None,
            "Baseline Price": ref,
            "% Below Baseline": round((ref - l.unit_price) / ref, 4) if ref else None,
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
            elif name in PCT_COLS:
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
    p.add_argument("--force", "--allproducts", dest="force", action="store_true",
                   help="check every selected product now, ignoring the per-priority cadence (CADENCE_DAYS)")
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


def load_dotenv(path: Path) -> None:
    """Minimal .env reader for local runs (Cursor/VS Code terminal). Lines like SERPAPI_KEY=abc123.
    Existing environment variables win, so GitHub Actions secrets are never overridden.
    NEVER commit .env - the repo is public. It is listed in .gitignore."""
    try:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip().removeprefix("export ").strip(), v.strip().strip('"').strip("'")
            if k and k not in os.environ:
                os.environ[k] = v
    except OSError:
        pass


def main(argv=None) -> int:
    for env_file in (Path(".env"), Path(__file__).resolve().with_name(".env")):
        load_dotenv(env_file)
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
        if ctx.serp.disabled:
            log("WARNING: SERPAPI_KEY is not set - Google Shopping (the main source for Amazon, Best Buy, "
                "Home Depot and every unlisted retailer) is OFF. Add it to .env locally or to repo secrets on GitHub.")
        ctx.serp.check_credits()
    if not args.no_ebay and ctx.ebay.disabled and any(i.open_used for i in todo):
        log("NOTE: eBay keys not set - used/refurbished listings will only come from Google Shopping.")

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
