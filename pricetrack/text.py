"""Text normalisation, pack sizes, prices, conditions and vendor-name keys."""
from __future__ import annotations

import re
from typing import Optional


def norm_text(s: str) -> str:
    """Lower-case and normalise text so '4-INCH', '4"' and '4 inch' all compare equal."""
    s = (s or "").lower()
    s = s.replace("’", "'").replace("″", '"').replace("”", '"').replace("“", '"')
    s = re.sub(r'(\d+(?:\.\d+)?)\s*(?:"|-?\s*inch(?:es)?\b|-?\s*in\.(?=\s|$))', r"\1 inch ", s)   # 4" / 4-inch -> "4 inch"
    s = re.sub(r"\b(\d)(?:st|nd|rd|th)[\s-]*gen(?:eration)?\b", r"gen\1", s)         # "2nd Gen" -> "gen2"
    s = re.sub(r"\bgeneration\s*(\d+)\b", r"gen\1", s)
    s = re.sub(r"\bgen\s*-?\s*(\d+)\b", r"gen\1", s)                                    # "Gen 3" -> "gen3"
    s = re.sub(r"\bwi[\s-]?fi\b", "wifi", s)
    s = re.sub(r"(?<=\d)(?=(?:ghz|mhz|ft)\b)", " ", s)                                  # "2.4ghz" -> "2.4 ghz"
    s = re.sub(r"[^a-z0-9.]+", " ", s)
    s = re.sub(r"(?<!\d)\.|\.(?!\d)", " ", s)                                           # keep decimals only
    return re.sub(r"\s+", " ", s).strip()


_NUMWORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
             "eight": 8, "nine": 9, "ten": 10, "twelve": 12}
_NUM = r"(\d{1,3}|" + "|".join(_NUMWORDS) + r")"
PACK_RES = [
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
    for rx in PACK_RES:
        m = rx.search(title or "")
        if m and (best_pos is None or m.start() < best_pos):
            best_pos = m.start()
            best_qty = 2 if not m.groups() else max(1, _to_int(m.group(1)))
    return best_qty


def strip_pack_phrases(title: str) -> str:
    for rx in PACK_RES:
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
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"(\d[\d,]*\.?\d*)", str(v))
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


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


# ---- Vendor-name keys -----------------------------------------------------------------------------
_VENDOR_STOP = {"the", "inc", "llc", "ltd", "co", "corp", "com", "net", "us", "usa", "official",
                "store", "shop", "direct", "online", "seller", "from", "www"}
_VENDOR_ALIASES: dict = {}


def vendor_key(name: str) -> str:
    """'Amazon.com - Seller' -> 'amazon', 'The Home Depot' -> 'homedepot'."""
    s = (name or "").lower().replace("&", "").replace("'", "").replace("’", "")
    s = re.sub(r"^from\s+", "", s)
    s = re.sub(r"^https?://", "", s)
    s = re.sub(r"\.(com|net|org|co)\b.*$", "", s)
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
