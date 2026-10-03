"""
Identifier-first product matching.

Precedence (strongest first):
    GTIN/UPC/EAN  >  manufacturer SKU / MPN / model  >  trusted page metadata  >  title/spec match

Every listing is first checked for HARD CONFLICTS - any one of them makes it Low regardless of how good
the title looks:
    * a different GTIN (single-unit listings) or a near-variant model code (RAYG1US1BLK vs RAYG1EU1BLK)
    * region (EU/UK/International version, 220-240V), generation/version, colour, voltage, dimensions
    * accessories ('mount for ...', 'compatible with ...'), bundles ('+ Sub Mini', 'combo'),
      configuration words (Pro/Max/Mini/...) and sibling models of the same brand ('Sonos Beam' vs 'Ray')
Then positive evidence decides High / Medium / Low.

Sibling models are detected GENERICALLY (brand + model-token rule). The old hand-maintained
MODEL_FAMILIES table is kept only as SUPPLEMENTAL_FAMILY_RULES (extra protection, not the main defence).
"""
from __future__ import annotations

import difflib
import hashlib
import os
import re
from dataclasses import dataclass
from typing import Iterable, Optional

from .models import Item
from .text import norm_text, strip_pack_phrases

HEAD_MIN_CHARS = 70           # only the first N chars of a listing title are compared to the product name:
HEAD_EXTRA_CHARS = 40         #   max(70, len(name) + 40). Long SEO tails are ignored.

# =============================================================================
# GTIN helpers
# =============================================================================

def gtin_valid(digits: str) -> bool:
    if not digits.isdigit() or len(digits) not in (8, 12, 13, 14):
        return False
    body, check = digits[:-1], int(digits[-1])
    total = sum(int(d) * (3 if i % 2 == 0 else 1) for i, d in enumerate(reversed(body)))
    return (10 - total % 10) % 10 == check


def normalize_gtin(v) -> Optional[str]:
    """Any UPC-A / EAN-13 / GTIN-14 / EAN-8 with a valid check digit -> zero-padded GTIN-14."""
    if v is None:
        return None
    d = re.sub(r"\D", "", str(v))
    if len(d) in (11,):           # UPC printed without its leading zero is NOT accepted (ambiguous)
        return None
    if gtin_valid(d):
        return d.zfill(14)
    return None


def gtins_in_text(text: str) -> set:
    """Valid 12-14 digit GTINs appearing as standalone numbers in text/URLs."""
    out = set()
    for m in re.finditer(r"(?<!\d)(\d{12,14})(?!\d)", text or ""):
        g = normalize_gtin(m.group(1))
        if g:
            out.add(g)
    return out


def norm_code(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", s or "").upper()


def looks_like_sku(tok: str) -> bool:
    """True for model-number-like keywords ('RAYG1US1BLK', '609404') but not 'Gen3' / '4 inch'."""
    t = (tok or "").strip()
    if len(t) < 5 or " " in t or not re.search(r"\d", t):
        return False
    if re.fullmatch(r"(?i)gen\d+|v\d+|\d+(?:\.\d+)?(?:ghz|inch|ft|w|v|mm|cm|k|pack|pk)", t):
        return False
    return re.fullmatch(r"[A-Za-z0-9\-_/.]+", t) is not None


def model_codes_in_text(text: str) -> set:
    """Model-code-like tokens (letters AND digits, 6+ chars) - used for near-variant conflicts."""
    out = set()
    for tok in re.findall(r"[A-Za-z0-9][A-Za-z0-9\-]{4,}[A-Za-z0-9]", text or ""):
        c = norm_code(tok)
        if len(c) >= 6 and re.search(r"\d", c) and re.search(r"[A-Z]", c) and not re.fullmatch(r"GEN\d+|\d+(?:GHZ|MHZ|MM|CM|FT|IN|INCH|W|V|K|MAH|PACK|PK)", c):
            out.add(c)
    return out


def _shape(c: str) -> str:
    return re.sub(r"\d", "9", re.sub(r"[A-Z]", "A", c))


def near_variant(a: str, b: str) -> bool:
    """RAYG1US1BLK vs RAYG1US1WHT (colour) / RAYG1EU1BLK (region): same length & shape, long shared
    prefix or suffix, but not equal. Different products, never the same offer."""
    a, b = norm_code(a), norm_code(b)
    if a == b or len(a) < 6 or len(a) != len(b):
        return False
    diff = sum(1 for x, y in zip(a, b) if x != y)
    if diff > 3:
        return False
    prefix = len(os.path.commonprefix([a, b]))
    suffix = len(os.path.commonprefix([a[::-1], b[::-1]]))
    return (prefix + suffix) >= len(a) - 3 and (_shape(a) == _shape(b) or prefix >= 4)


# =============================================================================
# Attribute vocabulary (variant conflicts)
# =============================================================================

ACCESSORY_WORDS = {"case", "cover", "skin", "stand", "bracket", "mount", "mounting", "holder", "protector",
                   "sticker", "decal", "replacement", "cable", "adapter", "charger", "sleeve", "faceplate",
                   "trim", "bezel", "remote", "battery", "batteries", "filter", "strap", "pouch"}
HARD_VARIANT_WORDS = {"pro", "plus", "max", "mini", "lite", "ultra", "xl", "se", "duo"}
SOFT_VARIANT_WORDS = {"ambiance", "essential", "essentials", "outdoor", "portable", "slim", "flex", "dual"}
COLOR_WORDS = {"black", "white", "silver", "gray", "grey", "gold", "blue", "red", "green", "pink", "purple",
               "yellow", "orange", "brown", "beige", "navy", "charcoal", "graphite", "titanium", "bronze",
               "copper", "champagne", "ivory", "cream", "teal", "rose", "walnut", "oak", "chrome", "nickel"}
COLOR_CODES = {"BLK": "black", "WHT": "white", "WH": "white", "BK": "black", "SLV": "silver", "GRY": "gray",
               "GRAY": "gray", "BLU": "blue", "RED": "red", "GRN": "green", "PNK": "pink", "GLD": "gold"}
# Light colour / colour-temperature phrases are NOT product colours ('White and Color Ambiance').
LIGHT_COLOR_PHRASES = re.compile(r"\b(?:white and colou?r|white ambiance|colou?r ambiance|tunable white|soft white|"
                                 r"warm white|cool white|bright white|daylight white|full colou?r|white colou?r|"
                                 r"rgbw?w?|rgbic|colou?r changing|multicolou?r)\b")
GENERIC_WORDS = {"smart", "new", "the", "a", "an", "official", "wireless", "wifi", "zigbee", "bluetooth",
                 "matter", "thread", "z", "wave", "zwave", "compact", "all", "in", "one", "home", "and", "with",
                 "for", "by", "of", "premium", "original", "genuine", "brand", "latest", "version", "edition",
                 "model", "series", "led", "colour", "color", "ambiance", "works", "alexa", "google", "assistant",
                 "apple", "homekit", "hub", "required", "free", "us", "usa", "certified", "inch", "pack", "set",
                 "gen", "generation", "system", "kit", "philips", "signify", "ghz", "mhz", "ft", "w", "v", "mm",
                 "cm", "watt", "watts", "volt", "lumens", "k"} | COLOR_WORDS
CATEGORY_WORDS = {"soundbar", "speaker", "speakers", "plug", "plugs", "outlet", "bulb", "bulbs", "downlight",
                  "downlights", "light", "lights", "lamp", "strip", "bridge", "gateway", "switch", "dimmer",
                  "sensor", "camera", "doorbell", "lock", "thermostat", "controller", "button", "router",
                  "extender", "display", "tv", "television", "subwoofer", "sub", "fixture", "panel", "fan",
                  "vacuum", "purifier", "receiver", "dongle", "adapter", "recessed", "can", "retrofit",
                  "voice", "assistant", "satellite", "repeater", "module", "relay", "hub", "lightbulb"}
BUNDLE_NOUNS = {"sub", "subwoofer", "speaker", "speakers", "soundbar", "bulb", "bulbs", "bridge", "hub", "mount",
                "stand", "remote", "camera", "plug", "plugs", "switch", "sensor", "lamp", "strip", "controller",
                "dimmer", "beam", "arc", "era", "move", "roam", "echo", "dot", "tv"}

# Supplemental protection only (the generic sibling rule below is the main defence).
SUPPLEMENTAL_FAMILY_RULES = {
    "sonos": {"ray", "beam", "arc", "era", "five", "move", "roam", "sub", "ace", "amp", "port",
              "playbar", "playbase", "symfonisk"},
    "hue": {"lightstrip", "bloom", "iris", "signe", "gradient", "centris", "festavia", "datura",
            "filament", "candle", "dimmer"},
    "thirdreality": {"nightlight", "button", "e2", "zp1", "zp2"},
}

_REGION_RX = re.compile(
    r"\b(?:eu|uk|au|nz|jp|cn|euro(?:pean)?|international|intl|imported|import|asian|asia|japan(?:ese)?|china|"
    r"chinese|global|india|german|french|canadian|ca)\s*[-/]?\s*(?:version|model|plug|spec|specs|edition|import|only|"
    r"socket|standard)\b|\b(?:220|230|240)\s*-?\s*(?:240)?\s*v(?:olts?|ac)?\b|\bnon[- ]us\b|\btype[- ]g\b|\buk plug\b", re.I)
_VOLT_RX = re.compile(r"\b(\d{2,3})\s*(?:-\s*\d{2,3}\s*)?v(?:olts?|ac|dc)?\b", re.I)
_SIZE_RX = re.compile(r"\b(\d+(?:\.\d+)?)\s*(inch|in(?!\s+\d)|ft|feet|foot|mm|cm|m(?!\s+\d)|meter|meters|metre|metres)\b")
_TO_MM = {"inch": 25.4, "in": 25.4, "ft": 304.8, "feet": 304.8, "foot": 304.8, "mm": 1.0, "cm": 10.0, "m": 1000.0,
          "meter": 1000.0, "meters": 1000.0, "metre": 1000.0, "metres": 1000.0}
_BUNDLE_RX = re.compile(r"\b(?:bundle|combo|starter\s+(?:kit|set|pack)|value\s+kit)\b", re.I)


_WITH_RX = re.compile(r"(?:(?<=\s)\+|\+(?=\s)|\bw/|\bbundled?\s+with|\bincludes?|\bplus|(?<!compatible )(?<!works )"
                      r"(?<!work )(?<!requires )(?<!require )(?<!use )(?<!connects )(?<!pairs )\bwith)"
                      r"\s*(?:(?:a|an|the|\d+|two|three|four)\s*x?\s+)?([A-Za-z][A-Za-z\- ]{0,30})", re.I)


def _bundle_extra(title: str, itxt_set: set) -> Optional[str]:
    """'Sonos Ray + Sub Mini', 'Hue Bridge with 2 Bulbs', 'Soundbar w/ Wall Mount' -> the extra product noun."""
    for m in _WITH_RX.finditer((title or "")[:140]):
        toks = norm_text(m.group(1)).split()[:3]
        extra = [t for t in toks if t in BUNDLE_NOUNS and t not in itxt_set]
        if extra:
            return extra[0]
    return None


def _strip_universal(text: str) -> str:
    """'100-240V' universal power supplies are sold in the US - not a regional variant."""
    return re.sub(r"\b1[0-2]0\s*[-~/]\s*2[2-4]0\s*v(?:ac)?\b", " ", text or "", flags=re.I)


def _sizes(t_norm: str) -> set:
    out = set()
    for v, unit in _SIZE_RX.findall(t_norm):
        try:
            out.add(round(float(v) * _TO_MM[unit], 1))
        except (KeyError, ValueError):
            pass
    return out


def _gens(t_norm: str) -> set:
    g = set(re.findall(r"\bgen(\d+)\b", t_norm))
    g |= {m for m in re.findall(r"\b(?:v|mk|mark|version)\s?(\d)\b", t_norm)}
    return g


def _colors(text: str) -> set:
    t = LIGHT_COLOR_PHRASES.sub(" ", norm_text(text))
    return {w for w in t.split() if w in COLOR_WORDS} - {"rose"} | ({"rose gold"} if "rose gold" in t else set())


def _code_colors(codes: Iterable[str]) -> set:
    out = set()
    for c in codes:
        c = norm_code(c)
        for code, col in sorted(COLOR_CODES.items(), key=lambda kv: -len(kv[0])):
            if len(c) >= 6 and c.endswith(code) and re.search(r"\d", c[:-len(code)]):
                out.add(col)
                break
    return out


def _volts(text: str) -> set:
    return {int(v) for v in _VOLT_RX.findall(text or "") if 3 <= int(v) <= 480}


# =============================================================================
# Item identity
# =============================================================================

def item_brand(item: Item) -> str:
    b = norm_text(item.brand or item.learned_brand)
    toks = norm_text(item.product).split()
    if b:
        # 'Philips' + 'Hue Color ...' -> the name's own first token is still the brand-line marker
        return b.split()[0] if b.split()[0] in toks else (toks[0] if toks else b.split()[0])
    return toks[0] if toks else ""


def item_model_token(item: Item) -> Optional[str]:
    """The token that names the model: digit-bearing tokens win ('gen3'), else the first distinctive
    token after the brand ('ray' in 'Sonos Ray Soundbar', 'slim' in 'Hue Color Slim Downlight')."""
    toks = norm_text(item.product).split()
    brand = item_brand(item)
    rest = toks[1:] if toks and toks[0] == brand else toks
    for t in rest:
        if re.search(r"[a-z]", t) and re.search(r"\d", t):
            return t
    for t in rest:
        if t not in GENERIC_WORDS and t not in CATEGORY_WORDS and not t.isdigit() and len(t) >= 2:
            return t
    return None


def item_colors(item: Item) -> set:
    cols = set()
    for s in item.specs + [k for k in item.keywords if not looks_like_sku(k)]:
        cols |= _colors(s)
    return cols | _code_colors(item.all_mpns)


def item_text(item: Item) -> str:
    return " ".join([item.product] + list(item.specs) + [k for k in item.keywords if not looks_like_sku(k)])


def excluded_phrase(item: Item, title: str = "", mpns: Iterable[str] = (), slug: str = "") -> Optional[str]:
    """First '!phrase' (Master Sheet specs / keywords) that this listing contains, else None.
    Whole-word, in-order match on the normalised title ('Lite' never matches 'Satellite'); a model-code-like
    phrase also matches the listing's MPNs / a compact title match. URL-slug words (order is lost) are only
    tested for single-word phrases. A phrase made only of words of the product's own name is ignored (a
    contradictory sheet row must not block every listing)."""
    if not item.exclude:
        return None
    t_toks = norm_text(title).split()
    t_compact = norm_code(title)
    slug_set = set(norm_text(slug).split())
    name_set = set(norm_text(item.product).split())
    codes = {norm_code(m) for m in mpns}
    for phrase in item.exclude:
        p_toks = norm_text(phrase).split()
        if not p_toks or set(p_toks) <= name_set:
            continue
        n = len(p_toks)
        if any(t_toks[i:i + n] == p_toks for i in range(len(t_toks) - n + 1)):
            return phrase
        code = norm_code(phrase)
        if looks_like_sku(phrase) and (code in codes or (len(code) >= 5 and code in t_compact)):
            return phrase
        if n == 1 and p_toks[0] in slug_set:
            return phrase
    return None


def fingerprint(brand: str, model: str, gtin: str = "", attrs: str = "", pack: int = 1) -> str:
    """Canonical product fingerprint: brand + model/MPN + GTIN + material attributes + pack."""
    raw = "|".join([norm_text(brand), norm_code(model), gtin or "", norm_text(attrs), str(pack)])
    return hashlib.sha1(raw.encode()).hexdigest()[:12]


def item_fingerprint(item: Item, pack: int = 1) -> str:
    g = sorted(item.all_gtins)[0] if item.all_gtins else ""
    m = item.all_mpns[0] if item.all_mpns else item.product
    return fingerprint(item_brand(item), m, g, " ".join(sorted(item_colors(item))), pack)


# =============================================================================
# Matching
# =============================================================================

@dataclass
class Match:
    confidence: str          # High | Medium | Low
    reason: str
    evidence: str = ""       # gtin | mpn | page_metadata | title | ''

    def __iter__(self):      # allows `conf, reason = classify_confidence(...)` (backwards compatible)
        return iter((self.confidence, self.reason))

    def __getitem__(self, i):
        return (self.confidence, self.reason)[i]


def _name_present(tok: str, head_set: set, head_compact: str) -> bool:
    return tok in head_set or (len(tok) >= 6 and tok in head_compact)


def _in_order(name_tokens: list, head_tokens: list, head_set: set, head_compact: str) -> bool:
    pos = 0
    for tok in name_tokens:
        if tok in head_tokens[pos:]:
            pos = head_tokens.index(tok, pos) + 1
        elif _name_present(tok, head_set, head_compact):
            continue
        else:
            return False
    return True


def _for_product(title: str, item: Item) -> bool:
    """'Wall Mount compatible with Sonos Ray' / 'Case for Hue ...' -> accessory listing."""
    brand = item_brand(item)
    model = item_model_token(item) or ""
    keys = [k for k in {brand, model} if k]
    if not keys:
        return False
    t = norm_text(title)
    rx = re.compile(r"\b(?:compatible with|for use with|designed for|fits|replacement for|made for|for)\s+(?:the\s+)?"
                    r"(?:\w+\s+){0,2}?(?:" + "|".join(re.escape(k) for k in keys) + r")\b")
    m = rx.search(t)
    if not m:
        return False
    lead = t[:m.start()].split()
    return not (set(lead[:3]) & {brand, model}) or bool(set(lead) & ACCESSORY_WORDS)


def conflicts(item: Item, title: str, gtins: set = frozenset(), mpns: set = frozenset(), pack_qty: int = 1,
              color: str = "", allow_codes: Iterable[str] = ()) -> Optional[str]:
    """Return a human-readable hard-conflict reason, or None."""
    t_norm = norm_text(strip_pack_phrases(title))
    name_norm = norm_text(item.product)
    name_set = set(name_norm.split())
    itxt = norm_text(item_text(item))
    itxt_set = set(itxt.split())
    head_len = max(HEAD_MIN_CHARS, len(name_norm) + HEAD_EXTRA_CHARS)
    head = t_norm[:head_len]
    head_set = set(head.split())
    allowed_codes = {norm_code(c) for c in allow_codes}

    # --- '!' exclusions from the Master Sheet (Product specifications / Search Keywords) ---------------
    ex = excluded_phrase(item, title, mpns)
    if ex:
        return f"excluded by '!{ex}' in the Master Sheet"

    # --- identifiers -------------------------------------------------------------------------------
    i_gtins = item.all_gtins
    l_gtins = set(gtins) | gtins_in_text(title)
    if i_gtins and l_gtins and not (i_gtins & l_gtins) and pack_qty == 1:
        return f"GTIN mismatch ({sorted(l_gtins)[0].lstrip('0')})"
    item_codes = [norm_code(m) for m in item.all_mpns if len(norm_code(m)) >= 6]
    l_codes = {norm_code(m) for m in mpns} | model_codes_in_text(title)
    for ic in item_codes:
        if ic in l_codes or ic in norm_code(title):
            continue
        for lc in l_codes:
            # a listing carrying another of the item's OWN model codes (e.g. 1-pack P1SPD1Z vs 4-pack
            # P1SPD4Z, both listed in the MPN cell) is the same product line, not a different variant
            if lc in item_codes:
                continue
            if lc not in allowed_codes and near_variant(ic, lc):
                return f"different variant of model {ic} ({lc})"

    # --- accessories / bundles ------------------------------------------------------------------------
    acc = (head_set & ACCESSORY_WORDS) - itxt_set
    if acc:
        return f"accessory word '{sorted(acc)[0]}'"
    if _for_product(title, item):
        return "accessory 'for/compatible with' listing"
    if _BUNDLE_RX.search(title or "") and not _BUNDLE_RX.search(item_text(item)):
        return "bundle"
    extra = _bundle_extra(title, itxt_set)
    if extra:
        return f"bundle (+ {extra})"

    # --- region ---------------------------------------------------------------------------------------
    reg = _REGION_RX.search(_strip_universal(title))
    if reg and not _REGION_RX.search(_strip_universal(item_text(item))):
        return f"non-US / regional variant ('{reg.group(0).strip()}')"

    # --- generation / version -------------------------------------------------------------------------
    ig, lg = _gens(itxt), _gens(t_norm)
    if ig and lg and not (ig & lg):
        return "different generation/version"

    # --- colour ---------------------------------------------------------------------------------------
    icol = item_colors(item)
    lcol = (_colors(title) - _colors(item.product)) | ({c for c in _colors(color)} if color else set()) | _code_colors(l_codes)
    if icol and lcol and not (icol & lcol):
        return f"different colour '{sorted(lcol)[0]}'"

    # --- voltage / dimensions ------------------------------------------------------------------------
    iv, lv = _volts(item_text(item)), _volts(title)
    if iv and lv and not (iv & lv):
        return f"different voltage ({sorted(lv)[0]}V)"
    isz, lsz = _sizes(itxt), _sizes(t_norm)
    if isz and lsz and not any(abs(a - b) <= max(1.0, 0.04 * a) for a in isz for b in lsz):
        return "different size/dimensions"

    # --- configuration / sibling models ---------------------------------------------------------------
    hv = (head_set & HARD_VARIANT_WORDS) - itxt_set
    if hv:
        return f"different configuration '{sorted(hv)[0]}'"
    brand = item_brand(item)
    model = item_model_token(item)
    toks = head.split()
    if brand and model and brand in toks and model not in head_set and not (len(model) >= 6 and model in head.replace(" ", "")):
        after = [t for t in toks[toks.index(brand) + 1:toks.index(brand) + 5]
                 if t not in GENERIC_WORDS and not t.isdigit() and t not in CATEGORY_WORDS]
        other = [t for t in after if t not in itxt_set and len(t) >= 2]
        if other:
            return f"different {brand} model '{other[0]}'"
    for fam, models in SUPPLEMENTAL_FAMILY_RULES.items():
        if fam in name_set or fam in name_norm.replace(" ", ""):
            other = (head_set & models) - itxt_set
            if other:
                return f"different {fam} model '{sorted(other)[0]}'"
    return None


def classify(item: Item, title: str, anchor_head: Optional[str] = None, *, gtins: set = frozenset(),
             mpns: set = frozenset(), brand: str = "", pack_qty: int = 1, color: str = "") -> Match:
    """Is this listing the exact product? Identifier-first; see module docstring."""
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

    # Positive identifiers first (they may also legitimise a model code that would otherwise look like a variant)
    l_gtins = set(gtins) | gtins_in_text(title)
    gtin_hit = bool(item.all_gtins & l_gtins)
    l_codes = {norm_code(m) for m in mpns}
    mpn_hit = any(norm_code(m) and (norm_code(m) in l_codes or norm_code(m).lower() in t_compact)
                  for m in item.all_mpns if len(norm_code(m)) >= 4)

    bad = conflicts(item, title, gtins, mpns, pack_qty, color)
    if bad:
        return Match("Low", bad)

    if gtin_hit:
        return Match("High", "GTIN match", "gtin")

    present = [_name_present(tok, head_set, head_compact) for tok in name_tokens]
    coverage = (sum(present) / len(name_tokens)) if name_tokens else 0.0
    exact = coverage == 1.0 and _in_order(name_tokens, head_tokens, head_set, head_compact)
    ratio = difflib.SequenceMatcher(None, name_norm, head[:len(name_norm) + 10]).ratio()
    if coverage < 1.0 and ((head_set & (HARD_VARIANT_WORDS | SOFT_VARIANT_WORDS)) - set(name_tokens)):
        return Match("Low", "variant word present and name word missing")
    missing_model = [t for t in name_tokens if re.search(r"[a-z]", t) and re.search(r"\d", t)
                     and t not in t_set and t not in t_compact]
    if missing_model and not mpn_hit:
        return Match("Low", f"model '{missing_model[0]}' not in title")

    specs_ok = True
    for spec in item.spec_phrases:
        toks = norm_text(spec).split()
        if toks and not all(t in t_set or (len(t) >= 5 and t in t_compact) for t in toks):
            specs_ok = False
            break

    if mpn_hit:
        src = "page_metadata" if l_codes and any(norm_code(m) in l_codes for m in item.all_mpns) else "mpn"
        return Match("High", "SKU/MPN + specs match", src) if specs_ok else \
            Match("Medium", "SKU/MPN match, some specs not in title", src)
    if exact:
        if specs_ok:
            return Match("High", "exact name, specs match" if item.spec_phrases else "exact name, no specs defined", "title")
        return Match("Medium", "exact name, specs not in title", "title")
    if anchor_head and specs_ok and coverage >= 0.75:
        if difflib.SequenceMatcher(None, anchor_head[:head_len], head).ratio() >= 0.8:
            return Match("High", "mirrors SKU-verified listing", "title")
    if coverage >= 0.75 or ratio >= 0.85:
        return Match("Medium", "strong name match", "title")
    return Match("Low", "weak name match", "title")


def classify_confidence(item: Item, title: str, anchor_head: Optional[str] = None) -> Match:
    """Backwards-compatible wrapper (title only)."""
    return classify(item, title, anchor_head)


def anchor_from(item: Item, title: str) -> str:
    return norm_text(strip_pack_phrases(title))[:max(HEAD_MIN_CHARS, len(item.product) + HEAD_EXTRA_CHARS)]


def has_identifier_hit(item: Item, title: str, gtins: set = frozenset(), mpns: set = frozenset()) -> bool:
    t_compact = norm_text(title).replace(" ", "")
    if item.all_gtins & (set(gtins) | gtins_in_text(title)):
        return True
    l_codes = {norm_code(m) for m in mpns}
    return any(norm_code(m) and (norm_code(m) in l_codes or norm_code(m).lower() in t_compact)
               for m in item.all_mpns if len(norm_code(m)) >= 4)


# =============================================================================
# Validating a trusted-candidate page (Master Sheet Product URL / cached discovery)
# =============================================================================

def brand_owns_domain(item: Item, host: str) -> bool:
    """True when the page's host is the item's OWN brand's site (samsung.com for a Samsung TV, us.govee.com for a
    Govee strip). Needs the Brand named in the Master Sheet (or learned from a validated page) - the first word of
    the product name is not trusted for this."""
    comp = re.sub(r"[^a-z0-9]", "", (host or "").lower())
    brand = norm_text(item.brand or item.learned_brand)
    cands = {brand.replace(" ", ""), brand.split()[0] if brand else ""}
    return bool(comp) and any(len(b) >= 4 and b in comp for b in cands)


def validate_page(item: Item, title: str, gtins: set, mpns: set, brand_hint: str = "", domain_brand: str = "",
                  slug: str = "", pack_qty: int = 1, color: str = "", host: str = "") -> Match:
    """See _validate_page. A page on the brand's own site that is not contradicted by anything (no hard conflict:
    right model code, size, colour, pack ...) is the product: an unconfirmed (Medium) page there is promoted to High."""
    m = _validate_page(item, title, gtins, mpns, brand_hint, domain_brand, slug, pack_qty, color)
    if m.confidence == "Medium" and brand_owns_domain(item, host):
        return Match("High", f"Product URL on the brand's own site ({host}), no conflicts - {m.reason}", "brand_site")
    return m


def _validate_page(item: Item, title: str, gtins: set, mpns: set, brand_hint: str = "", domain_brand: str = "",
                   slug: str = "", pack_qty: int = 1, color: str = "") -> Match:
    """A Product URL is a trusted CANDIDATE, not proof. Accept its page as the product when:
         * an identifier matches (GTIN / MPN) and nothing conflicts            -> High (gtin / mpn)
         * or the page title (plus URL slug) carries the distinctive name words -> High (product_url)
           (brand implied by the domain and generic category words like 'soundbar' may be absent)
       Hard conflicts -> Low (identity mismatch). Weak evidence -> Medium (does NOT establish trusted price)."""
    bad = conflicts(item, title, gtins, mpns, pack_qty, color)
    if not bad:
        ex = excluded_phrase(item, "", slug=slug)
        bad = f"excluded by '!{ex}' in the page URL" if ex else None
    if bad:
        return Match("Low", f"Product URL page conflicts: {bad}")
    full = classify(item, title, gtins=gtins, mpns=mpns, pack_qty=pack_qty, color=color)
    if full.confidence == "High":
        return Match("High", f"Product URL validated ({full.reason})", full.evidence or "title")
    if full.confidence == "Low" and full.reason.startswith(("model '", "variant word")):
        return full
    text = norm_text(f"{title} {slug}")
    tset, tcomp = set(text.split()), text.replace(" ", "")
    implied = set(norm_text(domain_brand + " " + brand_hint).split())
    toks = [t for t in norm_text(item.product).split()
            if t not in implied and t not in CATEGORY_WORDS and t not in {"smart", "the", "and", "new"}]
    if not toks:
        toks = norm_text(item.product).split()
    have = [t for t in toks if t in tset or (len(t) >= 4 and t in tcomp)]
    digit_toks = [t for t in toks if re.search(r"\d", t) and re.search(r"[a-z]", t)]
    if all(t in have for t in digit_toks) and len(have) / len(toks) >= 0.75:
        return Match("High", "Product URL validated (title/slug match)", "product_url")
    return Match("Medium", f"Product URL identity unconfirmed ({full.reason})", full.evidence)
