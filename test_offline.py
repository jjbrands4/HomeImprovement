#!/usr/bin/env python3
"""
test_offline.py - verifies price_tracker.py end-to-end WITHOUT internet or API keys.

It copies your real workbook to a temp folder, feeds fake-but-realistic SerpApi / eBay / direct-site
responses through the real scoring + deal logic, and checks:
  * matching rules (packs, accessories, generation/colour variants, SKU anchor)
  * deal rules (15% below pool average, 30-day trend) and the separate New / Resale pools
  * Run Data + Deals Data are written, and NO other worksheet is changed
  * the script survives with no keys / no network (source failures become "N/A" notes)

Run:  python test_offline.py            (from the repo folder containing the workbook)
"""
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

import price_tracker as pt

SRC = next(iter(sorted(Path(".").glob(pt.WORKBOOK_GLOB))), None)
if SRC is None:
    sys.exit("Put this file next to your workbook (Home Wishlist.xlsx) and re-run.")

# Work on a FIXTURE copy so the tests don't depend on what's currently in your real Master Sheet:
# item 4 (MiBoxer, priority 1) is added if missing; it is used for the cadence + matching tests.
_fx_dir = Path(tempfile.mkdtemp())
FIXTURE = _fx_dir / "Home Wishlist.xlsx"
shutil.copy(SRC, FIXTURE)
_fx = load_workbook(FIXTURE)
_ms = _fx[pt.SHEET_MASTER]
_hm = pt.header_map(_ms)
if not any(str(_ms.cell(r, _hm["wishlistitem"]).value).strip() == "4" for r in range(2, _ms.max_row + 1)):
    _r = _ms.max_row + 1
    for k, v in {"wishlistitem": 4, "product": "MiBoxer 2.4 GHz WiFi Gateway", "location": "Overall",
                 "priority1low5high": 1, "quantityneeded": 1, "targetpriceperitemusd": 25,
                 "openusedorhighqualityrefurbished": "No", "openusedorhighqualityrefurbished?": "No",
                 "isbulkoption": "No"}.items():
        c = pt.find_col(_hm, k)
        if c:
            _ms.cell(_r, c, v)
_fx.save(FIXTURE)
SRC = FIXTURE

passed = failed = 0


def check(label, cond, extra=""):
    global passed, failed
    if cond:
        passed += 1
        print(f"  PASS  {label}")
    else:
        failed += 1
        print(f"  FAIL  {label}  {extra}")


# ---------------------------------------------------------------------------
print("\n[1] Pure-function checks")
check("pack: '2 Pack'", pt.extract_pack_qty("Plug 2 Pack") == 2)
check("pack: 'Pack of 4'", pt.extract_pack_qty("Plug, Pack of 4") == 4)
check("pack: '4-pack' / '2pk' / 'six pack'", (pt.extract_pack_qty("4-pack"), pt.extract_pack_qty("2pk"),
                                              pt.extract_pack_qty("six pack")) == (4, 2, 6))
check("pack: none -> 1", pt.extract_pack_qty("Sonos Ray Soundbar") == 1)
check("bulk sizes '2Pack; 4 Pack'", pt.parse_bulk_sizes("2Pack; 4 Pack") == {2, 4})
check("vendor: 'The Home Depot' ~ 'Home Depot'", pt.keys_match(pt.vendor_key("The Home Depot"), pt.vendor_key("Home Depot")))
check("vendor: 'B&H' ~ 'B&H Photo Video'", pt.keys_match(pt.vendor_key("B&H"), pt.vendor_key("B&H Photo Video")))
check("vendor: 'from eBay' ~ 'ebay'", pt.keys_match(pt.vendor_key("from eBay"), pt.vendor_key("ebay")))
check("vendor: 'Amazon.com - Seller' ~ 'Amazon'", pt.keys_match(pt.vendor_key("Amazon.com - Seller"), pt.vendor_key("Amazon")))
check("vendor: Hue != Home Depot", not pt.keys_match(pt.vendor_key("Philips Hue"), pt.vendor_key("Home Depot")))

wb0 = load_workbook(SRC)
items = {str(i.wid): i for i in pt.load_items(wb0[pt.SHEET_MASTER])}
hue, plug, sonos, mibox = items["1"], items["2"], items["3"], items["4"]
check("Master Sheet parsed 4 items", len(items) == 4)
check("Hue SKU detected, '4 inch' treated as spec", hue.skus == ["609404"] and hue.spec_phrases == ["4 inch"])
check("Sonos SKU detected", sonos.skus == ["RAYG1US1BLK"])
check("ThirdReality bulk sizes {2,4}", plug.bulk and plug.bulk_sizes == {2, 4})

cc = pt.classify_confidence
check("Hue: SKU + '4-inch' -> High", cc(hue, 'Philips Hue White and Color Ambiance 4" Slim Downlight 609404')[0] == "High")
check("Hue: exact name, no SKU -> High", cc(hue, "Philips Hue Color Slim Downlight 4-inch")[0] == "High")
check("Hue: White Ambiance (not Color) -> Low", cc(hue, "Philips Hue White Ambiance Slim Downlight 4 inch")[0] == "Low")
check("Hue: accessory 'Case for' -> Low", cc(hue, "Case for Hue Color Slim Downlight 4-inch")[0] == "Low")
check("Plug: 'Zigbee' inserted in name, still exact -> High",
      cc(plug, "ThirdReality Zigbee Smart Plug Gen3, Power Meter, Works with Home Assistant")[0] == "High")
check("Plug: pack phrase ignored -> High", cc(plug, "THIRDREALITY Smart Plug Gen3 4 Pack, Precise Real-time Power Meter")[0] == "High")
check("Plug: Gen2 -> Low", cc(plug, "THIRDREALITY Smart Plug Gen2")[0] == "Low")
check("Plug: older model 'Smart Plug E2' -> Low", cc(plug, "THIRDREALITY Smart Plug E2")[0] == "Low")
check("Plug: 'Smart Dual Plug ZP1' -> Low", cc(plug, "THIRDREALITY Smart Dual Plug ZP1")[0] == "Low")
check("Plug: other brand -> Low", cc(plug, "Kasa Smart Plug Mini 15A")[0] == "Low")
check("Sonos: SKU but 'Unmounted' missing -> Medium", cc(sonos, "Sonos Ray Compact Soundbar Black RAYG1US1BLK")[0] == "Medium")
check("Sonos: white variant SKU -> Low", cc(sonos, "Sonos Ray Soundbar White RAYG1US1WHT")[0] == "Low")
check("Sonos: wall mount -> Low", cc(sonos, "Sonos Ray Wall Mount Bracket")[0] == "Low")
check("MiBoxer: near-exact -> Medium or High", cc(mibox, "MiBoxer WL-Box2 2.4GHz WiFi Gateway")[0] in ("Medium", "High"))

# ---------------------------------------------------------------------------
print("\n[2] End-to-end with fake sources (real scoring, deal rules, and workbook writing)")


def S(title, price, source, **kw):          # fake SerpApi Google Shopping row
    return {"title": title, "extracted_price": price, "source": source,
            "product_link": "https://www.google.com/shopping/product/123", **kw}


def E(title, price, cond, pct, cnt, user="seller1"):   # fake eBay Browse API row
    return {"title": title, "price": {"value": str(price), "currency": "USD"}, "condition": cond,
            "itemWebUrl": "https://www.ebay.com/itm/1", "seller": {"username": user,
            "feedbackPercentage": str(pct), "feedbackScore": cnt}}


FAKE_SERP = {
    "hue": [
        S('Philips Hue White and Color Ambiance 4" Slim Downlight 609404', 69.99, "Amazon.com"),
        S("Philips Hue Color Slim Downlight 4-inch", 64.99, "Best Buy"),
        S("Philips Hue Color Slim Downlight 4 inch 609404", 72.00, "The Home Depot"),
        S("Philips Hue Color Slim Downlight 4-inch", 66.00, "Crutchfield"),
        S("Philips Hue Color Slim Downlight 4-inch", 49.99, "B&H"),                # expected primary deal
        S("Philips Hue Color Slim Downlight 4-inch (4 Pack)", 250.00, "Amazon.com"),  # pack not allowed
        S("Case for Hue Color Slim Downlight 4-inch", 12.00, "Etsy"),             # accessory
        S("Philips Hue White Ambiance Slim Downlight 4 inch", 40.00, "Best Buy"),  # wrong model
    ],
    "thirdreality": [
        S("THIRDREALITY Smart Plug Gen3 4 Pack, Precise Real-time Power Meter", 39.96, "Amazon.com"),  # unit 9.99
        S("ThirdReality Zigbee Smart Plug Gen3", 14.99, "Best Buy"),
        S("THIRDREALITY Smart Plug Gen3", 13.99, "Home Depot"),
        S("THIRDREALITY Smart Plug Gen3", 12.00, "Zigbee Shop"),                   # unlisted: baseline only (open_used = No)
        S("THIRDREALITY Smart Plug Gen3 Pack of 2", 22.00, "Amazon.com - Seller"),  # unit 11.00
        S("THIRDREALITY Smart Plug Gen2", 9.00, "Amazon.com"),                      # wrong generation
        S("THIRDREALITY Smart Plug Gen3 6 Pack", 60.00, "Amazon.com"),              # 6 not in bulk keywords
    ],
    "sonos": [
        S("Sonos Ray Compact Soundbar Black RAYG1US1BLK", 169.00, "Amazon.com"),
        S("Sonos Ray Compact Soundbar RAYG1US1BLK", 179.00, "Best Buy"),
        S("Sonos Ray Soundbar RAYG1US1BLK Unmounted", 179.00, "Sonos"),
        S("Sonos Ray Soundbar White RAYG1US1WHT", 149.00, "Best Buy"),             # colour variant
        S("Sonos Ray Wall Mount", 29.00, "Amazon.com"),                            # accessory
        S("Sonos Ray Soundbar RAYG1US1BLK - Renewed", 120.00, "Amazon.com"),       # resale pool
    ],
    "miboxer": [
        S("MiBoxer WL-Box2 2.4GHz WiFi Gateway", 24.99, "Amazon.com"),
        S("MiBoxer 2.4 GHz WiFi Gateway", 29.99, "SuperBrightLEDs"),
        S("MiBoxer 2.4 GHz WiFi Gateway", 27.50, "Best Buy"),
        S("MiBoxer 2.4 GHz WiFi Gateway", 17.00, "Home Depot"),
    ],
}
FAKE_EBAY = {
    "hue": [
        E('Philips Hue Color Slim Downlight 4" 609404', 41.00, "Used", 99.4, 500),
        E("Philips Hue Color Slim Downlight 4-inch", 44.00, "Used", 99.0, 900),
        E("Philips Hue Color Slim Downlight 4-inch 609404", 46.00, "Open box", 98.0, 300),
        E("Philips Hue Color Slim Downlight 4-inch", 29.99, "Open box", 98.5, 1200, "gooddeals"),  # resale deal
        E("Philips Hue Color Slim Downlight 4-inch", 25.00, "Used", 80.0, 40, "shadyseller"),     # low feedback
    ],
    "sonos": [
        E("Sonos Ray Soundbar RAYG1US1BLK", 130.00, "Used", 99.0, 800),
        E("Sonos Ray Soundbar RAYG1US1BLK", 135.00, "Used", 99.0, 700),
        E("Sonos Ray Soundbar RAYG1US1BLK", 140.00, "Used", 99.5, 2000),
        E("Sonos Ray Soundbar RAYG1US1BLK", 95.00, "Used", 99.5, 2000),   # <0.4x... stays plausible but tests outlier math
    ],
}


def key_of(query):
    q = query.lower()
    return next(k for k in ("hue", "thirdreality", "sonos", "miboxer") if k in q)


def fake_direct(vendor, query, session):
    k = key_of(query)
    if k == "thirdreality":      # sold-out on vendor's site -> must be excluded
        return [pt.Listing("THIRDREALITY Smart Plug Gen3", "https://thirdreality.com/products/smart-plug-gen3",
                           14.99, vendor.name, "direct", in_stock=False)], f"Direct[{vendor.name}]: 1 listings (fake)"
    if k == "sonos":
        return [pt.Listing("Sonos Ray Soundbar Black RAYG1US1BLK", "https://www.sonos.com/en-us/shop/ray",
                           179.00, vendor.name, "direct", in_stock=True)], f"Direct[{vendor.name}]: 1 listings (fake)"
    return [], f"Direct[{vendor.name}]: N/A (fake)"


calls = {"serp": 0, "ebay": 0, "direct": []}


def fake_shopping(self, query):
    calls["serp"] += 1
    return FAKE_SERP[key_of(query)]


def fake_ebay(self, query):
    calls["ebay"] += 1
    return FAKE_EBAY.get(key_of(query), [])


_orig = (pt.SerpApiClient.shopping, pt.SerpApiClient.check_credits, pt.EbayClient.search)
pt.SerpApiClient.shopping = fake_shopping
pt.SerpApiClient.check_credits = lambda self: None
pt.EbayClient.search = fake_ebay
_orig_direct = pt.direct_vendor_listings
pt.direct_vendor_listings = lambda v, it, s: (calls["direct"].append(v.name), fake_direct(v, it.product, s))[1]
pt.time.sleep = lambda *_: None
os.environ.update({"SERPAPI_KEY": "dummy", "EBAY_CLIENT_ID": "dummy", "EBAY_CLIENT_SECRET": "dummy"})

tmp = Path(tempfile.mkdtemp())
wbpath = tmp / "Home Wishlist.xlsx"
shutil.copy(SRC, wbpath)

# Seed 30-day history: 3 past runs for Sonos (item 3, priority 2 -> 6-day cadence) 8-16 days ago averaging $185,
# so the trend rule can fire; plus one run 1 day ago for item 4 (priority 1, 9-day cadence) to prove the cadence skip.
wbs = load_workbook(wbpath)
hm = pt.ensure_columns(wbs[pt.SHEET_RUN], pt.RUN_COLS)
seed = []
for days in (8, 12, 16):
    seed.append({"RowDateTime": datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=days), "runID": "seed", "WishlistItem": 3,
                 "Product": "Sonos Ray Soundbar", "Avg Price (New)": 185.0})
seed.append({"RowDateTime": datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=1), "runID": "seed", "WishlistItem": 4,
             "Product": "MiBoxer 2.4 GHz WiFi Gateway", "Avg Price (New)": 26.0})
pt.append_rows(wbs[pt.SHEET_RUN], pt.RUN_COLS, seed)
wbs.save(wbpath)
before = load_workbook(wbpath)            # snapshot AFTER seeding, to prove nothing else changes

rc = pt.main(["--workbook", str(wbpath)])  # no --force: item 4 must be skipped by cadence
check("main() exit code 0", rc == 0)

after = load_workbook(wbpath)
run, deals = after[pt.SHEET_RUN], after[pt.SHEET_DEALS]
rh = pt.header_map(run)
rows = [{h: run.cell(r, c).value for h, c in rh.items()} for r in range(2, run.max_row + 1)]
new_rows = [r for r in rows if r[pt.norm_header("runID")] != "seed"]
by_item = {str(r[pt.norm_header("WishlistItem")]): r for r in new_rows}
check("3 new Run Data rows; priority-1 item 4 skipped by cadence", len(new_rows) == 3 and "4" not in by_item, str(list(by_item)))

dh = pt.header_map(deals)
drows = [{h: deals.cell(r, c).value for h, c in dh.items()} for r in range(2, deals.max_row + 1)]
g = lambda d, name: d.get(pt.norm_header(name))
for d in drows:
    print(f"     deal: item {g(d,'WishlistItem')} ${g(d,'Price')} {g(d,'Vendor')} primary={g(d,'isPrimaryVendor')} "
          f"conf={g(d,'RightProductConfidence')} resale={g(d,'Secondary Vendor?')} | {g(d,'Deal Rule')}")

hue_d = [d for d in drows if str(g(d, "WishlistItem")) == "1"]
plug_d = [d for d in drows if str(g(d, "WishlistItem")) == "2"]
sonos_d = [d for d in drows if str(g(d, "WishlistItem")) == "3"]
check("Hue: B&H $49.99 flagged as PRIMARY deal and listed first",
      hue_d and g(hue_d[0], "Price") == 49.99 and g(hue_d[0], "isPrimaryVendor") == 1)
check("Hue: eBay resale deal $29.99 has seller comment with % and count",
      any(g(d, "Price") == 29.99 and "98.5% positive" in (g(d, "Secondary Vendor Comments") or "")
          and "1,200" in (g(d, "Secondary Vendor Comments") or "") and g(d, "Secondary Vendor?") == "Yes" for d in hue_d))
check("Hue: low-feedback eBay seller ($25) NOT reported", all(g(d, "Price") != 25.0 for d in hue_d))
check("Hue: 4-pack, accessory and wrong-model listings never reported",
      all(g(d, "Price") not in (250.0, 12.0, 40.0, 62.5) for d in hue_d))
check("Hue: primary deals sorted before resale", [g(d, "isPrimaryVendor") for d in hue_d][:1] == [1])
check("Plug: Amazon 4-pack reported at per-item $9.99 (Pack Qty 4, Listed 39.96)",
      any(g(d, "Price") == 9.99 and g(d, "Pack Qty") == 4 and g(d, "Listed Price") == 39.96 for d in plug_d))
check("Plug: unlisted vendor 'Zigbee Shop' $12 NOT reported (item not open to wider search)",
      all(g(d, "Vendor") != "Zigbee Shop" for d in plug_d))
check("Plug: sold-out direct listing excluded", all(g(d, "Source") != "direct" for d in plug_d))
check("Plug: no resale rows for an item not open to used", all(g(d, "Secondary Vendor?") == "No" for d in plug_d))
check("Sonos: trend rule fires (below 30d trend $185)", any("30d trend" in (g(d, "Deal Rule") or "") for d in sonos_d))
check("Sonos: white variant $149 + wall mount $29 never reported", all(g(d, "Price") not in (149.0, 29.0) for d in sonos_d))
check("Sonos: direct Sonos listing marked In Stock Verified = Yes",
      any(g(d, "Source") == "direct" and g(d, "In Stock Verified") == "Yes" for d in sonos_d) or
      not any(g(d, "Source") == "direct" for d in sonos_d))

r1 = by_item["1"]
check("Run Data: Hue row has averages stored for trend",
      g(r1, "Avg Price (New)") and g(r1, "Avg Price (Resale)"), str(r1))
check("Run Data: Hue 'Target Price or better found' counts $49.99 + $35-ish listings <= $60",
      (g(r1, "Target Price or better found") or 0) >= 1)
check("Run Data: Sonos trend column populated from history", g(by_item["3"], "30d Trend (New)") == 185.0)
check("Run Data: Matching Listings recorded", all(isinstance(g(r, "Matching Listings"), int) for r in new_rows))
check("Run Data: notes list SerpApi/eBay/Direct outcomes", "SerpApi" in g(r1, "Source Notes") and "eBay" in g(r1, "Source Notes"))
check("Direct search only for name-matching DTC vendors (Hue, Sonos, ThirdReality)",
      sorted(set(calls["direct"])) == ["Philips Hue Direct", "Sonos Direct", "THIRDREALITY Store"], str(calls["direct"]))
check("eBay only queried for items open to used (Hue, Sonos)", calls["ebay"] == 2, str(calls["ebay"]))

untouched = [n for n in before.sheetnames if n not in (pt.SHEET_RUN, pt.SHEET_DEALS)]
same = all([[c.value for c in row] for row in before[n].iter_rows()] ==
           [[c.value for c in row] for row in after[n].iter_rows()] for n in untouched)
check(f"No other worksheet changed ({', '.join(untouched)})", same)
check("Deals Data keeps original 8 headers in place",
      [deals.cell(1, c).value for c in range(1, 9)] ==
      ["runID", "WishlistItem", "Product", "URL", "isPrimaryVendor", "Target Price", "Price", "RightProductConfidence"])

# ---------------------------------------------------------------------------
print("\n[3] --dry-run leaves the file byte-identical")
h1 = wbpath.read_bytes()
pt.main(["--workbook", str(wbpath), "--force", "--dry-run", "--items", "1"])
check("dry-run wrote nothing", h1 == wbpath.read_bytes())

# ---------------------------------------------------------------------------
print("\n[4] No keys + no network: must not crash, must still log rows")
pt.direct_vendor_listings = _orig_direct
tmp2 = Path(tempfile.mkdtemp())
wb2 = tmp2 / "Home Wishlist.xlsx"
shutil.copy(SRC, wb2)
env = {k: v for k, v in os.environ.items() if k not in ("SERPAPI_KEY", "EBAY_CLIENT_ID", "EBAY_CLIENT_SECRET")}
proc = subprocess.run([sys.executable, "price_tracker.py", "--workbook", str(wb2), "--force", "--no-direct"],
                      env=env, capture_output=True, text=True, timeout=300)
check("exit code 0 with no keys", proc.returncode == 0, proc.stderr[-300:])
w2 = load_workbook(wb2)
check("4 Run Data rows written with N/A source notes",
      w2[pt.SHEET_RUN].max_row == 5 and "N/A" in str(w2[pt.SHEET_RUN].cell(2, pt.header_map(w2[pt.SHEET_RUN])["sourcenotes"]).value))

# ---------------------------------------------------------------------------
print("\n[5] Failure paths: quota exhaustion, bad keys, blocked sites, page parsing")
pt.SerpApiClient.shopping, pt.SerpApiClient.check_credits, pt.EbayClient.search = _orig
pt.direct_vendor_listings = _orig_direct


class Resp:
    def __init__(self, status=200, data=None, text="", ctype="application/json"):
        self.status_code, self._d, self.text, self.headers = status, data, text, {"content-type": ctype}
        self.ok = status < 400
        self.content = (text or "").encode()

    def json(self):
        return self._d


class Sess:
    """Fake requests.Session that replays queued responses (or raises) and counts calls."""
    def __init__(self, *responses):
        self.q, self.n = list(responses), 0

    def _next(self, *a, **k):
        self.n += 1
        r = self.q.pop(0)
        if isinstance(r, Exception):
            raise r
        return r
    get = post = _next


c = pt.SerpApiClient("k")
c.session = Sess(Resp(200, {"error": "Your account has run out of searches."}))
check("SerpApi quota error -> returns None, client disabled", c.shopping("x") is None and c.disabled)
check("SerpApi disabled -> later calls make NO request", c.shopping("y") is None and c.session.n == 1)
c = pt.SerpApiClient("k"); c.session = Sess(Resp(200, {"error": "Google hasn't returned any results for this query."}))
check("SerpApi 'no results' -> [] and stays enabled", c.shopping("x") == [] and not c.disabled)
c = pt.SerpApiClient("k"); c.session = Sess(Resp(429, {}))
check("SerpApi HTTP 429 -> disabled, no crash", c.shopping("x") is None and c.disabled)
c = pt.SerpApiClient("k"); c.session = Sess(ConnectionError("boom"))
check("SerpApi network exception -> None, NOT disabled", c.shopping("x") is None and not c.disabled)
c = pt.SerpApiClient("k"); c.session = Sess(Resp(200, {"total_searches_left": 2}))
c.check_credits()
check("SerpApi credit reserve (<=3 left) -> disabled before any search", c.disabled)
check("SerpApi with no key -> disabled from the start", pt.SerpApiClient(None).disabled)

e = pt.EbayClient("id", "secret"); e.session = Sess(Resp(401, {}))
check("eBay token failure -> disabled, search returns None", e.search("x") is None and e.disabled)
check("eBay with no keys -> disabled from the start", pt.EbayClient(None, None).disabled)
e = pt.EbayClient("id", "secret"); e.session = Sess(Resp(200, {"access_token": "t"}), Resp(200, {"itemSummaries": [1]}))
check("eBay happy path returns itemSummaries", e.search("x") == [1])
ls = pt.parse_ebay([E("Sonos Ray", 99.5, "Certified - Refurbished", 99.3, 1234)])
check("eBay parse: refurbished + seller comment", ls[0].condition == "refurbished" and "99.3% positive" in ls[0].seller_comment
      and "1,234" in ls[0].seller_comment and ls[0].seller_ok)

html = """<html><script type="application/ld+json">{"@context":"https://schema.org","@graph":[
  {"@type":"Organization","name":"X"},
  {"@type":"Product","name":"THIRDREALITY Smart Plug Gen3","offers":[
     {"@type":"Offer","price":"14.99","priceCurrency":"USD","availability":"https://schema.org/OutOfStock"}]}]}</script></html>"""
info = pt.parse_jsonld_product(html)
check("JSON-LD (@graph): price + OutOfStock detected", info and info["price"] == 14.99 and info["in_stock"] is False)
info = pt.parse_jsonld_product('<script type="application/ld+json">{"@type":"Product","name":"A","offers":{"price":99,"priceCurrency":"USD","availability":"InStock"}}</script>')
check("JSON-LD: InStock single offer", info and info["in_stock"] is True and info["price"] == 99.0)
check("JSON-LD: garbage HTML -> None", pt.parse_jsonld_product("<html>no data</html>") is None)


# ---------------------------------------------------------------------------
print("\n[6] Direct-site discovery: Shopify variants, sitemaps, search fallbacks, Product URLs, .env")


class RSess:
    """Fake session that answers by URL substring (first match wins) and logs every request."""
    def __init__(self, routes):
        self.routes, self.log = routes, []

    def _req(self, url, **kw):
        self.log.append(url)
        for pat, resp in self.routes:
            if pat in url:
                if isinstance(resp, Exception):
                    raise resp
                return resp
        return Resp(404, None, "not found", "text/html")
    get = post = _req


def page(name, price, avail="InStock"):
    return Resp(200, None, '<script type="application/ld+json">{"@type":"Product","name":"%s","offers":{"price":"%s",'
                '"priceCurrency":"USD","availability":"https://schema.org/%s"}}</script>' % (name, price, avail), "text/html")


# --- Shopify: only the matching product is expanded, and every pack variant becomes a listing
plug_vendor = pt.Vendor("THIRDREALITY Store", "thirdreality", True, "thirdreality.com")
shop = RSess([
    ("/search/suggest.json", Resp(200, {"resources": {"results": {"products": [
        {"title": "THIRDREALITY Smart Plug Gen3", "handle": "smart-plug-gen3", "url": "/products/smart-plug-gen3?_pos=1", "price": "14.99"},
        {"title": "THIRDREALITY Smart Bulb ZL1", "handle": "smart-bulb-zl1", "price": "17.99"},
        {"title": "THIRDREALITY Smart Plug E2 Gen2", "handle": "smart-plug-e2", "price": "11.99"},
        {"title": "THIRDREALITY Motion Sensor", "handle": "motion", "price": "13.99"}]}}})),
    ("/products/smart-plug-gen3.js", Resp(200, {"variants": [
        {"id": 1, "title": "1 Pack", "price": 1499, "available": False},
        {"id": 2, "title": "2 Pack", "price": 2799, "available": True},
        {"id": 3, "title": "4 Pack", "price": 4999, "available": True}]})),
])
ls, note = pt.direct_vendor_listings(plug_vendor, plug, shop)
check("Shopify: 4 search results but only the Gen3 plug kept", "4 results, 1 matching" in note, note)
check("Shopify: 3 pack variants priced from cents", sorted(l.price for l in ls) == [14.99, 27.99, 49.99], str([l.price for l in ls]))
check("Shopify: sold-out 1-pack flagged out of stock", any(l.price == 14.99 and l.in_stock is False for l in ls))
check("Shopify: only ONE variants request (unrelated products not fetched)", sum(u.endswith(".js") for u in shop.log) == 1, str(shop.log))
check("Shopify: variant titles carry pack size", sorted(pt.extract_pack_qty(l.title) for l in ls) == [1, 2, 4])

# --- Non-Shopify site: sitemap index -> US product sitemap -> best page -> JSON-LD
pt._SITEMAP_CACHE.clear()
sonos_vendor = pt.Vendor("Sonos Direct", "sonos", True, "www.sonos.com")
sm = RSess([
    ("/search/suggest.json", Resp(404, None, "<html>nope</html>", "text/html")),
    ("robots.txt", Resp(200, None, "User-agent: *\nSitemap: https://www.sonos.com/sitemap-index.xml\n", "text/plain")),
    ("sitemap-index.xml", Resp(200, None, "<sitemapindex><sitemap><loc>https://www.sonos.com/en-gb/sitemap-products.xml</loc></sitemap>"
                                         "<sitemap><loc>https://www.sonos.com/en-us/sitemap-pages.xml</loc></sitemap>"
                                         "<sitemap><loc>https://www.sonos.com/en-us/sitemap-products.xml</loc></sitemap></sitemapindex>", "application/xml")),
    ("en-us/sitemap-products.xml", Resp(200, None, "<urlset><url><loc>https://www.sonos.com/en-us/shop/ray-wall-mount</loc></url>"
                                                   "<url><loc>https://www.sonos.com/en-us/shop/arc</loc></url>"
                                                   "<url><loc>https://www.sonos.com/en-us/shop/ray</loc></url></urlset>", "application/xml")),
    ("en-us/sitemap-pages.xml", Resp(200, None, "<urlset><url><loc>https://www.sonos.com/en-us/blog/ray-review</loc></url></urlset>", "application/xml")),
    ("/en-us/shop/ray", page("Sonos Ray", "179.00")),
])
ls, note = pt.direct_vendor_listings(sonos_vendor, sonos, sm)
check("Sitemap: Sonos Ray found at /en-us/shop/ray via sitemap", len(ls) == 1 and ls[0].url.endswith("/en-us/shop/ray") and ls[0].price == 179.0, note)
check("Sitemap: product sitemap read before pages sitemap; en-gb skipped",
      not any("en-gb" in u for u in sm.log) and sm.log.index("https://www.sonos.com/en-us/sitemap-products.xml") < sm.log.index("https://www.sonos.com/en-us/sitemap-pages.xml"))
check("Sitemap: wall mount, blog and other products never fetched",
      not any(x in u for u in sm.log for x in ("ray-wall-mount", "/shop/arc", "blog/ray")), str(sm.log))
check("Ranking: brand implied by domain ignored ('sonos' not needed in slug)",
      pt.rank_product_urls(["https://www.sonos.com/en-us/shop/ray"], sonos, sonos_vendor) == ["https://www.sonos.com/en-us/shop/ray"])

# --- No sitemap; DuckDuckGo (both) blocked with 202 -> Bing redirect link decoded -> page
import base64 as _b64
pt._SITEMAP_CACHE.clear()
hue_vendor = pt.Vendor("Philips Hue Direct", "philipshue", True, "www.philips-hue.com")
target = "https://www.philips-hue.com/en-us/p/hue-white-and-color-ambiance-slim-downlight-4-inch/046677609404"
bing_link = "https://www.bing.com/ck/a?!&&p=abc&u=a1" + _b64.urlsafe_b64encode(target.encode()).decode().rstrip("=") + "&ntb=1"
se = RSess([
    ("html.duckduckgo.com", Resp(202, None, "anomaly", "text/html")),
    ("lite.duckduckgo.com", Resp(202, None, "anomaly", "text/html")),
    ("bing.com/search", Resp(200, None, f'<li class="b_algo"><h2><a href="{bing_link}">Hue</a></h2></li>', "text/html")),
    ("046677609404", page('Hue White and color ambiance Slim downlight 4 inch', "59.99")),
])
ls, note = pt.direct_vendor_listings(hue_vendor, hue, se)
check("Search fallback: DDG 202 x2 -> Bing link decoded -> priced", len(ls) == 1 and ls[0].price == 59.99 and "Bing" in note, note)

pt._SITEMAP_CACHE.clear()
blocked = RSess([("duckduckgo", Resp(202, None, "", "text/html")), ("bing.com", Resp(429, None, "", "text/html"))])
ls, note = pt.direct_vendor_listings(hue_vendor, hue, blocked)
check("Everything blocked -> N/A note names each engine, no crash",
      ls == [] and "DuckDuckGo blocked (HTTP 202)" in note and "Bing blocked (HTTP 429)" in note, note)

pt._SITEMAP_CACHE.clear()
forbidden = RSess([("robots.txt", Resp(200, None, "Sitemap: https://www.sonos.com/sm.xml", "text/plain")),
                   ("sm.xml", Resp(200, None, "<urlset><url><loc>https://www.sonos.com/en-us/shop/ray</loc></url></urlset>", "application/xml")),
                   ("/shop/ray", Resp(403, None, "denied", "text/html"))])
ls, note = pt.direct_vendor_listings(sonos_vendor, sonos, forbidden)
check("Product page 403 -> explains the site blocks automated access", ls == [] and "blocked automated access" in note, note)

meta = '<meta property="og:title" content="Sonos Ray"><meta property="product:price:amount" content="199.00"><meta property="product:availability" content="out of stock">'
info = pt.parse_meta_price(meta)
check("Meta-tag fallback: price + out of stock", info and info["price"] == 199.0 and info["in_stock"] is False, str(info))

# --- Product URLs column (optional) overrides discovery for that vendor
import copy as _copy
sonos_u = _copy.deepcopy(sonos)
sonos_u.urls = ["https://www.bestbuy.com/site/sonos-ray/6505005.p", "https://www.sonos.com/en-us/shop/ray"]
ctx_stub = type("C", (), {})()
ctx_stub.primary = pt.load_primary_vendors(wb0[pt.SHEET_PRIMARY])
ctx_stub.session = RSess([("bestbuy.com", page("Sonos Ray Soundbar", "169.99")), ("sonos.com/en-us/shop/ray", page("Sonos Ray", "179"))])
ls, status, covered, failed_u = pt.product_url_listings(sonos_u, ctx_stub)
note = str(status)
check("Product URLs: both pages priced", sorted(l.price for l in ls) == [169.99, 179.0], note)
check("Product URLs: marked as from_url (ground truth)", all(l.from_url for l in ls) and not failed_u)
check("Product URLs: Sonos page credited to 'Sonos Direct' vendor", any(l.vendor == "Sonos Direct" for l in ls))
check("Product URLs: Best Buy page matches primary vendor 'Best Buy'",
      pt.keys_match(pt.vendor_key(next(l.vendor for l in ls if "bestbuy" in l.url)), pt.vendor_key("Best Buy")))
check("Product URLs: domains marked covered (discovery skipped)", covered == {"bestbuy.com", "sonos.com"})

# --- .env loading never overrides real environment variables
envf = Path(tempfile.mkdtemp()) / ".env"
envf.write_text('# comment\nSERPAPI_KEY="from_file"\nexport EBAY_CLIENT_ID=file_id\n')
os.environ.pop("SERPAPI_KEY", None); os.environ["EBAY_CLIENT_ID"] = "real_env"
pt.load_dotenv(envf)
check(".env: loads quoted values", os.environ.get("SERPAPI_KEY") == "from_file")
check(".env: does not override existing env (GitHub secrets win)", os.environ.get("EBAY_CLIENT_ID") == "real_env")
check("--allproducts is an alias of --force", pt.parse_args(["--allproducts"]).force is True)

# ---------------------------------------------------------------------------
print("\n[7] Small samples: average always recorded, but excluded from deal rules and trend")
from datetime import datetime as _dt, timedelta as _td
now7 = _dt(2026, 10, 1)
hist = [{"wid": "2", "dt": now7 - _td(days=d), "avg_new": v, "avg_res": None, "n_new": n, "n_res": None}
        for d, v, n in ((3, 13.0, 5), (6, 12.0, 4), (9, 99.0, 1), (12, 14.0, None))]   # None = legacy row
check("Trend skips the 1-listing run ($99) and keeps legacy rows", pt.trend_for(hist, 2, "new", now7) == 13.0,
      str(pt.trend_for(hist, 2, "new", now7)))

pt.direct_vendor_listings = lambda v, it, s: ([
    pt.Listing("THIRDREALITY Smart Plug Gen3 - 1 Pack", "https://thirdreality.com/p?v=1", 14.99, v.name, "direct", in_stock=False),
    pt.Listing("THIRDREALITY Smart Plug Gen3 - 2 Pack", "https://thirdreality.com/p?v=2", 26.99, v.name, "direct", in_stock=True),
    pt.Listing("THIRDREALITY Smart Plug Gen3 - 4 Pack", "https://thirdreality.com/p?v=3", 49.99, v.name, "direct", in_stock=True)],
    "Direct[THIRDREALITY Store]: test")
os.environ.pop("SERPAPI_KEY", None)
tmp7 = Path(tempfile.mkdtemp()) / "Home Wishlist.xlsx"
shutil.copy(SRC, tmp7)
pt.main(["--workbook", str(tmp7), "--items", "2", "--force", "--no-ebay"])
w7 = load_workbook(tmp7)[pt.SHEET_RUN]
h7 = pt.header_map(w7)
row7 = {k: w7.cell(2, c).value for k, c in h7.items()}
check("ThirdReality-only run: Avg Price (New) includes the sold-out vendor 1-pack list price",
      row7["avgpricenew"] == 13.66, str(row7["avgpricenew"]))
check("...with Avg Sample (New) = 3", row7["avgsamplenew"] == 3, str(row7["avgsamplenew"]))
check("...vendor baseline = sold-out 1-pack list price $14.99", row7["vendorbaselinenew"] == 14.99, str(row7["vendorbaselinenew"]))
check("...the vendor's normal 2/4-pack pricing is NOT a deal (compared per pack size)", row7["dealsfound"] == 0)

# ---------------------------------------------------------------------------
print("\n[8] Regressions from run 20261001-020327-5949 (Product URLs, baselines, false deals)")

# --- 8a. page extractors: Dell-style microdata / embedded JSON, sold-out JSON-LD, retries, Best Buy API
dell_html = '<html><title>Philips Hue Slim Downlight | Dell USA</title><div itemscope><span itemprop="price" content="64.99"></span></div></html>'
check("Dell-style microdata price parsed", (pt.parse_microdata_price(dell_html) or {}).get("price") == 64.99)
js_html = '<script>window.__STATE__={"product":{"name":"x","dellPrice":"$64.99","financing":"$6/mo"}}</script>'
check("Embedded-JSON price parsed (JS-rendered page)", (pt.parse_embedded_price(js_html, 60) or {}).get("price") == 64.99)
check("Embedded-JSON skips implausible prices (financing $6)",
      pt.parse_embedded_price('{"salePrice": 6}', 60) is None)
oos_ld = ('<script type="application/ld+json">{"@type":"Product","name":"Hue Slim downlight 4 inch",'
          '"offers":{"price":"69.99","priceCurrency":"USD","availability":"https://schema.org/OutOfStock"}}</script>')
info = pt.parse_jsonld_product(oos_ld)
check("Sold-out Hue page still yields its price", info and info["price"] == 69.99 and info["in_stock"] is False, str(info))

flaky = Sess(ConnectionError("reset"), ConnectionError("reset"), Resp(200, None, dell_html, "text/html"))
r, why = pt.fetch_page("https://www.dell.com/x", flaky)
check("fetch_page retries network errors (BestBuy-style) and recovers", r is not None and flaky.n == 3, why)
dead = Sess(*[ConnectionError("reset")] * (pt.PAGE_RETRIES + 1))
_cffi, pt.cffi_requests = pt.cffi_requests, None
r, why = pt.fetch_page("https://www.bestbuy.com/x", dead)
pt.cffi_requests = _cffi
check("fetch_page gives a clear reason when every attempt fails", r is None and "network error" in why, why)
check("Best Buy SKU parsed from old + new URL styles",
      pt._bestbuy_sku("https://www.bestbuy.com/site/sonos-ray/6505005.p?skuId=6505005") == "6505005"
      and pt._bestbuy_sku("https://www.bestbuy.com/product/sonos-ray/J3ZYG/sku/6505005") == "6505005")
os.environ["BESTBUY_API_KEY"] = "k"
bb = Sess(Resp(200, {"products": [{"sku": 6505005, "name": "Sonos - Ray Soundbar - Black", "salePrice": 219.0,
                                   "onlineAvailability": True}]}))
ls, why = pt.bestbuy_api_listing("https://www.bestbuy.com/site/sonos-ray/6505005.p", "Best Buy", bb)
os.environ.pop("BESTBUY_API_KEY")
check("Best Buy API path prices the page without scraping", ls and ls[0].price == 219.0 and ls[0].in_stock, why)

# --- 8b. matching: Beam is not Ray; accessories 'for Sonos Ray' are not the Ray
check("eBay 'Sonos Beam' does NOT match 'Sonos Ray' -> Low",
      cc(sonos, "Sonos Beam (Gen 2) Smart Soundbar Black")[0] == "Low")
check("'Wall Bracket compatible with Sonos Ray' -> Low",
      cc(sonos, "Mounting Kit Compatible with Sonos Ray Soundbar")[0] == "Low")
check("Genuine 'Sonos Ray Soundbar' still matches", cc(sonos, "Sonos Ray Compact Soundbar Black RAYG1US1BLK")[0] != "Low")

# --- 8c. end-to-end replay of the bad run with fake sources
FAKE_SERP["sonos"] = [
    S("Sonos Ray Soundbar RAYG1US1BLK", 287.42, "Zaytoun"),                       # overpriced reseller
    S("Sonos Ray Soundbar + Sub Mini Bundle", 699.00, "Amazon.com"),              # bundle: skews the avg
    S("Sonos Ray Soundbar Black", 649.00, "Walmart - Seller"),                    # gouged marketplace
    S("Sonos Ray Soundbar Black RAYG1US1BLK", 219.00, "Best Buy"),                # fills the failed BB URL
    S("Sonos Ray Soundbar Black RAYG1US1BLK", 219.00, "Sonos"),                   # duplicate of sonos.com page
    S("Sonos Ray Compact Soundbar RAYG1US1BLK", 179.00, "Target"),               # a real ~18% deal
]
FAKE_EBAY["sonos"] = [E("Sonos Beam (Gen 2) Soundbar Black", 199.99, "Used", 99.5, 900)]
pt.SerpApiClient.shopping = fake_shopping
pt.SerpApiClient.check_credits = lambda self: None
pt.EbayClient.search = fake_ebay
pt.direct_vendor_listings = lambda v, it, s: ([], f"Direct[{v.name}]: N/A (fake)")
os.environ.update({"SERPAPI_KEY": "dummy", "EBAY_CLIENT_ID": "dummy", "EBAY_CLIENT_SECRET": "dummy"})

hue_ld = oos_ld
routes = {"philips-hue.com": Resp(200, None, hue_ld, "text/html"),
          "dell.com": Resp(200, None, dell_html, "text/html"),
          "sonos.com": page("Sonos Ray", "219.00")}


class URLSess:
    def get(self, url, **kw):
        if "bestbuy.com" in url:
            raise ConnectionError("reset by peer")
        for k, v in routes.items():
            if k in url:
                return v
        return Resp(404, None, "", "text/html")
    post = get


tmp8 = Path(tempfile.mkdtemp()) / "Home Wishlist.xlsx"
shutil.copy(SRC, tmp8)
w = load_workbook(tmp8)
ms = w[pt.SHEET_MASTER]
mh = pt.ensure_columns(ms, ["Product URLs"])
for r in range(2, ms.max_row + 1):
    wid = str(ms.cell(r, mh["wishlistitem"]).value)
    ms.cell(r, mh["producturls"], {
        "1": "https://www.philips-hue.com/en-us/p/hue-slim-downlight-4/046677\nhttps://www.dell.com/en-us/shop/hue/apd/ab1",
        "3": "https://www.sonos.com/en-us/shop/ray; https://www.bestbuy.com/site/sonos-ray/6505005.p"}.get(wid))
w.save(tmp8)
_orig_session = pt.requests.Session
pt.requests.Session = URLSess
_cffi, pt.cffi_requests = pt.cffi_requests, None
pt.main(["--workbook", str(tmp8), "--items", "1,3", "--force"])
pt.requests.Session, pt.cffi_requests = _orig_session, _cffi

w8 = load_workbook(tmp8)
rw, dw = w8[pt.SHEET_RUN], w8[pt.SHEET_DEALS]
rh8, dh8 = pt.header_map(rw), pt.header_map(dw)
rr = {str(rw.cell(r, rh8["wishlistitem"]).value): {k: rw.cell(r, c).value for k, c in rh8.items()} for r in range(2, rw.max_row + 1)}
dd = [{k: dw.cell(r, c).value for k, c in dh8.items()} for r in range(2, dw.max_row + 1)]
h1, s3 = rr.get("1", {}), rr.get("3", {})
print("     Hue :", h1.get("avgpricenew"), h1.get("vendorbaselinenew"), "|", h1.get("producturlstatus"))
print("     Sonos:", s3.get("avgpricenew"), s3.get("vendorbaselinenew"), "|", s3.get("producturlstatus"))
for d in dd:
    print(f"     deal: item {d['wishlistitem']} ${d['price']} {d['vendor']} | {d['dealrule']} | "
          f"%tgt={d['belowtarget']} %base={d['belowbaseline']}")
check("Hue: Avg (New) recorded even though the Hue page is sold out", h1.get("avgpricenew") is not None, str(h1))
check("Hue: vendor baseline from both Product URLs (69.99 sold-out + 64.99 Dell)",
      h1.get("vendorbaselinenew") == 67.49, str(h1.get("vendorbaselinenew")))
check("Hue: Product URL Status shows Dell priced + Hue OUT OF STOCK",
      "dell.com: $64.99" in (h1.get("producturlstatus") or "") and "OUT OF STOCK" in (h1.get("producturlstatus") or ""))
check("Sonos: vendor baseline = sonos.com $219", s3.get("vendorbaselinenew") == 219.0, str(s3.get("vendorbaselinenew")))
check("Sonos: Best Buy URL network error filled from Best Buy's Google Shopping row",
      "bestbuy.com: $219.00 via Google Shopping fallback" in (s3.get("producturlstatus") or ""), s3.get("producturlstatus"))
check("Sonos: Avg (New) anchored near $219 (not ~$454)", s3.get("avgpricenew") and 170 <= s3["avgpricenew"] <= 260,
      str(s3.get("avgpricenew")))
sonos_deals = [d for d in dd if str(d["wishlistitem"]) == "3"]
check("Sonos: Zaytoun $287.42 is NOT a deal", all(d["price"] != 287.42 for d in sonos_deals))
check("Sonos: eBay 'Beam' $199.99 is NOT a deal", all(d["price"] != 199.99 for d in sonos_deals))
check("Sonos: Target $179 IS a deal vs the $219 baseline, with correct % columns",
      any(d["price"] == 179.0 and d["baselineprice"] == 219.0 and abs(d["belowbaseline"] - 0.1826) < 0.001
          and abs(d["belowtarget"] - (175 - 179) / 175) < 0.001 for d in sonos_deals), str(sonos_deals))
check("Sonos: sonos.com's Google row not double-counted alongside the scraped page",
      "dropped - same store already priced" in (s3.get("sourcenotes") or ""))

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
