# HomeImprovement
Home automation and improvement wishlist as well as stored automations for at home.

## Price tracker

`price_tracker.py` reads **Home Wishlist.xlsx** (Master Sheet + vendor tabs), prices every wishlist item and
appends results to the **Run Data** and **Deals Data** tabs. Nothing else in the workbook is changed.
`dashboard.html` visualises those two tabs (open it in a browser and load the workbook, or connect it to this repo).

```
pip install -r requirements.txt
python price_tracker.py --dry-run --force      # everything, write nothing
python price_tracker.py --items 1,3 --force    # specific WishlistItem ids
python test_offline.py                         # offline test-suite (no network, no keys)
```

### How a price becomes trusted

| Layer | Sources | Evidence tier |
|---|---|---|
| Verification | Master Sheet **Product URLs** (validated, never blindly trusted), Best Buy Products API, eBay Browse API | `verified_direct` |
| Verification | Merchant pages the tracker found itself: Shopify stores (auto-detected), retailer site search / sitemaps, DTC vendor sites, merchant links found in Google Shopping | `verified_discovered` |
| Discovery / corroboration | Google Shopping via SerpApi | `market_snapshot` |

* **Identifier-first matching**: GTIN/UPC/EAN > MPN/SKU/model > page metadata > title/specs. Identifiers are
  extracted automatically (JSON-LD, microdata, Shopify barcodes, Best Buy UPC/model, GTINs in URLs) and remembered
  in `tracker_state/`. Optional Master Sheet columns `GTIN`/`UPC`, `MPN`/`Model`, `Brand` are used when present.
* **Hard conflicts** make a listing Low no matter how good its title looks: different GTIN, near-variant model code
  (e.g. `RAYG1US1BLK` vs `RAYG1EU1BLK`), region/voltage, generation, colour, size, accessory, bundle, configuration
  (Pro/Mini/...), sibling model of the same brand.
* **Product URLs** are re-validated each run; stale links redirected to a category/other product are reported as
  `identity_mismatch`.
* Only **verified + High + new + ordinary-priced** listings feed the verified baseline, the 30-day verified median,
  the verified low and the offer history. A Google Shopping row standing in for a blocked Product URL stays a
  `market_snapshot`. Medium listings count only after independent corroboration promotes them.
* **Price fields**: current, regular/compare-at/MSRP, shipping, effective (pre-tax) price, condition,
  availability, pack size. Coupon / membership / subscription / financing / trade-in prices are flagged as
  conditional and never mixed into ordinary averages.
* **Reference price** (what "normal" costs, used for deal rule A): official MSRP/regular/compare-at → regular-price
  consensus → verified history. Current market statistics are computed separately.
* **Quantity Needed**: the cheapest valid combination of singles and allowed multipacks (Bulk Keywords) is recorded
  as *Best Qty Plan*; one-off resale listings are used at most once.

### Master Sheet syntax

`;` separates the entries of **Product specifications**, **Search Keywords**, **Product URLs** and **Bulk Keywords**
(a line break inside the cell works too). **Notes** is free text and is never parsed.

* **`!` exclusions** - an entry that starts with `!` is *excluded*: `W50; !Lite` means "must be a W50, never a Lite".
  The phrase runs up to the next `;`. It is never used as a search word (Google queries get `-Lite`), a listing or
  Product URL page whose title / URL / model code contains it is rejected (whole words only: `Lite` ≠ `Elite`).
  Works in both Product specifications and Search Keywords.
* **Product specifications** are *required words* (every entry must appear in a listing title for High confidence).
  Put only attributes a title will actually contain (`Black`, `50 inch`); a bare `50 in` is read as `50 inch`, and
  synonyms (`50 in; 50 inch; 50"`) collapse into one. Setup context ("Unmounted") belongs in Notes.
* **Only Check Primary Links** (`Yes` / `No`) - `Yes` prices the item from its **Product URLs**. The Primary and
  Secondary vendor lists, retailer discovery, Best Buy keyword search and eBay are skipped, and the vendors of the
  item's own links count as its primary vendors (so a deal at e.g. Article or Etsy is reported; only the listing's own
  condition makes it resale). How much SerpApi searching follows depends on how many **vendors** have a working link
  (`LINKS_ENOUGH_VENDORS = 3`):
  * **3 or more vendors work** - no search of any kind.
  * **2 or fewer work** - a general Google Shopping search runs *and* every vendor whose link(s) all failed gets a
    same-vendor search: that vendor's Google row is identity-checked, its merchant page is fetched and priced there
    (`verified_discovered`), or - when the page can't be fetched - the row stays a labelled `market_snapshot`.
    A vendor with a working link is never searched (and its rows in the general results are ignored). A replacement
    link that worked is re-tried first next time (0 credits). *Product URL Status* / *Source Notes* name the replacement
    so you can update the sheet.
  * `Yes` with **no Product URLs** - normal search.
* **Run Data flags** - three columns record what happened each run: **Search Mode** (normal / links only, no search /
  links + expanded search / the "Yes but no URLs" fallback), **Primary Link Failed** (`No`, `No links listed`, or
  `Yes: host [outcome]`) and **Expanded Search No Results** (`No`, `Not run (reason)`, or `Yes: ...` naming the general
  and/or same-vendor searches that found nothing). A `Yes` carries `(run N in a row)` - the tracker reads the previous
  Run Data rows back to count consecutive flagged runs.

### Reliability
Retries with exponential backoff + jitter (Retry-After honoured), one request at a time per domain with a minimum
interval, a circuit breaker after repeated blocks, conditional GET (ETag/Last-Modified) with cached parsed results,
strategy memory (plain HTTP vs `curl_cffi` Chrome impersonation), Playwright **only** when a page's product data is
JavaScript-only, and full isolation: one failing source/parser/item never stops the others. Every source reports a
category (`success`, `no_match`, `identity_mismatch`, `blocked`, `timeout_network`, `parser_failure`,
`unavailable`, `api_unavailable`, `skipped`) in **Retrieval Outcomes** / **Product URL Status**.

### History, caching and idempotency
`tracker_state/` (commit it) holds learned identifiers, validated retailer URLs (re-verified each run, rediscovered
only after 30 days; "not carried" results re-checked after 10), Shopify detection, SerpApi responses (20 h – a re-run
spends no credits), offer states and `observations.jsonl`. Unchanged offers are recorded as *unchanged since …*, not
as new price events. Re-using a run id (`--run-id`, or `PRICE_TRACKER_RUN_ID` – the workflow uses the GitHub run id)
replaces that run's rows instead of duplicating them. URLs are stored without tracking parameters.

### Keys (all optional, all free tiers)
`BESTBUY_API_KEY` (developer.bestbuy.com), `SERPAPI_KEY` (Google Shopping discovery – skipped automatically when
verified pages already price an item and discovery ran within 6 days), `EBAY_CLIENT_ID` / `EBAY_CLIENT_SECRET`
(eBay stays optional; without it the tracker runs normally). Put them in `.env` locally or in repository secrets.

### Code layout
`price_tracker.py` (settings, workbook I/O, orchestration) and `pricetrack/`: `identity.py`, `extract.py`,
`fetch.py`, `adapters/` (page, shopify, bestbuy, discovery, serpapi, ebay), `reconcile.py`, `pricing.py`,
`history.py`, `urls.py`, `text.py`, `models.py`.

The GitHub Actions workflow is stored in the root file `workflows`; GitHub only runs it from
`.github/workflows/price-tracker.yml`.
