"""
pricetrack - the retrieval / matching / history engine behind price_tracker.py.

Module map (each one is importable and testable on its own):
  models     - Item, Vendor, Listing, SourceResult, evidence tiers and retrieval-outcome categories
  urls       - URL normalisation (tracking params stripped, functional ids kept), merchant item ids
  identity   - identifier-first product matching (GTIN > MPN/SKU > page metadata > title/specs),
               variant-conflict detection, canonical product fingerprints
  extract    - structured-data extraction (JSON-LD > microdata > OpenGraph > hydration JSON > DOM)
  fetch      - resilient HTTP: retries w/ backoff+jitter, per-domain rate limits, circuit breaker,
               conditional GET cache, curl_cffi + Playwright fallbacks, outcome classification
  adapters   - SourceAdapter implementations (product page, Shopify, retailer
               sitemap/site-search, browser, SerpApi, eBay)
  reconcile  - stable offer ids, cross-source corroboration, promotion of uncertain listings
  pricing    - effective price, conditional pricing, reference-price hierarchy, market stats,
               quantity-aware pack optimisation, deal rules
  learning   - what a run teaches the next one: retailer yield, suggested vendors, Master Sheet suggestions,
               identifiers learned from verified pages
  history    - persistent state (discovery cache, learned identities, offer state, observations),
               verified-only historical statistics, price-change events, idempotent re-runs

Design notes borrowed (concepts only) from established open-source projects:
  * changedetection.io / urlwatch - per-target snapshots, "only record a change when it changed",
    conditional requests, isolated per-watch failures
  * Scrapy - per-domain concurrency/delay (AutoThrottle idea), retry middleware with backoff,
    HTTP cache middleware (ETag / Last-Modified), item pipelines (normalise -> dedupe -> store)
  * Playwright - render only when static HTML lacks the data; block heavy resources
  * Apprise - pluggable adapter registry where a missing/failed plugin never breaks the others
"""

__all__ = ["models", "urls", "identity", "extract", "fetch", "adapters", "reconcile", "pricing", "history", "learning"]
