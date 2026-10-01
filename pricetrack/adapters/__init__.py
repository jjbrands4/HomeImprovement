"""Source adapters: each turns a merchant/API/aggregator into normalised Listings + an Outcome."""
from .base import AdapterContext, SourceAdapter, page_listings, run_safely  # noqa: F401
from .bestbuy import BestBuyAdapter, bestbuy_sku  # noqa: F401
from .discovery import KNOWN_DOMAINS, RetailerDiscoveryAdapter, rank_product_urls  # noqa: F401
from .ebay import EbayAdapter, EbayClient, parse_ebay  # noqa: F401
from .page import ProductPageAdapter  # noqa: F401
from .serpapi import SerpApiAdapter, SerpApiClient, build_query, conditional_kind, parse_serpapi  # noqa: F401
from .shopify import ShopifyAdapter  # noqa: F401
