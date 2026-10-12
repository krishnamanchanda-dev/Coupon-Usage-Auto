import requests
import re
from typing import Generator, Optional
from config import SHOP_DOMAIN, SHOPIFY_API_VERSION, ORDERS_PAGE_SIZE, ORDERS_CREATED_AT_MIN, ORDERS_CREATED_AT_MAX


def _build_initial_params() -> dict:
    params = {
        "limit": ORDERS_PAGE_SIZE,
        "status": "any",
        "fields": "id,created_at,discount_codes,total_discounts",
    }
    if ORDERS_CREATED_AT_MIN:
        params["created_at_min"] = ORDERS_CREATED_AT_MIN
    if ORDERS_CREATED_AT_MAX:
        params["created_at_max"] = ORDERS_CREATED_AT_MAX
    return params


def _extract_next_page_url(link_header: Optional[str]) -> Optional[str]:
    """Parse Shopify's Link header to get the next page URL."""
    if not link_header:
        return None
    # Format: <url>; rel="next", <url>; rel="previous"
    match = re.search(r'<([^>]+)>;\s*rel="next"', link_header)
    return match.group(1) if match else None


def iter_all_orders(token: str) -> Generator[dict, None, None]:
    """Yields ALL orders (for computing total order count per day)."""
    session = requests.Session()
    session.headers.update({"X-Shopify-Access-Token": token})

    base_url = f"https://{SHOP_DOMAIN}/admin/api/{SHOPIFY_API_VERSION}/orders.json"
    url = base_url
    params = _build_initial_params()
    page = 0

    while url:
        response = session.get(url, params=params, timeout=60)
        response.raise_for_status()

        orders = response.json().get("orders", [])
        page += 1

        yield from orders

        params = None
        url = _extract_next_page_url(response.headers.get("Link"))

        if page % 10 == 0:
            print(f"  Fetched page {page} ({page * ORDERS_PAGE_SIZE:,} orders so far)...")


def iter_orders(token: str) -> Generator[dict, None, None]:
    """Yields only orders that have a discount applied."""
    for order in iter_all_orders(token):
        discount_codes = order.get("discount_codes", [])
        total_discounts = float(order.get("total_discounts", "0") or "0")
        if discount_codes or total_discounts > 0:
            yield order
