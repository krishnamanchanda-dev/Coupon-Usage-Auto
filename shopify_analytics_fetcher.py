"""
Fetches session metrics from Shopify Analytics (ShopifyQL) for a given date.
Returns total human sessions and sessions_with_cart_additions.
Used for Product Deep Dive cols E and G.
"""

import requests
from config import SHOP_DOMAIN, SHOPIFY_API_VERSION


def fetch_session_metrics(date_str: str, token: str) -> dict:
    """
    Returns dict with 'sessions' and 'cart_additions' for the given date (YYYY-MM-DD).
    Uses ShopifyQL via the GraphQL Admin API.
    """
    url = f"https://{SHOP_DOMAIN}/admin/api/{SHOPIFY_API_VERSION}/graphql.json"
    query = """
    query analyticsQuery($query: String!) {
      analyticsQuery(query: $query) {
        data
        parseErrors {
          message
          range { start { line column } end { line column } }
        }
      }
    }
    """
    shopify_ql = (
        f"FROM sessions "
        f"SHOW sessions, sessions_with_cart_additions "
        f"WHERE human_or_bot_session = 'human' "
        f"SINCE {date_str} UNTIL {date_str} "
        f"WITH TOTALS"
    )

    resp = requests.post(
        url,
        headers={
            "X-Shopify-Access-Token": token,
            "Content-Type": "application/json",
        },
        json={"query": query, "variables": {"query": shopify_ql}},
        timeout=30,
    )
    resp.raise_for_status()

    body = resp.json()
    aq = body.get("data", {}).get("analyticsQuery", {})

    errors = aq.get("parseErrors", [])
    if errors:
        raise RuntimeError(f"ShopifyQL parse error: {errors}")

    # `data` is a JSON string from Shopify — parse it
    import json as _json
    raw_data = aq.get("data", "")
    print(f"  DEBUG sessions raw_data type={type(raw_data).__name__} value={str(raw_data)[:300]}")
    if isinstance(raw_data, str):
        data = _json.loads(raw_data) if raw_data else {}
    else:
        data = raw_data or {}

    # Structure: {"columns": [...], "rowData": [["101319"]], "totals": {...}}
    # With a simple aggregate query (no TIMESERIES), rowData has one row.
    # columns order matches SHOW clause: sessions
    columns = [c.get("name", "").lower() for c in data.get("columns", [])]
    row_data = data.get("rowData", [])

    if not row_data and not data.get("totals"):
        print(f"  WARNING: No Shopify session data for {date_str}")
        return {"sessions": 0, "cart_additions": 0}

    def _get_metric(name: str) -> int:
        # Try totals first
        totals = data.get("totals", {})
        for key in (name, name.lower(), name.upper()):
            if key in totals:
                return int(totals[key])
        # Fall back to summing rowData
        idx = next((i for i, c in enumerate(columns) if c == name.lower()), None)
        if idx is None:
            return 0
        total = 0
        for row in row_data:
            try:
                total += int(row[idx])
            except (TypeError, ValueError, IndexError):
                pass
        return total

    sessions       = _get_metric("sessions")
    cart_additions = _get_metric("sessions_with_cart_additions")

    print(f"  Shopify sessions for {date_str}: total={sessions}, cart_additions={cart_additions}")
    return {"sessions": sessions, "cart_additions": cart_additions}


# Keep a convenience wrapper for backward compat
def fetch_total_sessions(date_str: str, token: str) -> int:
    return fetch_session_metrics(date_str, token)["sessions"]


def fetch_returning_customer_rate(date_str: str, token: str) -> float:
    """
    Returns returning_customer_rate as a float (e.g. 0.3554 for 35.54%)
    from Shopify ShopifyQL for the given date (YYYY-MM-DD).
    """
    url = f"https://{SHOP_DOMAIN}/admin/api/{SHOPIFY_API_VERSION}/graphql.json"
    gql = """
    query analyticsQuery($query: String!) {
      analyticsQuery(query: $query) {
        data
        parseErrors {
          message
          range { start { line column } end { line column } }
        }
      }
    }
    """
    shopify_ql = (
        f"FROM sales "
        f"SHOW returning_customer_rate "
        f"SINCE {date_str} UNTIL {date_str} "
        f"WITH TOTALS"
    )

    resp = requests.post(
        url,
        headers={
            "X-Shopify-Access-Token": token,
            "Content-Type": "application/json",
        },
        json={"query": gql, "variables": {"query": shopify_ql}},
        timeout=30,
    )
    resp.raise_for_status()

    body = resp.json()
    aq = body.get("data", {}).get("analyticsQuery", {})
    errors = aq.get("parseErrors", [])
    if errors:
        raise RuntimeError(f"ShopifyQL parse error: {errors}")

    import json as _json
    raw_data = aq.get("data", "")
    print(f"  DEBUG returning_rate raw_data type={type(raw_data).__name__} value={str(raw_data)[:300]}")
    data = _json.loads(raw_data) if isinstance(raw_data, str) and raw_data else (raw_data or {})

    # Try totals first
    totals = data.get("totals", {})
    for key in ("returning_customer_rate", "returningCustomerRate"):
        if key in totals:
            val = float(totals[key])
            print(f"  Shopify returning_customer_rate for {date_str}: {val:.4f}")
            return val

    # Fall back to rowData
    columns = [c.get("name", "").lower() for c in data.get("columns", [])]
    row_data = data.get("rowData", [])
    idx = next((i for i, c in enumerate(columns) if "returning" in c and "rate" in c), None)
    if idx is not None and row_data:
        val = float(row_data[-1][idx])  # last row = totals row with TIMESERIES + TOTALS
        print(f"  Shopify returning_customer_rate for {date_str}: {val:.4f}")
        return val

    print(f"  WARNING: returning_customer_rate not found for {date_str}, defaulting to 0")
    return 0.0
