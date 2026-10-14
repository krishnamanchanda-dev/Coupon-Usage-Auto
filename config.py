import os
from datetime import datetime, timedelta, timezone

# Store operates in IST (UTC+5:30); use IST midnight as day boundaries
_IST = timezone(timedelta(hours=5, minutes=30))


def _ist_midnight_utc(d) -> str:
    """Return UTC ISO string for midnight IST on date d."""
    midnight_ist = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=_IST)
    return midnight_ist.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _yesterday_iso() -> str:
    """Returns start of yesterday (IST midnight) expressed in UTC."""
    yesterday = datetime.now(_IST).date() - timedelta(days=1)
    return _ist_midnight_utc(yesterday)


def _today_iso() -> str:
    """Returns start of today (IST midnight) expressed in UTC — end of yesterday's window."""
    today = datetime.now(_IST).date()
    return _ist_midnight_utc(today)


# Token generation
TOKEN_API_URL = "https://backgroundprocessor.swiss-custom.site/api/public/token/generate"
TOKEN_API_KEY = os.getenv("TOKEN_API_KEY", "")
TOKEN_SCOPES = ["read_all_orders", "read_products", "read_customers", "read_inventory", "read_orders"]

# Shopify
SHOP_DOMAIN = os.getenv("SHOP_DOMAIN", "swiss-beauty-dev.myshopify.com")
SHOPIFY_API_VERSION = "2024-10"
ORDERS_PAGE_SIZE = 250  # max allowed by Shopify

# Google Sheets (fill in when API is shared)
SHEETS_API_URL = os.getenv("SHEETS_API_URL", "")
SHEETS_API_KEY = os.getenv("SHEETS_API_KEY", "")
SHEET_ID = os.getenv("SHEET_ID", "")

# Date range for order fetching.
# Default: yesterday's full day (for daily scheduled runs).
# Override via env vars for backfill: e.g. ORDERS_CREATED_AT_MIN=2024-10-11T00:00:00Z
# Default: yesterday's full day for daily runs.
# For backfill from Oct 11 onward, set ORDERS_CREATED_AT_MIN=2026-10-11T00:00:00Z
ORDERS_CREATED_AT_MIN = os.getenv("ORDERS_CREATED_AT_MIN", _yesterday_iso())
ORDERS_CREATED_AT_MAX = os.getenv("ORDERS_CREATED_AT_MAX", _today_iso())
