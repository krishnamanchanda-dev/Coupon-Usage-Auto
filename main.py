#!/usr/bin/env python3
"""
Coupon Automation: fetches Shopify orders, aggregates coupon usage by day,
and writes results into the 'test' Google Sheet tab.

Daily run:  fetches yesterday's data (default).
Backfill:   set ORDERS_CREATED_AT_MIN / ORDERS_CREATED_AT_MAX env vars.
"""

import sys
from collections import defaultdict
from config import ORDERS_CREATED_AT_MIN, ORDERS_CREATED_AT_MAX
from token_manager import get_shopify_token
from order_fetcher import iter_all_orders
from aggregator import aggregate_coupons_by_day
from test_tab_writer import write_test_tab, print_categorization


def main():
    print(f"Coupon Automation starting...")
    print(f"  Date range: {ORDERS_CREATED_AT_MIN} → {ORDERS_CREATED_AT_MAX}")

    print("\n[1/4] Generating Shopify token...")
    token = get_shopify_token()
    print(f"  Token obtained.")

    print("\n[2/4] Fetching all orders...")
    total_orders_by_day: dict[str, int] = defaultdict(int)
    discounted_orders = []

    for order in iter_all_orders(token):
        date = (order.get("created_at") or "")[:10]
        if date:
            total_orders_by_day[date] += 1
        discount_codes = order.get("discount_codes", [])
        total_disc = float(order.get("total_discounts", "0") or "0")
        if discount_codes or total_disc > 0:
            discounted_orders.append(order)

    print(f"  Orders by day: { {d: c for d, c in sorted(total_orders_by_day.items())} }")

    print("\n[3/4] Aggregating coupon data by day...")
    flat_rows = aggregate_coupons_by_day(iter(discounted_orders))

    if not flat_rows:
        print("  No discounted orders found. Nothing to write.")
        sys.exit(0)

    print_categorization(flat_rows)

    print(f"\n[4/4] Writing to 'test' tab in Google Sheet...")
    write_test_tab(flat_rows, dict(total_orders_by_day))

    print("\nDone.")


if __name__ == "__main__":
    main()
