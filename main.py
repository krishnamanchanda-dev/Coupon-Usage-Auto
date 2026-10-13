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
from test_tab_writer import (
    write_test_tab, print_categorization, _get_client, SPREADSHEET_ID,
    update_overall_customer_rates,
)
from ga4_fetcher import fetch_marketing_sessions
from shopify_analytics_fetcher import fetch_session_metrics, fetch_returning_customer_rate
from product_deep_dive_writer import update_product_deep_dive


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

    print(f"\n[4/5] Writing to 'Coupon Usage' tab and updating Overall tab...")
    write_test_tab(flat_rows, dict(total_orders_by_day))

    print(f"\n[5/5] Updating Product Deep Dive tab...")
    gc = _get_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    for date_str in total_orders_by_day:
        try:
            mktng_sessions = fetch_marketing_sessions(date_str)
            shopify = fetch_session_metrics(date_str, token)
            update_product_deep_dive(
                sh, date_str,
                mktng_sessions,
                shopify["sessions"],
                shopify["cart_additions"],
            )
        except Exception as e:
            print(f"  WARNING: Product Deep Dive update failed for {date_str}: {e}")

        try:
            ret_rate = fetch_returning_customer_rate(date_str, token)
            update_overall_customer_rates(sh, date_str, ret_rate)
        except Exception as e:
            print(f"  WARNING: Customer rate update failed for {date_str}: {e}")

    print("\nDone.")


if __name__ == "__main__":
    main()
