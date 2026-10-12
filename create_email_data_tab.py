"""One-shot script to populate Email_Data tab using today's already-fetched data."""
import os
from config import ORDERS_CREATED_AT_MIN, ORDERS_CREATED_AT_MAX
from token_manager import get_shopify_token
from order_fetcher import iter_all_orders
from aggregator import aggregate_coupons_by_day
from collections import defaultdict
from test_tab_writer import aggregate_for_tab, write_email_data_tab, _get_client, SPREADSHEET_ID

token = get_shopify_token()
print("Token obtained.")

total_orders_by_day = defaultdict(int)
discounted_orders = []
for order in iter_all_orders(token):
    date = (order.get("created_at") or "")[:10]
    if date:
        total_orders_by_day[date] += 1
    codes = order.get("discount_codes", [])
    total_disc = float(order.get("total_discounts", "0") or "0")
    if codes or total_disc > 0:
        discounted_orders.append(order)

print(f"Fetched {sum(total_orders_by_day.values())} orders.")

flat_rows = aggregate_coupons_by_day(iter(discounted_orders))
aggregated = aggregate_for_tab(flat_rows, dict(total_orders_by_day))

gc = _get_client()
sh = gc.open_by_key(SPREADSHEET_ID)
write_email_data_tab(sh, aggregated, flat_rows)
print("Done — Email_Data tab created.")
