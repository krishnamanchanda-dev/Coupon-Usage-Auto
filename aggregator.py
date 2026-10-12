from collections import defaultdict
from typing import Iterable


def _normalize(code: str) -> str:
    """Canonical key: lowercase, no spaces. 'buy 3' == 'buy3', 'free lipstick' == 'freelipstick'."""
    return code.lower().replace(" ", "").strip()


def aggregate_coupons_by_day(orders: Iterable[dict]) -> list[dict]:
    """
    Aggregates coupon usage per day across all orders.

    Returns rows sorted by date asc, then total_discount_amount desc:
      [{"date": "YYYY-MM-DD", "coupon_code": str, "times_used": int, "total_discount_amount": float}, ...]

    Edge cases handled:
    - Orders with multiple discount codes: each code counted separately per day
    - Manual/automatic discounts with no named code: bucketed under [NO CODE / MANUAL]
    - Codes with zero amount: included in count
    """
    # stats[date][code] = {"times_used": int, "total_discount_amount": float}
    stats: dict[str, dict[str, dict]] = defaultdict(lambda: defaultdict(lambda: {"times_used": 0, "total_discount_amount": 0.0}))

    # canonical_display[date][norm_key] = preferred display name (first seen)
    canonical_display: dict[str, dict[str, str]] = defaultdict(dict)

    total_orders_processed = 0
    for order in orders:
        total_orders_processed += 1

        created_at = order.get("created_at", "")
        date = created_at[:10] if created_at else "unknown"

        discount_codes = order.get("discount_codes", [])

        if discount_codes:
            for dc in discount_codes:
                raw_code = dc.get("code", "").strip()
                if not raw_code:
                    continue
                norm = _normalize(raw_code)
                # Keep the first non-spaced variant as display name, else uppercase original
                if norm not in canonical_display[date]:
                    canonical_display[date][norm] = raw_code.replace(" ", "").upper()
                amount = float(dc.get("amount", "0") or "0")
                stats[date][norm]["times_used"] += 1
                stats[date][norm]["total_discount_amount"] += amount
        else:
            amount = float(order.get("total_discounts", "0") or "0")
            if amount > 0:
                stats[date]["[no code / manual]"]["times_used"] += 1
                stats[date]["[no code / manual]"]["total_discount_amount"] += amount
                canonical_display[date]["[no code / manual]"] = "[NO CODE / MANUAL]"

    print(f"  Processed {total_orders_processed:,} discounted orders across {len(stats)} days.")

    rows = []
    for date in sorted(stats.keys()):
        day_stats = stats[date]
        day_rows = [
            {
                "date": date,
                "coupon_code": canonical_display[date].get(norm, norm.upper()),
                "coupon_key": norm,  # normalized key for column mapping
                "times_used": data["times_used"],
                "total_discount_amount": round(data["total_discount_amount"], 2),
            }
            for norm, data in day_stats.items()
        ]
        day_rows.sort(key=lambda r: r["total_discount_amount"], reverse=True)
        rows.extend(day_rows)

    return rows
