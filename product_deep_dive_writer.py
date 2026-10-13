"""
Updates the 'Product Deep Dive' tab daily.
Column A: dates (e.g. "12 Oct")
Column D: marketing sessions from GA4 (mktng sources)
Column E: total Shopify human sessions
Column F: E - D (organic/direct sessions)
Column G: sessions_with_cart_additions (Shopify)
Column H: G/F % (add-to-cart rate on organic sessions)
"""

import gspread
from test_tab_writer import _day_label

DEEP_DIVE_TAB = "Product Deep Dive"


def _find_date_row(ws, date_str: str) -> "int | None":
    label = _day_label(date_str)
    col_a = ws.col_values(1)
    for i, val in enumerate(col_a):
        if str(val).strip() == label:
            return i + 1
    return None


def update_product_deep_dive(sh, date_str: str, marketing_sessions: int,
                              total_sessions: int, cart_additions: int) -> None:
    """Write D–H for the row matching date_str."""
    try:
        ws = sh.worksheet(DEEP_DIVE_TAB)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  WARNING: '{DEEP_DIVE_TAB}' tab not found — skipping.")
        return

    row_num = _find_date_row(ws, date_str)
    if row_num is None:
        print(f"  WARNING: Date '{_day_label(date_str)}' not found in '{DEEP_DIVE_TAB}' col A — skipping.")
        return

    organic = total_sessions - marketing_sessions
    # H = G/F % — guard against division by zero
    atc_rate = round(cart_additions / organic * 100, 2) if organic > 0 else 0.0

    ws.update(
        f"D{row_num}:H{row_num}",
        [[marketing_sessions, total_sessions, organic, cart_additions, atc_rate]],
        value_input_option="RAW",
    )
    print(
        f"  Product Deep Dive row {row_num}: "
        f"D={marketing_sessions}, E={total_sessions}, F={organic}, "
        f"G={cart_additions}, H={atc_rate}%"
    )
