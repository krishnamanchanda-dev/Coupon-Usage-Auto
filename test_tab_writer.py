"""
Writes the 'test' tab with a clean styled layout matching the design:
  - 2 cyan decorative rows
  - Yellow merged coupon name headers (2 cols each, no sub-labels)
  - Clean white data rows, comma-formatted numbers
  - Coupon Usage % section below with a gap
"""

import os
import warnings
warnings.filterwarnings("ignore")

from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials
from categorizer import (
    NAMED_COUPONS, NAMED_ORDER, BUCKET_KEYS, BUCKET_DISPLAY, ALL_COLUMNS, categorize
)

SPREADSHEET_ID = os.getenv("SHEET_ID", "1RBIngrxsQR2fWoksu83OU5YGCW5Idp7DLiVnTKp04fE")
SERVICE_ACCOUNT_FILE = "service_account.json"
TEST_TAB_NAME = "Coupon Usage"

# ── Layout constants ──────────────────────────────────────────────────────────
CYAN_ROWS        = 2   # decorative rows at top
HEADER_ROW       = 3   # coupon name row (1-indexed)
DATA_START       = 4   # first data row
PCT_GAP          = 3   # blank rows between data table and % section
DATE_COL         = 0   # column A (0-indexed)
TOTAL_COL        = 1   # column B: "Coupons Used" (0-indexed)
COUPON_START_COL = 2   # column C onwards: coupon pairs (0-indexed)

# ── Colors ────────────────────────────────────────────────────────────────────
YELLOW = {"red": 1.0,  "green": 0.95, "blue": 0.20}
CYAN   = {"red": 0.53, "green": 0.89, "blue": 0.89}
WHITE  = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
LIGHT_GREY = {"red": 0.95, "green": 0.95, "blue": 0.95}

BORDER = {
    "style": "SOLID",
    "color": {"red": 0.7, "green": 0.7, "blue": 0.7},
}
BORDER_DARK = {
    "style": "SOLID",
    "color": {"red": 0.4, "green": 0.4, "blue": 0.4},
}


def _get_client():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds)


def _display(col_key: str) -> str:
    if col_key in BUCKET_DISPLAY:
        return BUCKET_DISPLAY[col_key]
    return NAMED_COUPONS.get(col_key, col_key.upper())


def _day_label(date_str: str) -> str:
    MONTHS = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
    }
    day = int(date_str[8:10])
    month = int(date_str[5:7])
    return f"{day} {MONTHS[month]}"


def _col_for(coupon_idx: int) -> int:
    """Return 0-indexed sheet column for the COUNT cell of coupon at position coupon_idx (0-based)."""
    return COUPON_START_COL + coupon_idx * 2


def _get_or_create_tab(sh):
    try:
        ws = sh.worksheet(TEST_TAB_NAME)
        ws.clear()
        print(f"  Cleared existing '{TEST_TAB_NAME}' tab.")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=TEST_TAB_NAME, rows=500, cols=100)
        print(f"  Created new '{TEST_TAB_NAME}' tab.")
    return ws


def _batch_update(sh, requests: list):
    sh.batch_update({"requests": requests})


def _rgb(c: dict):
    return {k: v for k, v in c.items()}


def _cell_format(row, col, bg=None, bold=False, font_size=10,
                 h_align="CENTER", v_align="MIDDLE", number_fmt=None) -> dict:
    """Returns a CellFormat update request (0-indexed row/col)."""
    fmt = {
        "textFormat": {"bold": bold, "fontSize": font_size, "fontFamily": "Arial"},
        "horizontalAlignment": h_align,
        "verticalAlignment": v_align,
    }
    if bg:
        fmt["backgroundColor"] = _rgb(bg)
    if number_fmt:
        fmt["numberFormat"] = number_fmt

    return {
        "repeatCell": {
            "range": {
                "sheetId": None,  # filled in caller
                "startRowIndex": row,
                "endRowIndex": row + 1,
                "startColumnIndex": col,
                "endColumnIndex": col + 1,
            },
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,"
                      "verticalAlignment,numberFormat)",
        }
    }


def _range_format(sheet_id, start_row, end_row, start_col, end_col,
                  bg=None, bold=False, font_size=10,
                  h_align="CENTER", v_align="MIDDLE", number_fmt=None) -> dict:
    fmt = {
        "textFormat": {"bold": bold, "fontSize": font_size, "fontFamily": "Arial"},
        "horizontalAlignment": h_align,
        "verticalAlignment": v_align,
    }
    if bg:
        fmt["backgroundColor"] = _rgb(bg)
    if number_fmt:
        fmt["numberFormat"] = number_fmt

    fields = "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment"
    if number_fmt:
        fields += ",numberFormat"
    fields += ")"

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            },
            "cell": {"userEnteredFormat": fmt},
            "fields": fields,
        }
    }


def _merge_request(sheet_id, row, start_col, end_col) -> dict:
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row,
                "endRowIndex": row + 1,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            },
            "mergeType": "MERGE_ALL",
        }
    }


def _border_request(sheet_id, start_row, end_row, start_col, end_col, border) -> dict:
    b = {"style": border["style"], "color": border["color"], "width": 1}
    return {
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            },
            "top": b, "bottom": b, "left": b, "right": b,
            "innerHorizontal": b, "innerVertical": b,
        }
    }


def _set_row_height(sheet_id, row_idx, height_px) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": row_idx,
                "endIndex": row_idx + 1,
            },
            "properties": {"pixelSize": height_px},
            "fields": "pixelSize",
        }
    }


def _set_col_width(sheet_id, col_idx, width_px) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": col_idx,
                "endIndex": col_idx + 1,
            },
            "properties": {"pixelSize": width_px},
            "fields": "pixelSize",
        }
    }


def aggregate_for_tab(flat_rows: list, total_orders_by_day: dict) -> dict:
    result = {}
    by_date = defaultdict(list)
    for r in flat_rows:
        by_date[r["date"]].append(r)

    for date_str, rows in by_date.items():
        buckets = {k: {"count": 0, "disc_amt": 0.0} for k in ALL_COLUMNS}
        for r in rows:
            bucket = categorize(r["coupon_code"])
            if bucket == "excluded":
                continue
            if bucket not in buckets:
                bucket = "others"
            buckets[bucket]["count"] += r["times_used"]
            buckets[bucket]["disc_amt"] += r["total_discount_amount"]
        total_coupon_used = sum(b["count"] for b in buckets.values())
        result[date_str] = {
            "buckets": buckets,
            "total_orders": total_orders_by_day.get(date_str, 0),
            "total_coupon_used": total_coupon_used,
        }
    return result


def write_test_tab(flat_rows: list, total_orders_by_day: dict) -> None:
    gc = _get_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = _get_or_create_tab(sh)
    sheet_id = ws.id

    aggregated = aggregate_for_tab(flat_rows, total_orders_by_day)
    sorted_dates = sorted(aggregated.keys())
    n_cols = len(ALL_COLUMNS)
    total_data_cols = COUPON_START_COL + n_cols * 2  # A + B + (count+amount) × 37

    # ── 1. Write cell values (all in one batch) ───────────────────────────────
    # Coupon header row
    header_row = ["", "Coupons Used"]
    for col_key in ALL_COLUMNS:
        header_row.append(_display(col_key))
        header_row.append("")  # merged away

    # Data rows
    data_rows = []
    for date_str in sorted_dates:
        day = aggregated[date_str]
        row = [_day_label(date_str), day["total_coupon_used"]]
        for col_key in ALL_COLUMNS:
            b = day["buckets"][col_key]
            row.append(b["count"] if b["count"] > 0 else 0)
            row.append(round(b["disc_amt"] * 0.25, 0) if b["disc_amt"] > 0 else 0)
        data_rows.append(row)

    # % section
    pct_label_row = DATA_START + len(sorted_dates) + PCT_GAP
    pct_header_row = pct_label_row + 1
    pct_data_start = pct_header_row + 1

    pct_hdr = ["Date"] + [_display(k) for k in ALL_COLUMNS]

    pct_data_rows = []
    for date_str in sorted_dates:
        day = aggregated[date_str]
        total = day["total_coupon_used"]
        row = [_day_label(date_str)]
        if total:
            exact = {k: day["buckets"][k]["count"] / total * 1000 for k in ALL_COLUMNS}
            floored = {k: int(v) for k, v in exact.items()}
            remainder = {k: exact[k] - floored[k] for k in exact}
            extra = 1000 - sum(floored.values())
            for k in sorted(remainder, key=lambda k: remainder[k], reverse=True)[:extra]:
                floored[k] += 1
            for col_key in ALL_COLUMNS:
                row.append(f"{floored[col_key] / 10:.1f}%")
        else:
            for col_key in ALL_COLUMNS:
                row.append("0.0%")
        pct_data_rows.append(row)

    # Build full block and write in one call
    blank_rows = [[] for _ in range(PCT_GAP)]
    all_rows = (
        [header_row]
        + data_rows
        + blank_rows
        + [["COUPON USAGE %"]]
        + [pct_hdr]
        + pct_data_rows
    )
    ws.update(f"A{HEADER_ROW}", all_rows, value_input_option="USER_ENTERED")

    print("  Values written. Applying formatting...")

    # ── 2. Batch format & style requests ─────────────────────────────────────
    requests = []

    # Unmerge all cells first (clears stale merges from previous runs)
    requests.append({
        "unmergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0, "endRowIndex": 500,
                "startColumnIndex": 0, "endColumnIndex": 200,
            }
        }
    })

    # Cyan rows 1 & 2 (0-indexed: 0, 1)
    requests.append(_range_format(sheet_id, 0, CYAN_ROWS, 0, total_data_cols,
                                  bg=CYAN, bold=False))
    requests.append(_set_row_height(sheet_id, 0, 22))
    requests.append(_set_row_height(sheet_id, 1, 22))

    # Yellow header row (0-indexed: HEADER_ROW - 1 = 2)
    header_row_0 = HEADER_ROW - 1
    requests.append(_range_format(sheet_id, header_row_0, header_row_0 + 1, 0, total_data_cols,
                                  bg=YELLOW, bold=True, font_size=11,
                                  h_align="CENTER", v_align="MIDDLE"))
    requests.append(_set_row_height(sheet_id, header_row_0, 46))

    # Merge coupon name cells in header row + border each pair
    for i, col_key in enumerate(ALL_COLUMNS):
        left_col_0 = COUPON_START_COL + i * 2   # 0-indexed
        right_col_0 = left_col_0 + 1

        # Merge the 2 header cells
        requests.append(_merge_request(sheet_id, header_row_0, left_col_0, right_col_0 + 1))

        # Dark border around each coupon pair in header
        requests.append(_border_request(sheet_id, header_row_0, header_row_0 + 1,
                                        left_col_0, right_col_0 + 1, BORDER_DARK))

    # Date column (col A): width + format
    requests.append(_set_col_width(sheet_id, 0, 72))
    # Coupons Used column (col B)
    requests.append(_set_col_width(sheet_id, TOTAL_COL, 100))
    # Coupon data columns: width
    for i in range(n_cols):
        c = COUPON_START_COL + i * 2
        requests.append(_set_col_width(sheet_id, c, 80))      # count col
        requests.append(_set_col_width(sheet_id, c + 1, 95))  # amount col

    # Data rows: white bg, number format, borders, right-aligned numbers
    NUM_FMT = {"type": "NUMBER", "pattern": "#,##0"}
    last_data_row_0 = DATA_START - 1 + len(sorted_dates)

    for i, date_str in enumerate(sorted_dates):
        row_0 = DATA_START - 1 + i
        bg = WHITE if i % 2 == 0 else LIGHT_GREY

        # Date cell: left-align, plain
        requests.append(_range_format(sheet_id, row_0, row_0 + 1, 0, 1,
                                      bg=bg, bold=False, font_size=10, h_align="LEFT"))

        # Number cells (Coupons Used + all coupon cols): right-align, comma format
        requests.append(_range_format(sheet_id, row_0, row_0 + 1, TOTAL_COL, total_data_cols,
                                      bg=bg, bold=False, font_size=10,
                                      h_align="RIGHT", number_fmt=NUM_FMT))

        # Border for full row
        requests.append(_border_request(sheet_id, row_0, row_0 + 1,
                                        0, total_data_cols, BORDER))

        requests.append(_set_row_height(sheet_id, row_0, 22))

    # Vertical dividers between coupon pairs in data section
    for i in range(n_cols):
        left_0 = COUPON_START_COL + i * 2
        for row_0 in range(DATA_START - 1, last_data_row_0):
            requests.append({
                "updateBorders": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_0,
                        "endRowIndex": row_0 + 1,
                        "startColumnIndex": left_0,
                        "endColumnIndex": left_0 + 2,
                    },
                    "left":  {"style": "SOLID", "color": {"red": 0.4, "green": 0.4, "blue": 0.4}},
                    "right": {"style": "SOLID", "color": {"red": 0.4, "green": 0.4, "blue": 0.4}},
                }
            })

    # % section header label
    pct_lbl_0 = pct_label_row - 1
    requests.append(_range_format(sheet_id, pct_lbl_0, pct_lbl_0 + 1, 0, total_data_cols,
                                  bg=YELLOW, bold=True, font_size=11))
    requests.append(_set_row_height(sheet_id, pct_lbl_0, 36))

    # % section column headers
    pct_hdr_0 = pct_header_row - 1
    requests.append(_range_format(sheet_id, pct_hdr_0, pct_hdr_0 + 1, 0, len(ALL_COLUMNS) + 1,
                                  bg=CYAN, bold=True, font_size=10, h_align="CENTER"))
    requests.append(_set_row_height(sheet_id, pct_hdr_0, 36))

    # % data rows
    for i in range(len(sorted_dates)):
        row_0 = pct_data_start - 1 + i
        bg = WHITE if i % 2 == 0 else LIGHT_GREY
        requests.append(_range_format(sheet_id, row_0, row_0 + 1, 0, len(ALL_COLUMNS) + 1,
                                      bg=bg, font_size=10, h_align="CENTER"))
        requests.append(_border_request(sheet_id, row_0, row_0 + 1,
                                        0, len(ALL_COLUMNS) + 1, BORDER))
        requests.append(_set_row_height(sheet_id, row_0, 22))

    # Send all formatting in one batch
    _batch_update(sh, requests)

    print(f"  Formatted {len(requests)} styling rules.")
    print(f"  {len(sorted_dates)} day(s) × {n_cols} coupon columns.")
    print(f"  Coupon Usage % section at row {pct_label_row}.")


def print_categorization(flat_rows: list) -> None:
    print(f"\n{'Coupon Code':<30} {'Category':<20} {'Count':>8} {'Disc Amt':>14}")
    print("-" * 76)
    for r in flat_rows:
        cat = categorize(r["coupon_code"])
        print(f"{r['coupon_code']:<30} {cat:<20} {r['times_used']:>8} {r['total_discount_amount']:>14,.2f}")
