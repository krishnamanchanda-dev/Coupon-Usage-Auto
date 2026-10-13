"""
Writes the 'test' tab with a clean styled layout matching the design:
  - 2 cyan decorative rows
  - Yellow merged coupon name headers (2 cols each, no sub-labels)
  - Clean white data rows, comma-formatted numbers
  - Coupon Usage % section below with a gap
"""

import os
import warnings
import datetime
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
OVERALL_TAB_NAME = "Overall"
AW_COL = 49  # column AW, 1-indexed (A=1 … Z=26, AA=27 … AW=49)
EMAIL_DATA_TAB = "Email_Data"

# ── Layout constants ──────────────────────────────────────────────────────────
CYAN_ROWS        = 2   # decorative rows at top
HEADER_ROW       = 3   # coupon name row (1-indexed)
DATA_START       = 4   # first data row
PCT_GAP          = 3   # blank rows between data table and % section
MAX_DATA_ROWS    = 31  # max top-table data rows (enough for a full month)
# % section fixed at row 38: DATA_START(4) + MAX_DATA_ROWS(31) + PCT_GAP(3) = 38
PCT_SECTION_ROW  = DATA_START + MAX_DATA_ROWS + PCT_GAP   # "COUPON USAGE %" label (1-indexed)
PCT_HEADER_ROW   = PCT_SECTION_ROW + 1                    # "Date | SAVE30 | …" header
PCT_DATA_START   = PCT_SECTION_ROW + 2                    # first % data row
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
    """Return (worksheet, is_new). Never clears an existing tab."""
    try:
        ws = sh.worksheet(TEST_TAB_NAME)
        print(f"  Using existing '{TEST_TAB_NAME}' tab.")
        return ws, False
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=TEST_TAB_NAME, rows=500, cols=100)
        print(f"  Created new '{TEST_TAB_NAME}' tab.")
        return ws, True


def _find_tab_state(ws) -> dict:
    """
    Read column A to find what date labels already exist in both tables.
    Returns:
      top_date_labels  — list of existing labels (e.g. "12 Oct") in the top table
      pct_date_labels  — list of existing labels in the % section
      is_first_run     — True if the % section has never been written
    """
    col_a = ws.col_values(1)  # 1-indexed; list of raw strings

    top_date_labels = []
    pct_date_labels = []

    for i, val in enumerate(col_a):
        row_1 = i + 1
        sv = str(val).strip()
        if not sv:
            continue

        if row_1 == PCT_SECTION_ROW:
            continue  # "COUPON USAGE %" label row
        if row_1 == PCT_HEADER_ROW:
            continue  # "Date" header row

        if DATA_START <= row_1 < PCT_SECTION_ROW:
            if sv not in ("Coupons Used", ""):
                top_date_labels.append(sv)
        elif row_1 >= PCT_DATA_START:
            pct_date_labels.append(sv)

    # Check whether the % section has been written at all
    pct_cell = ""
    if len(col_a) >= PCT_SECTION_ROW:
        pct_cell = str(col_a[PCT_SECTION_ROW - 1]).strip()
    is_first_run = "COUPON USAGE" not in pct_cell.upper()

    return {
        "top_date_labels": top_date_labels,
        "pct_date_labels": pct_date_labels,
        "is_first_run": is_first_run,
    }


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


def _parse_cell_date(cell_val) -> "datetime.date | None":
    """Try to parse a spreadsheet cell value as a date. Returns date or None."""
    if not cell_val:
        return None
    s = str(cell_val).strip()
    for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
                "%d %b %Y", "%d-%b-%Y", "%d %B %Y"]:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    # "11 Oct" / "1 Oct" — no year; infer from current IST date
    for fmt in ["%d %b", "%d %B"]:
        try:
            parsed = datetime.datetime.strptime(s, fmt)
            today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5, minutes=30))).date()
            # Try current year; if that puts the date more than 6 months in the future, use last year
            candidate = parsed.replace(year=today.year).date()
            if (candidate - today).days > 180:
                candidate = parsed.replace(year=today.year - 1).date()
            return candidate
        except ValueError:
            pass
    # Google Sheets stores dates as serial numbers (epoch = 1899-12-30)
    try:
        serial = float(s)
        if 1 < serial < 100000:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(serial)))
    except (ValueError, TypeError):
        pass
    return None


def update_overall_tab(sh, aggregated: dict) -> None:
    """Write coupons_used / total_orders ratio to column AW of Overall tab per date."""
    try:
        ws = sh.worksheet(OVERALL_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  WARNING: '{OVERALL_TAB_NAME}' tab not found — skipping Overall update.")
        return

    col_a = ws.col_values(1)  # 1-indexed; returns list of strings

    # Map ISO date string → 1-indexed row number
    date_to_row: dict[str, int] = {}
    for idx, cell_val in enumerate(col_a):
        d = _parse_cell_date(cell_val)
        if d:
            date_to_row[d.strftime("%Y-%m-%d")] = idx + 1

    updates = []
    for date_str, day in aggregated.items():
        row_num = date_to_row.get(date_str)
        if row_num is None:
            print(f"  WARNING: {date_str} not found in Overall tab col A — skipping.")
            continue
        total_orders = day["total_orders"]
        ratio = (day["total_coupon_used"] / total_orders) if total_orders > 0 else 0.0
        cell = gspread.utils.rowcol_to_a1(row_num, AW_COL)  # e.g. "AW5"
        updates.append({"range": cell, "values": [[ratio]]})

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        print(f"  Updated column AW in '{OVERALL_TAB_NAME}' for {len(updates)} date(s).")
    else:
        print(f"  No matching dates found in '{OVERALL_TAB_NAME}' tab.")


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


def write_email_data_tab(sh, aggregated: dict, flat_rows: list = None) -> None:
    """Write a flat, n8n-readable tab with yesterday's coupon data for email pickup.

    Row types:
      meta          — key/value pairs (date, totals)
      named         — individual named coupon rows
      bucket_header — PP / INFLUENCER/MKTNG / OTHERS section header with totals
      bucket_detail — individual coupon codes within a bucket
    """
    try:
        ws = sh.worksheet(EMAIL_DATA_TAB)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=EMAIL_DATA_TAB, rows=500, cols=5)

    date_str = sorted(aggregated.keys())[-1]
    day = aggregated[date_str]
    buckets = day["buckets"]
    total_coupon = day["total_coupon_used"]
    total_orders = day["total_orders"]
    overall_pct = f"{total_coupon / total_orders * 100:.1f}" if total_orders else "0.0"

    # Largest-remainder percentages
    if total_coupon:
        exact = {k: buckets[k]["count"] / total_coupon * 1000 for k in ALL_COLUMNS}
        floored = {k: int(v) for k, v in exact.items()}
        remainder = {k: exact[k] - floored[k] for k in exact}
        extra = 1000 - sum(floored.values())
        for k in sorted(remainder, key=lambda k: remainder[k], reverse=True)[:extra]:
            floored[k] += 1
        pcts = {k: f"{floored[k] / 10:.1f}%" for k in ALL_COLUMNS}
    else:
        pcts = {k: "0.0%" for k in ALL_COLUMNS}

    NAMED_DISPLAY = {k: NAMED_COUPONS[k] for k in NAMED_ORDER}
    BUCKET_LABEL = {"pp": "PP", "influencer": "INFLUENCER / MKTNG", "others": "OTHERS"}

    # Individual coupons within each bucket (sourced from flat_rows)
    bucket_details: dict = {k: [] for k in BUCKET_KEYS}
    if flat_rows:
        for r in flat_rows:
            if r["date"] != date_str:
                continue
            cat = categorize(r["coupon_code"])
            if cat in BUCKET_KEYS:
                bucket_details[cat].append([
                    "bucket_detail",
                    r["coupon_code"],
                    r["times_used"],
                    "",
                    round(r["total_discount_amount"] * 0.25, 0),
                ])
        for k in BUCKET_KEYS:
            bucket_details[k].sort(key=lambda x: x[2], reverse=True)

    # Named coupon rows (non-zero, sorted by count desc)
    named_rows = [
        ["named", NAMED_DISPLAY[k], buckets[k]["count"], pcts[k],
         round(buckets[k]["disc_amt"] * 0.25, 0)]
        for k in NAMED_ORDER if buckets[k]["count"] > 0
    ]
    named_rows.sort(key=lambda r: r[2], reverse=True)

    all_rows = [
        ["row_type", "col1", "col2", "col3", "col4"],  # string headers — no integer key reordering
        ["meta", "date",         date_str,     "", ""],
        ["meta", "total_orders", total_orders, "", ""],
        ["meta", "total_coupon", total_coupon, "", ""],
        ["meta", "overall_pct",  overall_pct,  "", ""],
    ]
    all_rows.extend(named_rows)

    for k in BUCKET_KEYS:
        if buckets[k]["count"] == 0:
            continue
        all_rows.append([
            "bucket_header", BUCKET_LABEL[k],
            buckets[k]["count"], pcts[k],
            round(buckets[k]["disc_amt"] * 0.25, 0),
        ])
        all_rows.extend(bucket_details.get(k, []))

    ws.update("A1", all_rows, value_input_option="RAW")  # RAW prevents "33.1%" being parsed as 0.331
    detail_count = sum(len(v) for v in bucket_details.values())
    print(f"  Email_Data tab: {len(named_rows)} named + {detail_count} bucket-detail rows for {date_str}.")


_MONTHS_IDX = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
}


def _sort_key_label(label: str) -> tuple:
    """Convert "12 Oct" → (10, 12) so rows sort chronologically."""
    parts = label.strip().split()
    if len(parts) == 2:
        try:
            return (_MONTHS_IDX.get(parts[1], 0), int(parts[0]))
        except (ValueError, KeyError):
            pass
    return (0, 0)


def _read_existing_rows(ws) -> tuple:
    """
    Read existing data rows from both tables.
    Returns (top_rows, pct_rows) — lists of lists, one per non-empty date row.
    """
    top_end = DATA_START + MAX_DATA_ROWS - 1
    pct_end = PCT_DATA_START + MAX_DATA_ROWS - 1
    top_raw = ws.get(f"A{DATA_START}:ZZZ{top_end}") or []
    pct_raw = ws.get(f"A{PCT_DATA_START}:ZZZ{pct_end}") or []
    top_rows = [r for r in top_raw if r and str(r[0]).strip()]
    pct_rows = [r for r in pct_raw if r and str(r[0]).strip()]
    return top_rows, pct_rows


def _build_pct_row(date_str: str, day: dict) -> list:
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
        row.extend(["0.0%"] * len(ALL_COLUMNS))
    return row


def _format_data_row(sheet_id, row_0: int, row_idx: int, total_data_cols: int, n_cols: int) -> list:
    """Return formatting requests for one data row in the top table."""
    NUM_FMT = {"type": "NUMBER", "pattern": "#,##0"}
    bg = WHITE if row_idx % 2 == 0 else LIGHT_GREY
    reqs = []
    reqs.append(_range_format(sheet_id, row_0, row_0 + 1, 0, 1,
                               bg=bg, bold=False, font_size=10, h_align="LEFT"))
    reqs.append(_range_format(sheet_id, row_0, row_0 + 1, TOTAL_COL, total_data_cols,
                               bg=bg, bold=False, font_size=10, h_align="RIGHT", number_fmt=NUM_FMT))
    reqs.append(_border_request(sheet_id, row_0, row_0 + 1, 0, total_data_cols, BORDER))
    reqs.append(_set_row_height(sheet_id, row_0, 22))
    for j in range(n_cols):
        left_0 = COUPON_START_COL + j * 2
        reqs.append({
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
    return reqs


def _format_pct_row(sheet_id, row_0: int, row_idx: int) -> list:
    """Return formatting requests for one data row in the % section."""
    bg = WHITE if row_idx % 2 == 0 else LIGHT_GREY
    reqs = []
    reqs.append(_range_format(sheet_id, row_0, row_0 + 1, 0, len(ALL_COLUMNS) + 1,
                               bg=bg, font_size=10, h_align="CENTER"))
    reqs.append(_border_request(sheet_id, row_0, row_0 + 1, 0, len(ALL_COLUMNS) + 1, BORDER))
    reqs.append(_set_row_height(sheet_id, row_0, 22))
    return reqs


def write_test_tab(flat_rows: list, total_orders_by_day: dict) -> None:
    gc = _get_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws, is_new = _get_or_create_tab(sh)
    sheet_id = ws.id

    aggregated = aggregate_for_tab(flat_rows, total_orders_by_day)
    sorted_dates = sorted(aggregated.keys())
    n_cols = len(ALL_COLUMNS)
    total_data_cols = COUPON_START_COL + n_cols * 2

    state = _find_tab_state(ws)
    is_first_run = is_new or state["is_first_run"]

    existing_top = set(state["top_date_labels"])
    new_dates = [d for d in sorted_dates if _day_label(d) not in existing_top]

    if not new_dates and not is_first_run:
        print("  All dates already present — nothing to append.")
        update_overall_tab(sh, aggregated)
        write_email_data_tab(sh, aggregated, flat_rows)
        return

    if is_first_run:
        # ── Full write (first time) ───────────────────────────────────────────
        ws.clear()

        # Coupon header row
        header_row = ["", "Coupons Used"]
        for col_key in ALL_COLUMNS:
            header_row.append(_display(col_key))
            header_row.append("")

        # Top-table data rows
        data_rows = []
        for date_str in sorted_dates:
            day = aggregated[date_str]
            row = [_day_label(date_str), day["total_coupon_used"]]
            for col_key in ALL_COLUMNS:
                b = day["buckets"][col_key]
                row.append(b["count"] if b["count"] > 0 else 0)
                row.append(round(b["disc_amt"] * 0.25, 0) if b["disc_amt"] > 0 else 0)
            data_rows.append(row)

        # % section (fixed at PCT_SECTION_ROW = 38)
        pct_hdr = ["Date"] + [_display(k) for k in ALL_COLUMNS]
        pct_data_rows = [_build_pct_row(d, aggregated[d]) for d in sorted_dates]

        # Write header block (rows 3+)
        blank_gap = [[] for _ in range(PCT_GAP)]
        top_block = [header_row] + data_rows
        ws.update(f"A{HEADER_ROW}", top_block, value_input_option="USER_ENTERED")

        # Write % section at fixed row
        pct_block = [["COUPON USAGE %"], pct_hdr] + pct_data_rows
        ws.update(f"A{PCT_SECTION_ROW}", pct_block, value_input_option="USER_ENTERED")

        print("  Values written. Applying formatting...")

        requests = []

        # Unmerge everything
        requests.append({
            "unmergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0, "endRowIndex": 500,
                    "startColumnIndex": 0, "endColumnIndex": 200,
                }
            }
        })

        # Cyan rows 1 & 2
        requests.append(_range_format(sheet_id, 0, CYAN_ROWS, 0, total_data_cols, bg=CYAN))
        requests.append(_set_row_height(sheet_id, 0, 22))
        requests.append(_set_row_height(sheet_id, 1, 22))

        # Yellow header row
        hdr0 = HEADER_ROW - 1
        requests.append(_range_format(sheet_id, hdr0, hdr0 + 1, 0, total_data_cols,
                                       bg=YELLOW, bold=True, font_size=11,
                                       h_align="CENTER", v_align="MIDDLE"))
        requests.append(_set_row_height(sheet_id, hdr0, 46))

        for i, col_key in enumerate(ALL_COLUMNS):
            lc = COUPON_START_COL + i * 2
            requests.append(_merge_request(sheet_id, hdr0, lc, lc + 2))
            requests.append(_border_request(sheet_id, hdr0, hdr0 + 1, lc, lc + 2, BORDER_DARK))

        # Column widths
        requests.append(_set_col_width(sheet_id, 0, 72))
        requests.append(_set_col_width(sheet_id, TOTAL_COL, 100))
        for i in range(n_cols):
            c = COUPON_START_COL + i * 2
            requests.append(_set_col_width(sheet_id, c, 80))
            requests.append(_set_col_width(sheet_id, c + 1, 95))

        # Top-table data rows
        for i, date_str in enumerate(sorted_dates):
            row_0 = DATA_START - 1 + i
            requests.extend(_format_data_row(sheet_id, row_0, i, total_data_cols, n_cols))

        # % section label row
        pct_lbl_0 = PCT_SECTION_ROW - 1
        requests.append(_range_format(sheet_id, pct_lbl_0, pct_lbl_0 + 1, 0, total_data_cols,
                                       bg=YELLOW, bold=True, font_size=11))
        requests.append(_set_row_height(sheet_id, pct_lbl_0, 36))

        # % section header row
        pct_hdr_0 = PCT_HEADER_ROW - 1
        requests.append(_range_format(sheet_id, pct_hdr_0, pct_hdr_0 + 1, 0, len(ALL_COLUMNS) + 1,
                                       bg=CYAN, bold=True, font_size=10, h_align="CENTER"))
        requests.append(_set_row_height(sheet_id, pct_hdr_0, 36))

        # % data rows
        for i in range(len(sorted_dates)):
            row_0 = PCT_DATA_START - 1 + i
            requests.extend(_format_pct_row(sheet_id, row_0, i))

        _batch_update(sh, requests)
        print(f"  Full write: {len(sorted_dates)} day(s), % section fixed at row {PCT_SECTION_ROW}.")

    else:
        # ── Merge-sort-rewrite: read existing rows, merge with new, sort, rewrite ──
        existing_top_raw, existing_pct_raw = _read_existing_rows(ws)

        # Index existing rows by date label
        existing_top_by_label = {str(r[0]).strip(): r for r in existing_top_raw}
        existing_pct_by_label = {str(r[0]).strip(): r for r in existing_pct_raw}

        # Generate fresh rows for all dates in the current aggregated run
        for date_str in sorted_dates:
            label = _day_label(date_str)
            day = aggregated[date_str]
            top_row = [label, day["total_coupon_used"]]
            for col_key in ALL_COLUMNS:
                b = day["buckets"][col_key]
                top_row.append(b["count"] if b["count"] > 0 else 0)
                top_row.append(round(b["disc_amt"] * 0.25, 0) if b["disc_amt"] > 0 else 0)
            existing_top_by_label[label] = top_row
            existing_pct_by_label[label] = _build_pct_row(date_str, day)

        # Sort all labels chronologically (oldest first)
        sorted_labels = sorted(existing_top_by_label.keys(), key=_sort_key_label)
        sorted_top_rows = [existing_top_by_label[l] for l in sorted_labels]
        sorted_pct_rows = [
            existing_pct_by_label.get(l, [l] + ["0.0%"] * len(ALL_COLUMNS))
            for l in sorted_labels
        ]

        # Clear data sections then rewrite in sorted order
        top_end = DATA_START + MAX_DATA_ROWS - 1
        pct_end = PCT_DATA_START + MAX_DATA_ROWS - 1
        ws.batch_clear([
            f"A{DATA_START}:ZZZ{top_end}",
            f"A{PCT_DATA_START}:ZZZ{pct_end}",
        ])
        ws.update(f"A{DATA_START}", sorted_top_rows, value_input_option="USER_ENTERED")
        ws.update(f"A{PCT_DATA_START}", sorted_pct_rows, value_input_option="USER_ENTERED")

        # Reformat all rows (reset bg first to clear stale alternating colors)
        requests = []
        requests.append(_range_format(sheet_id, DATA_START - 1, top_end, 0, total_data_cols, bg=WHITE))
        requests.append(_range_format(sheet_id, PCT_DATA_START - 1, pct_end, 0, len(ALL_COLUMNS) + 1, bg=WHITE))
        for i, label in enumerate(sorted_labels):
            requests.extend(_format_data_row(sheet_id, DATA_START - 1 + i, i, total_data_cols, n_cols))
            requests.extend(_format_pct_row(sheet_id, PCT_DATA_START - 1 + i, i))

        _batch_update(sh, requests)
        print(f"  Rewrote {len(sorted_labels)} date(s) in sorted order: {', '.join(sorted_labels)}.")

    print(f"\n[5/5] Updating Overall tab column AW...")
    update_overall_tab(sh, aggregated)

    print(f"\n[6/6] Writing Email_Data tab for n8n pickup...")
    write_email_data_tab(sh, aggregated, flat_rows)


def print_categorization(flat_rows: list) -> None:
    print(f"\n{'Coupon Code':<30} {'Category':<20} {'Count':>8} {'Disc Amt':>14}")
    print("-" * 76)
    for r in flat_rows:
        cat = categorize(r["coupon_code"])
        print(f"{r['coupon_code']:<30} {cat:<20} {r['times_used']:>8} {r['total_discount_amount']:>14,.2f}")
