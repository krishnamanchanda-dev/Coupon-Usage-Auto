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
COUPON_PCT_COL   = 2   # column C: "Coupon Usage %" (0-indexed)
# Summary columns D, E, F (0-indexed 3, 4, 5)
TOTAL_ORDERS_COL = 3   # column D: Total Orders
EXCLUDED_COL     = 4   # column E: Excluded Count
NO_COUPON_COL    = 5   # column F: No Coupon Count
COUPON_START_COL = 6   # column G onwards: coupon pairs (0-indexed)

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

    # Detect layout version: new layout has "Total Orders" at column D (0-indexed 3) of header row
    row3 = ws.row_values(HEADER_ROW) if not is_first_run else []
    old_layout = not is_first_run and (
        len(row3) < COUPON_START_COL or
        str(row3[TOTAL_ORDERS_COL]).strip() != "Total Orders"
    )
    # Old layout → treat as first run so we clear and rewrite with new layout
    if old_layout:
        is_first_run = True

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


# Mapping: our coupon key → (count_col, amount_col) in Overall tab, 1-indexed.
# amount_col=None means the column is a single-column bucket aggregate (count only).
# Column numbers below reflect the layout AFTER 2 columns are inserted before the old CP.
# Old CP (Others=94) shifts to CR (96); old CQ (Influencer=95) shifts to CS (97); etc.
# New CP (94) = Excluded count, new CQ (95) = No Coupon count.
OVERALL_COUPON_COLS: dict[str, tuple] = {
    # New summary columns inserted before Others
    '__excluded__':  (94,  None),  # CP  "Excluded"
    '__no_coupon__': (95,  None),  # CQ  "No Coupon"
    # Bucket aggregates — count only (all shifted +2)
    'others':        (96,  None),  # CR  "Other"
    'influencer':    (97,  None),  # CS  "Influencer/MKTNG"
    'pp':            (128, None),  # DX  "PP"
    # Named coupons — (count, amount) pairs (all shifted +2)
    '5pcs':         (98,  99),    # CT  "5Pcs - 1999"
    '3pcs':         (100, 101),   # CV  "3Pcs - 899"
    'b3p2':         (102, 103),   # CX  "B3P2"
    'buy3at699':    (104, 105),   # CZ  "Buy 3 at 699"
    'freevanity':   (106, 107),   # DB  "Vanity/FREEVANITY"
    'luxevanity':   (108, 109),   # DD  "LUXEVANITY"
    'freelipstick': (110, 111),   # DF  "FreeLipstick"
    'lips5':        (112, 113),   # DH  "LIPS5"
    'freekit':      (114, 115),   # DJ  "FREEKIT"
    'lipluxe':      (116, 117),   # DL  "LIPLUXE"
    'modern20':     (118, 119),   # DN  "MODERN20"
    'b2g2':         (120, 121),   # DP  "B2G2"
    'buy6':         (122, 123),   # DR  "BUY6"
    'buy10':        (124, 125),   # DT  "BUY10"
    'get4':         (126, 127),   # DV  "GET4"
    'free6':        (129, 130),   # DY  "FREE6"
    'free5':        (131, 132),   # EA  "FREE5"
    'special4':     (133, 134),   # EC  "Special4"
    'special5':     (135, 136),   # EE  "Special5"
    'special6':     (137, 138),   # EG  "Special6"
    'buy5':         (139, 140),   # EI  "Buy5"
    'gpay':         (141, 142),   # EK  "Gpay"
    'buy4':         (143, 144),   # EM  "Buy4"
    'save25':       (145, 146),   # EO  "SAVE25"
    'save30':       (147, 148),   # EQ  "SAVE30"
    'save35':       (149, 150),   # ES  "SAVE35"
    'stealdeal':    (151, 152),   # EU  "STEALDEAL"
    'fullcb':       (153, 154),   # EW  "FULLCB"
    'luxe3':        (157, 158),   # EZ  "LUXE3"
    'buy3':         (161, 162),   # FE  "BUY3"
    'get2':         (163, 164),   # FG  "Get2"
    '6pcskit':      (165, 166),   # FI  "6 Pcs Kit"
}

# Mapping: coupon key → 1-indexed column in Overall tab for per-coupon USAGE % (BM–CL).
OVERALL_PCT_COLS: dict[str, int] = {
    '__excluded__':  65,   # BM  "Excluded"
    '4pcs':          66,   # BN  "4Pcs"
    '3pcs':          67,   # BO  "3Pcs (899)"
    'freevanity':    68,   # BP  "Vanity 14.6%"
    'luxevanity':    69,   # BQ  "LUXEVANITY"
    'freelipstick':  70,   # BR  "Lipstick 5.9%"
    'lips5':         71,   # BS  "LIPS5"
    'free5':         72,   # BT  "Free5"
    'free6':         73,   # BU  "Free6"
    'freekit':       74,   # BV  "Freekit"
    'lipluxe':       75,   # BW  "Lipluxe"
    'save25':        76,   # BX  "SAVE25"
    'save30':        77,   # BY  "SAVE30"
    'save35':        78,   # BZ  "SAVE35"
    'stealdeal':     79,   # CA  "STEALDEAL"
    'pp':            80,   # CB  "Partnership"
    '__no_coupon__': 81,   # CC  "No Coupon"
    'influencer':    82,   # CD  "Influencer/MKTNG"
    'gpay':          83,   # CE  "Sale Offer"
    'others':        84,   # CF  "Others"
    'special4':      85,   # CG  "Special4"
    'special5':      86,   # CH  "Special5"
    'special6':      87,   # CI  "Special6"
    'b3p2':          88,   # CJ  "B3P2"
    'modern20':      89,   # CK  "MODERN20"
    'buy3':          90,   # CL  "BUY3"
}


def _ensure_overall_excluded_cols(ws, sh) -> None:
    """
    One-time migration: insert 2 columns before the old CP (Others) column so that
    CP = Excluded count and CQ = No Coupon count, with Others shifting to CR.
    Idempotent — skips if CP header is already 'Excluded'.
    """
    cp_header = str(ws.cell(3, 94).value or "").strip().lower()
    if cp_header == "excluded":
        return  # already done

    sheet_id = ws.id
    # Insert 2 blank columns at 0-indexed position 93 (before 1-indexed col 94 = CP)
    sh.batch_update({"requests": [{
        "insertDimension": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": 93,
                "endIndex": 95,
            },
            "inheritFromBefore": False,
        }
    }]})
    # Write column headers in row 3 (matching the Overall tab's header row)
    ws.update("CP3", [["Excluded", "No Coupon"]], value_input_option="RAW")
    print("  One-time migration: inserted Excluded (CP) and No Coupon (CQ) before Others.")


def update_overall_coupon_cols(sh, aggregated: dict) -> None:
    """Write per-coupon count and 0.25× disc_amt into Overall tab columns CP–FI."""
    try:
        ws = sh.worksheet(OVERALL_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  WARNING: '{OVERALL_TAB_NAME}' tab not found — skipping coupon col update.")
        return

    _ensure_overall_excluded_cols(ws, sh)

    col_a = ws.col_values(1)
    date_to_row: dict[str, int] = {}
    for idx, cell_val in enumerate(col_a):
        d = _parse_cell_date(cell_val)
        if d:
            date_to_row[d.strftime("%Y-%m-%d")] = idx + 1

    updates = []
    for date_str, day in aggregated.items():
        row_num = date_to_row.get(date_str)
        if row_num is None:
            print(f"  WARNING: {date_str} not found in Overall tab — skipping coupon cols.")
            continue
        buckets = day["buckets"]
        no_coupon = max(0, day["total_orders"] - day["total_coupon_used"] - day["excluded_count"])
        for key, (count_col, amount_col) in OVERALL_COUPON_COLS.items():
            if key == "__excluded__":
                count = day["excluded_count"]
            elif key == "__no_coupon__":
                count = no_coupon
            else:
                b = buckets.get(key, {"count": 0, "disc_amt": 0.0})
                count = b["count"]
                if amount_col is not None:
                    amt = round(b["disc_amt"] * 0.25, 0)
                    amt_cell = gspread.utils.rowcol_to_a1(row_num, amount_col)
                    updates.append({"range": amt_cell, "values": [[amt]]})
            count_cell = gspread.utils.rowcol_to_a1(row_num, count_col)
            updates.append({"range": count_cell, "values": [[count]]})

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        print(f"  Updated coupon cols (CP–FI) in '{OVERALL_TAB_NAME}' for {len(aggregated)} date(s).")
    else:
        print(f"  No matching dates found in '{OVERALL_TAB_NAME}' for coupon cols.")


def update_overall_pct_cols(sh, aggregated: dict) -> None:
    """Write per-coupon usage % (as decimals) into Overall tab columns BM–CL."""
    try:
        ws = sh.worksheet(OVERALL_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  WARNING: '{OVERALL_TAB_NAME}' tab not found — skipping pct col update.")
        return

    col_a = ws.col_values(1)
    date_to_row: dict[str, int] = {}
    for idx, cell_val in enumerate(col_a):
        d = _parse_cell_date(cell_val)
        if d:
            date_to_row[d.strftime("%Y-%m-%d")] = idx + 1

    updates = []
    for date_str, day in aggregated.items():
        row_num = date_to_row.get(date_str)
        if row_num is None:
            continue
        total_ord = day["total_orders"]
        if not total_ord:
            continue

        # Largest-remainder percentages (same logic as _build_pct_row)
        no_coupon = max(0, total_ord - day["total_coupon_used"] - day["excluded_count"])
        all_keys = list(ALL_COLUMNS) + ["__excluded__", "__no_coupon__"]
        all_counts = {k: day["buckets"][k]["count"] for k in ALL_COLUMNS}
        all_counts["__excluded__"] = day["excluded_count"]
        all_counts["__no_coupon__"] = no_coupon
        exact = {k: all_counts[k] / total_ord * 1000 for k in all_keys}
        floored = {k: int(v) for k, v in exact.items()}
        remainder = {k: exact[k] - floored[k] for k in exact}
        extra = 1000 - sum(floored.values())
        for k in sorted(remainder, key=lambda k: remainder[k], reverse=True)[:extra]:
            floored[k] += 1

        for key, col_num in OVERALL_PCT_COLS.items():
            pct_decimal = floored.get(key, 0) / 1000  # e.g. 0.223
            cell = gspread.utils.rowcol_to_a1(row_num, col_num)
            updates.append({"range": cell, "values": [[pct_decimal]]})

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        print(f"  Updated pct cols (BM–CL) in '{OVERALL_TAB_NAME}' for {len(aggregated)} date(s).")
    else:
        print(f"  No matching dates found in '{OVERALL_TAB_NAME}' for pct cols.")


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
    for date_str in aggregated:
        row_num = date_to_row.get(date_str)
        if row_num is None:
            print(f"  WARNING: {date_str} not found in Overall tab col A — skipping.")
            continue
        # Read total_orders from Coupon Usage tab col D — same source as the emailer
        tab_row = _read_coupon_usage_row(sh, date_str)
        if tab_row is None:
            print(f"  WARNING: {date_str} not found in Coupon Usage tab — skipping AW.")
            continue
        total_orders = tab_row["total_orders"]
        coupons_used = tab_row["total_coupon_used"]
        ratio = (coupons_used / total_orders) if total_orders > 0 else 0.0
        cell = gspread.utils.rowcol_to_a1(row_num, AW_COL)  # e.g. "AW5"
        updates.append({"range": cell, "values": [[ratio]]})

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        print(f"  Updated column AW in '{OVERALL_TAB_NAME}' for {len(updates)} date(s).")
    else:
        print(f"  No matching dates found in '{OVERALL_TAB_NAME}' tab.")


# AJ = col 36 (new customer rate), AK = col 37 (returning customer rate)
OVERALL_NEW_CUST_COL       = 36   # AJ
OVERALL_RETURNING_CUST_COL = 37   # AK


def update_overall_customer_rates(sh, date_str: str, returning_rate: float) -> None:
    """
    Write new_customer_rate (AJ) and returning_customer_rate (AK) for a single date.
    returning_rate is a float (e.g. 0.3554); new_rate = 1 - returning_rate.
    Both written as decimals so the sheet can format them as %.
    """
    try:
        ws = sh.worksheet(OVERALL_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  WARNING: '{OVERALL_TAB_NAME}' tab not found — skipping customer rates.")
        return

    col_a = ws.col_values(1)
    row_num = None
    for idx, cell_val in enumerate(col_a):
        d = _parse_cell_date(cell_val)
        if d and d.strftime("%Y-%m-%d") == date_str:
            row_num = idx + 1
            break

    if row_num is None:
        print(f"  WARNING: {date_str} not found in Overall tab — skipping customer rates.")
        return

    new_rate = round(1.0 - returning_rate, 6)
    aj = gspread.utils.rowcol_to_a1(row_num, OVERALL_NEW_CUST_COL)
    ak = gspread.utils.rowcol_to_a1(row_num, OVERALL_RETURNING_CUST_COL)
    ws.batch_update(
        [
            {"range": aj, "values": [[new_rate]]},
            {"range": ak, "values": [[returning_rate]]},
        ],
        value_input_option="USER_ENTERED",
    )
    print(
        f"  Overall AJ/AK for {date_str}: new={new_rate:.2%}, returning={returning_rate:.2%}"
    )


def aggregate_for_tab(flat_rows: list, total_orders_by_day: dict) -> dict:
    result = {}
    by_date = defaultdict(list)
    for r in flat_rows:
        by_date[r["date"]].append(r)

    for date_str, rows in by_date.items():
        buckets = {k: {"count": 0, "disc_amt": 0.0} for k in ALL_COLUMNS}
        excluded_count = 0
        for r in rows:
            bucket = categorize(r["coupon_code"])
            if bucket == "excluded":
                excluded_count += r["times_used"]
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
            "excluded_count": excluded_count,
        }
    return result


def _read_coupon_usage_row(sh, date_str: str) -> "dict | None":
    """
    Read a single date's data row from the Coupon Usage tab.
    Returns normalised dict with disc values already 0.25×, or None if not found.
    """
    try:
        ws = sh.worksheet(TEST_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return None

    target_label = _day_label(date_str)
    col_a = ws.col_values(1)
    row_num = next((i + 1 for i, v in enumerate(col_a) if str(v).strip() == target_label), None)
    if row_num is None:
        return None

    row = ws.row_values(row_num)

    def _n(idx):
        try:
            return float(str(row[idx]).replace(",", "").strip()) if len(row) > idx else 0.0
        except (ValueError, TypeError):
            return 0.0

    buckets = {}
    for i, col_key in enumerate(ALL_COLUMNS):
        c_idx = COUPON_START_COL + i * 2
        a_idx = c_idx + 1
        # disc_amt column in the sheet already stores the 0.25× value
        buckets[col_key] = {"count": int(_n(c_idx)), "disc": _n(a_idx)}

    return {
        "total_coupon_used": int(_n(TOTAL_COL)),
        "total_orders":      int(_n(TOTAL_ORDERS_COL)),
        "excluded_count":    int(_n(EXCLUDED_COL)),
        "no_coupon":         int(_n(NO_COUPON_COL)),
        "buckets":           buckets,
    }


def write_email_data_tab(sh, aggregated: dict, flat_rows: list = None) -> None:
    """Write a flat, n8n-readable tab with yesterday's coupon data for email pickup.

    Summary numbers (total_orders, coupons_used, excluded, per-coupon counts/disc)
    are read from the Coupon Usage tab so they match exactly what is shown there.
    Individual bucket/excluded code breakdowns still come from flat_rows.

    Row types:
      meta          — key/value pairs (date, totals)
      named         — individual named coupon rows
      bucket_header — PP / INFLUENCER/MKTNG / OTHERS / EXCLUDED section header
      bucket_detail — individual coupon codes within a bucket

    col4 (disc_amt) is only populated for groups PP, OTHERS, EXCLUDED and for
    specific named coupons; all others get "" so the emailer renders "—".
    """
    SHOW_DISC_NAMED = {
        "FREELIPSTICK", "MODERN20", "FREEMAKEUPVANITY",
        "FREEVANITY", "LUXEVANITY", "FREEPERFUME",
    }
    SHOW_DISC_BUCKETS: set = set()  # no group headers ever show disc amt

    try:
        ws = sh.worksheet(EMAIL_DATA_TAB)
        ws.clear()
        # clear() removes values but not cell formats; reset number format so
        # integer values don't render as "202800.00%" in percentage-formatted cells
        sh.batch_update({"requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": 0, "endRowIndex": 500,
                    "startColumnIndex": 0, "endColumnIndex": 10,
                },
                "cell": {"userEnteredFormat": {}},
                "fields": "userEnteredFormat.numberFormat",
            }
        }]})
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=EMAIL_DATA_TAB, rows=500, cols=5)

    date_str = sorted(aggregated.keys())[-1]

    # ── Source of truth: Coupon Usage tab ────────────────────────────────────
    tab = _read_coupon_usage_row(sh, date_str)
    if tab:
        total_coupon  = tab["total_coupon_used"]
        total_orders  = tab["total_orders"]
        excluded_count = tab["excluded_count"]
        # buckets[k]["count"] and buckets[k]["disc"] (already 0.25×) from the tab
        buckets = tab["buckets"]
        print(f"  Email_Data: using Coupon Usage tab values for {_day_label(date_str)}.")
    else:
        # Fallback to aggregated if tab row not found
        day = aggregated[date_str]
        total_coupon   = day["total_coupon_used"]
        total_orders   = day["total_orders"]
        excluded_count = day["excluded_count"]
        buckets = {k: {"count": day["buckets"][k]["count"],
                       "disc":  round(day["buckets"][k]["disc_amt"] * 0.25, 0)}
                   for k in ALL_COLUMNS}
        print(f"  Email_Data: Coupon Usage tab row not found, falling back to aggregated.")

    overall_pct = f"{total_coupon / total_orders * 100:.1f}" if total_orders else "0.0"

    # Largest-remainder percentages over total_coupon for named/bucket rows
    if total_coupon:
        exact   = {k: buckets[k]["count"] / total_coupon * 1000 for k in ALL_COLUMNS}
        floored = {k: int(v) for k, v in exact.items()}
        rem     = {k: exact[k] - floored[k] for k in exact}
        extra   = 1000 - sum(floored.values())
        for k in sorted(rem, key=lambda k: rem[k], reverse=True)[:extra]:
            floored[k] += 1
        pcts = {k: f"{floored[k] / 10:.1f}%" for k in ALL_COLUMNS}
    else:
        pcts = {k: "0.0%" for k in ALL_COLUMNS}

    excl_pct = f"{excluded_count / total_orders * 100:.1f}%" if total_orders else "0.0%"

    NAMED_DISPLAY = {k: NAMED_COUPONS[k] for k in NAMED_ORDER}
    BUCKET_LABEL  = {
        "pp": "PP", "influencer": "INFLUENCER / MKTNG",
        "others": "OTHERS", "excluded": "EXCLUDED",
    }

    # Individual codes within each bucket + excluded (from flat_rows — not in tab)
    bucket_details: dict = {k: [] for k in list(BUCKET_KEYS) + ["excluded"]}
    excluded_total_disc = 0.0
    if flat_rows:
        for r in flat_rows:
            if r["date"] != date_str:
                continue
            cat = categorize(r["coupon_code"])
            disc_val = round(r["total_discount_amount"] * 0.25, 0)
            if cat in BUCKET_KEYS:
                bucket_details[cat].append([
                    "bucket_detail", r["coupon_code"], r["times_used"], "", "",
                ])
            elif cat == "excluded":
                excluded_total_disc += r["total_discount_amount"]
                bucket_details["excluded"].append([
                    "bucket_detail", r["coupon_code"], r["times_used"], "", "",
                ])
        for k in list(BUCKET_KEYS) + ["excluded"]:
            bucket_details[k].sort(key=lambda x: x[2], reverse=True)

    # Named coupon rows — use counts from tab; disc only for allowed coupons
    named_rows = []
    for k in NAMED_ORDER:
        if buckets[k]["count"] == 0:
            continue
        display = NAMED_DISPLAY[k]
        disc = buckets[k]["disc"] if display in SHOW_DISC_NAMED else ""
        named_rows.append(["named", display, buckets[k]["count"], pcts[k], disc])
    named_rows.sort(key=lambda r: r[2], reverse=True)

    all_rows = [
        ["row_type", "col1", "col2", "col3", "col4"],
        ["meta", "date",         date_str,     "", ""],
        ["meta", "total_orders", total_orders, "", ""],
        ["meta", "total_coupon", total_coupon, "", ""],
        ["meta", "overall_pct",  overall_pct,  "", ""],
    ]
    all_rows.extend(named_rows)

    for k in BUCKET_KEYS:
        if buckets[k]["count"] == 0:
            continue
        all_rows.append(["bucket_header", BUCKET_LABEL[k], buckets[k]["count"], pcts[k], ""])
        all_rows.extend(bucket_details.get(k, []))

    if excluded_count > 0:
        all_rows.append(["bucket_header", BUCKET_LABEL["excluded"], excluded_count, excl_pct, ""])
        all_rows.extend(bucket_details.get("excluded", []))

    ws.update("A1", all_rows, value_input_option="RAW")
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


def _read_overall_orders(sh) -> dict:
    """Read date → total orders from Overall tab column C (authoritative source)."""
    try:
        ws = sh.worksheet(OVERALL_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return {}
    col_a = ws.col_values(1)
    col_c = ws.col_values(3)
    result = {}
    for date_val, order_val in zip(col_a, col_c):
        d = _parse_cell_date(date_val)
        if d and order_val:
            try:
                result[d.strftime("%Y-%m-%d")] = int(str(order_val).replace(",", "").strip())
            except ValueError:
                pass
    return result


def _build_pct_row(date_str: str, day: dict) -> list:
    """Build % row using total_orders as denominator; guaranteed to sum to 100%."""
    total_ord = day["total_orders"]
    row = [_day_label(date_str)]
    if total_ord:
        no_coupon = max(0, total_ord - day["total_coupon_used"] - day["excluded_count"])
        # Apply largest-remainder across ALL categories so they sum exactly to 100%
        all_keys = list(ALL_COLUMNS) + ["__excluded__", "__no_coupon__"]
        all_counts = {k: day["buckets"][k]["count"] for k in ALL_COLUMNS}
        all_counts["__excluded__"] = day["excluded_count"]
        all_counts["__no_coupon__"] = no_coupon
        exact = {k: all_counts[k] / total_ord * 1000 for k in all_keys}
        floored = {k: int(v) for k, v in exact.items()}
        remainder = {k: exact[k] - floored[k] for k in exact}
        extra = 1000 - sum(floored.values())
        for k in sorted(remainder, key=lambda k: remainder[k], reverse=True)[:extra]:
            floored[k] += 1
        for col_key in ALL_COLUMNS:
            row.append(f"{floored[col_key] / 10:.1f}%")
        row.append(f"{floored['__excluded__'] / 10:.1f}%")
        row.append(f"{floored['__no_coupon__'] / 10:.1f}%")
    else:
        row.extend(["0.0%"] * (len(ALL_COLUMNS) + 2))
    return row


def _format_data_row(sheet_id, row_0: int, row_idx: int, total_data_cols: int, n_cols: int) -> list:
    """Return formatting requests for one data row in the top table."""
    NUM_FMT = {"type": "NUMBER", "pattern": "#,##0"}
    PCT_FMT = {"type": "NUMBER", "pattern": "0.0%"}
    bg = WHITE if row_idx % 2 == 0 else LIGHT_GREY
    reqs = []
    reqs.append(_range_format(sheet_id, row_0, row_0 + 1, 0, 1,
                               bg=bg, bold=False, font_size=10, h_align="LEFT"))
    reqs.append(_range_format(sheet_id, row_0, row_0 + 1, TOTAL_COL, total_data_cols,
                               bg=bg, bold=False, font_size=10, h_align="RIGHT", number_fmt=NUM_FMT))
    # Override Coupon Usage % column with percentage format (not integer)
    reqs.append(_range_format(sheet_id, row_0, row_0 + 1, COUPON_PCT_COL, COUPON_PCT_COL + 1,
                               bg=bg, bold=False, font_size=10, h_align="RIGHT", number_fmt=PCT_FMT))
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
    pct_total_cols = len(ALL_COLUMNS) + 3  # Date + coupons + Excluded% + No Coupon%
    reqs = []
    reqs.append(_range_format(sheet_id, row_0, row_0 + 1, 0, pct_total_cols,
                               bg=bg, font_size=10, h_align="CENTER"))
    reqs.append(_border_request(sheet_id, row_0, row_0 + 1, 0, pct_total_cols, BORDER))
    reqs.append(_set_row_height(sheet_id, row_0, 22))
    return reqs


def write_test_tab(flat_rows: list, total_orders_by_day: dict) -> None:
    gc = _get_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws, is_new = _get_or_create_tab(sh)
    sheet_id = ws.id

    aggregated = aggregate_for_tab(flat_rows, total_orders_by_day)
    sorted_dates = sorted(aggregated.keys())

    # Override total_orders with authoritative values from Overall tab col C
    overall_orders = _read_overall_orders(sh)
    for date_str, day in aggregated.items():
        if date_str in overall_orders:
            day["total_orders"] = overall_orders[date_str]

    n_cols = len(ALL_COLUMNS)
    total_data_cols = COUPON_START_COL + n_cols * 2

    state = _find_tab_state(ws)
    is_first_run = is_new or state["is_first_run"]

    existing_top = set(state["top_date_labels"])
    new_dates = [d for d in sorted_dates if _day_label(d) not in existing_top]

    if not new_dates and not is_first_run:
        print("  All dates already present — nothing to append.")
        update_overall_tab(sh, aggregated)
        update_overall_coupon_cols(sh, aggregated)
        update_overall_pct_cols(sh, aggregated)
        write_email_data_tab(sh, aggregated, flat_rows)
        return

    if is_first_run:
        # ── Full write (layout migration or new tab) ──────────────────────────
        # Safety: read existing rows BEFORE clearing so historical dates survive
        # even when a layout migration forces a full rewrite.
        preserved_top = {}   # label → row list
        preserved_pct = {}   # label → row list
        if not is_new:
            existing_top_raw, existing_pct_raw = _read_existing_rows(ws)
            new_labels = {_day_label(d) for d in aggregated}
            for r in existing_top_raw:
                label = str(r[0]).strip()
                if label and label not in new_labels:
                    preserved_top[label] = r
            for r in existing_pct_raw:
                label = str(r[0]).strip()
                if label and label not in new_labels:
                    preserved_pct[label] = r
            if preserved_top:
                print(f"  Preserving {len(preserved_top)} existing row(s): {', '.join(sorted(preserved_top, key=_sort_key_label))}.")
                # Migrate old-layout rows (5 header cols) to new layout (6 header cols)
                # by inserting Coupon Usage % at index 2.
                old_header_cols = COUPON_START_COL - 1  # old layout had one fewer header col
                for lbl in list(preserved_top.keys()):
                    r = preserved_top[lbl]
                    if len(r) > 0 and len(r) < COUPON_START_COL + 1:
                        try:
                            tot = float(str(r[TOTAL_ORDERS_COL - 1]).replace(",", "")) if len(r) > TOTAL_ORDERS_COL - 1 else 0
                            used = float(str(r[TOTAL_COL]).replace(",", "")) if len(r) > TOTAL_COL else 0
                            pct = used / tot if tot > 0 else 0.0
                        except (ValueError, TypeError):
                            pct = 0.0
                        preserved_top[lbl] = list(r[:2]) + [pct] + list(r[2:])

        ws.clear()

        # Step 1: unmerge BEFORE writing values — ws.clear() doesn't clear merges,
        # so old merge ranges would hide values written to secondary cells.
        _batch_update(sh, [{
            "unmergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0, "endRowIndex": 500,
                    "startColumnIndex": 0, "endColumnIndex": 200,
                }
            }
        }])

        # Coupon header row (includes Coupon Usage %, Total Orders, Excluded, No Coupon after Coupons Used)
        header_row = ["", "Coupons Used", "Coupon Usage %", "Total Orders", "Excluded", "No Coupon"]
        for col_key in ALL_COLUMNS:
            header_row.append(_display(col_key))
            header_row.append("")

        # Build fresh rows for dates in current aggregated run
        fresh_top = {}
        fresh_pct = {}
        for date_str in sorted_dates:
            label = _day_label(date_str)
            day = aggregated[date_str]
            no_coupon = max(0, day["total_orders"] - day["total_coupon_used"] - day["excluded_count"])
            pct = day["total_coupon_used"] / day["total_orders"] if day["total_orders"] > 0 else 0.0
            row = [label, day["total_coupon_used"], pct,
                   day["total_orders"], day["excluded_count"], no_coupon]
            for col_key in ALL_COLUMNS:
                b = day["buckets"][col_key]
                row.append(b["count"] if b["count"] > 0 else 0)
                row.append(round(b["disc_amt"] * 0.25, 0) if b["disc_amt"] > 0 else 0)
            fresh_top[label] = row
            fresh_pct[label] = _build_pct_row(date_str, day)

        # Merge fresh + preserved, sort chronologically
        all_top = {**preserved_top, **fresh_top}
        all_pct = {**preserved_pct, **fresh_pct}
        all_labels = sorted(all_top.keys(), key=_sort_key_label)
        data_rows = [all_top[l] for l in all_labels]
        pct_data_rows_merged = [
            all_pct.get(l, [l] + ["0.0%"] * (len(ALL_COLUMNS) + 2))
            for l in all_labels
        ]

        # % section (fixed at PCT_SECTION_ROW = 38)
        pct_hdr = ["Date"] + [_display(k) for k in ALL_COLUMNS] + ["Excluded %", "No Coupon %"]

        # Step 2: write values (merges are now gone so all cells are writable)
        top_block = [header_row] + data_rows
        ws.update(f"A{HEADER_ROW}", top_block, value_input_option="USER_ENTERED")

        pct_block = [["COUPON USAGE %"], pct_hdr] + pct_data_rows_merged
        ws.update(f"A{PCT_SECTION_ROW}", pct_block, value_input_option="USER_ENTERED")

        print("  Values written. Applying formatting...")

        # Step 3: format + re-merge (no unmergeCells here)
        requests = []

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

        # Summary columns (Coupon Usage %, Total Orders, Excluded, No Coupon) — cyan bg, bold
        requests.append(_range_format(sheet_id, hdr0, hdr0 + 1,
                                       COUPON_PCT_COL, NO_COUPON_COL + 1,
                                       bg=CYAN, bold=True, font_size=11,
                                       h_align="CENTER", v_align="MIDDLE"))

        # Column widths
        requests.append(_set_col_width(sheet_id, 0, 72))
        requests.append(_set_col_width(sheet_id, TOTAL_COL, 100))
        requests.append(_set_col_width(sheet_id, COUPON_PCT_COL, 110))
        requests.append(_set_col_width(sheet_id, TOTAL_ORDERS_COL, 100))
        requests.append(_set_col_width(sheet_id, EXCLUDED_COL, 90))
        requests.append(_set_col_width(sheet_id, NO_COUPON_COL, 100))
        for i in range(n_cols):
            c = COUPON_START_COL + i * 2
            requests.append(_set_col_width(sheet_id, c, 80))
            requests.append(_set_col_width(sheet_id, c + 1, 95))

        # Top-table data rows (all_labels includes preserved + fresh, sorted)
        for i in range(len(all_labels)):
            row_0 = DATA_START - 1 + i
            requests.extend(_format_data_row(sheet_id, row_0, i, total_data_cols, n_cols))

        # % section label row
        pct_lbl_0 = PCT_SECTION_ROW - 1
        requests.append(_range_format(sheet_id, pct_lbl_0, pct_lbl_0 + 1, 0, total_data_cols,
                                       bg=YELLOW, bold=True, font_size=11))
        requests.append(_set_row_height(sheet_id, pct_lbl_0, 36))

        # % section header row  (Date + 37 coupons + Excluded% + No Coupon% = n+3 cols)
        pct_hdr_0 = PCT_HEADER_ROW - 1
        pct_total_cols = len(ALL_COLUMNS) + 3
        requests.append(_range_format(sheet_id, pct_hdr_0, pct_hdr_0 + 1, 0, pct_total_cols,
                                       bg=CYAN, bold=True, font_size=10, h_align="CENTER"))
        requests.append(_set_row_height(sheet_id, pct_hdr_0, 36))

        # % data rows
        for i in range(len(all_labels)):
            row_0 = PCT_DATA_START - 1 + i
            requests.extend(_format_pct_row(sheet_id, row_0, i))

        _batch_update(sh, requests)
        print(f"  Full write: {len(all_labels)} day(s) ({len(fresh_top)} new + {len(preserved_top)} preserved), "
              f"% section fixed at row {PCT_SECTION_ROW}.")

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
            no_coupon = max(0, day["total_orders"] - day["total_coupon_used"] - day["excluded_count"])
            pct = day["total_coupon_used"] / day["total_orders"] if day["total_orders"] > 0 else 0.0
            top_row = [label, day["total_coupon_used"], pct,
                       day["total_orders"], day["excluded_count"], no_coupon]
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

    print(f"\n[5b] Updating Overall tab coupon columns (CP–FH)...")
    update_overall_coupon_cols(sh, aggregated)

    print(f"\n[5c] Updating Overall tab pct columns (BM–CL)...")
    update_overall_pct_cols(sh, aggregated)

    print(f"\n[6/6] Writing Email_Data tab for n8n pickup...")
    write_email_data_tab(sh, aggregated, flat_rows)


def print_categorization(flat_rows: list) -> None:
    print(f"\n{'Coupon Code':<30} {'Category':<20} {'Count':>8} {'Disc Amt':>14}")
    print("-" * 76)
    for r in flat_rows:
        cat = categorize(r["coupon_code"])
        print(f"{r['coupon_code']:<30} {cat:<20} {r['times_used']:>8} {r['total_discount_amount']:>14,.2f}")
