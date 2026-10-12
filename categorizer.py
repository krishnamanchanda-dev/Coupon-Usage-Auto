"""
Categorizes Shopify coupon codes into named coupons or buckets for the test tab.

Priority order:
  1. Excluded   — c2p, prepaid, nil, bundle (skip entirely)
  2. Named      — any coupon from the master list (exact normalized match)
  3. PP bucket  — contains: unloq, wishlink, ck; or random alphanumeric code
  4. Influencer — person-name + trailing digits (e.g. SANYA25, PRABHKIRAT25)
  5. Others     — everything else
"""

import re

# ── Master list of all named coupons (from Copy of Overall) ──────────────────
# Key: normalized (lowercase, no spaces). Value: clean display name for sheet header.
NAMED_COUPONS: dict[str, str] = {
    "save30":       "SAVE30",
    "save25":       "SAVE25",
    "save35":       "SAVE35",
    "stealdeal":    "STEALDEAL",
    "buy3":         "BUY3",
    "buy3at699":    "BUY3@699",
    "buy4":         "BUY4",
    "buy5":         "BUY5",
    "buy6":         "BUY6",
    "buy10":        "BUY10",
    "b3p2":         "B3P2",
    "b2g2":         "B2G2",
    "modern20":     "MODERN20",
    "freelipstick": "FREELIPSTICK",
    "freevanity":   "FREEVANITY",
    "luxevanity":   "LUXEVANITY",
    "lips5":        "LIPS5",
    "free5":        "FREE5",
    "free6":        "FREE6",
    "freekit":      "FREEKIT",
    "lipluxe":      "LIPLUXE",
    "5pcs":         "5PCS",
    "3pcs":         "3PCS",
    "4pcs":         "4PCS",
    "special4":     "SPECIAL4",
    "special5":     "SPECIAL5",
    "special6":     "SPECIAL6",
    "gpay":         "GPAY",
    "get4":         "GET4",
    "get2":         "GET2",
    "fullcb":       "FULLCB",
    "luxe3":        "LUXE3",
    "6pcskit":      "6PCS KIT",
    "4pcskit":      "4PCS KIT",
}

# Substrings that identify PP (Paid Promotion) codes
PP_SUBSTRINGS = ["unloq", "wishlink", "ck"]

# Substrings/exact values for exclusion
EXCLUDE_SUBSTRINGS = ["c2p", "prepaid", "bundle", "cashback"]
EXCLUDE_EXACT = {"nil", "creatororder", "customdiscount"}
# Note: "custom discount" normalizes to "customdiscount", "cashback discount" → contains "cashback"

# Normalized key prefixes that indicate a brand/business promo code → Others
# (prevent these from being caught by the influencer pattern)
OTHERS_PREFIXES = ["emp", "swiss", "welcome", "beauty", "sbn", "sb", "bogo"]

# Random-code pattern: 8+ chars, only uppercase letters + digits
_RANDOM_CODE_RE = re.compile(r'^[A-Z0-9]{8,}$')

# Influencer pattern: 3+ letters followed by 2+ digits
_INFLUENCER_RE = re.compile(r'^[A-Za-z]{3,}\d{2,}$')


def _norm(code: str) -> str:
    return code.lower().replace(" ", "").replace("-", "").strip()


def categorize(raw_code: str) -> str:
    """
    Returns one of:
      - a key from NAMED_COUPONS (e.g. 'save30', 'buy3')
      - 'pp'
      - 'influencer'
      - 'others'
      - 'excluded'
    """
    key = _norm(raw_code)
    upper = raw_code.upper().strip()

    # 1. Exclude
    if key in EXCLUDE_EXACT:
        return "excluded"
    for sub in EXCLUDE_SUBSTRINGS:
        if sub in key:
            return "excluded"

    # 2. Named coupon (exact normalized match)
    if key in NAMED_COUPONS:
        return key

    # 3. PP substrings (must come before random check to catch wishlink/unloq combos)
    for sub in PP_SUBSTRINGS:
        if sub in key:
            return "pp"

    # 4. Brand/business promo prefix → Others (before influencer check so EMP50, SWISS25 etc. don't
    #    get caught by the letters+digits influencer pattern)
    for prefix in OTHERS_PREFIXES:
        if key.startswith(prefix):
            return "others"

    # 5. Influencer pattern: word(s) + trailing digits — check BEFORE random alphanumeric
    #    so PRABHKIRAT25, SANYA25, DEVANGI10 are caught here, not as random codes
    if _INFLUENCER_RE.match(raw_code.strip()):
        return "influencer"

    # 5. Random alphanumeric codes (digits interspersed, no clear word pattern) → PP
    if _RANDOM_CODE_RE.match(upper):
        return "pp"

    # 6. Everything else
    return "others"


# Ordered column list for the sheet: all named coupons + 3 buckets
BUCKET_KEYS = ["pp", "influencer", "others"]
BUCKET_DISPLAY = {"pp": "PP", "influencer": "INFLUENCER/MKTNG", "others": "OTHERS"}

# Sorted named coupons in a logical order (matches Copy of Overall flow)
NAMED_ORDER = [
    "save30", "save25", "save35", "stealdeal",
    "buy3", "buy3at699", "buy4", "buy5", "buy6", "buy10",
    "b3p2", "b2g2",
    "modern20", "gpay", "fullcb",
    "freelipstick", "lips5", "free5", "free6",
    "freevanity", "luxevanity", "freekit", "lipluxe",
    "5pcs", "3pcs", "4pcs",
    "special4", "special5", "special6",
    "get4", "get2", "luxe3",
    "6pcskit", "4pcskit",
]

ALL_COLUMNS = NAMED_ORDER + BUCKET_KEYS
