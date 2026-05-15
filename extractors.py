"""
extractors.py — Spreadsheet parsing logic.

Responsible for reading uploaded .xlsx files and returning structured
monthly data and ordered category lists. Supports three template formats:
  - Simple 2-column (no month sheets)
  - YYYY Month (e.g. "2025 May")
  - Daily tracking (date columns + TOTALS column)
"""

import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ── Month constants ────────────────────────────────────────────────────────────

ALL_MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
              "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# Maps any variation of a month name (short or full, any case) -> standard key
_MONTH_ALIASES = {
    "JAN": "JAN", "JANUARY": "JAN",
    "FEB": "FEB", "FEBRUARY": "FEB",
    "MAR": "MAR", "MARCH": "MAR",
    "APR": "APR", "APRIL": "APR",
    "MAY": "MAY",
    "JUN": "JUN", "JUNE": "JUN",
    "JUL": "JUL", "JULY": "JUL",
    "AUG": "AUG", "AUGUST": "AUG",
    "SEP": "SEP", "SEPT": "SEP", "SEPTEMBER": "SEP",
    "OCT": "OCT", "OCTOBER": "OCT",
    "NOV": "NOV", "NOVEMBER": "NOV",
    "DEC": "DEC", "DECEMBER": "DEC",
}

# ── Bucket labels ──────────────────────────────────────────────────────────────

BUCKET_LABELS = ["Necessities", "Luxuries", "Future Self"]

# ── Skip / structural label sets ──────────────────────────────────────────────

# Column-A labels that are structural dividers, not spendable categories.
# Matched with startswith so partial labels like "EXPENSES BY D.A." are caught.
_SKIP_PREFIXES = ("ATM WITHDRAWALS", "EXPENSES")

# Structural header labels that should never be treated as expense categories.
_STRUCTURAL_HEADERS = frozenset({
    "INCOME",
    "TOTAL INCOME BEFORE TAXES",
    "EXPENSES BY D.A. CATEGORIES",
    "TOTAL EXPENSES",
})

_EXCLUDED_COLORS = {"00000000", "FFFFFFFF", "FFF3F3F3"}

# Row labels in the simple template that are structural dividers, not spendable categories.
_SKIP_WORDS = {"necessities", "luxuries", "future self", "subtotal", "total", "budget template"}

# Spend-column keywords for the YYYY Month template
_SPEND_KEYWORDS = {"total", "totals", "spent", "spend", "monthly spend", "monthly spent"}

# ── Known category lookup ──────────────────────────────────────────────────────

# Known category keywords — recognized regardless of formatting, bold, or case.
# Each entry is a tuple of (canonical_name, [aliases...]).
_KNOWN_CATEGORIES = [
    ("Spiritual",         ["spiritual"]),
    ("Shelter",           ["shelter"]),
    ("Utilities",         ["utilities"]),
    ("Internet",          ["internet"]),
    ("Food",              ["food"]),
    ("Home Items",        ["home items", "home item"]),
    ("Groceries",         ["groceries", "grocery"]),
    ("Transportation",    ["transportation"]),
    ("Phone Bill",        ["phone bill"]),
    ("Laundry",           ["laundry"]),
    ("Storage",           ["storage"]),
    ("Moving",            ["moving"]),
    ("Gym",               ["gym"]),
    ("Clothing",          ["clothing"]),
    ("Personal Care",     ["personal care"]),
    ("Health Care",       ["health care", "healthcare"]),
    ("Inner Child",       ["inner child"]),
    ("Entertainment",     ["entertainment"]),
    ("Education",         ["education"]),
    ("Vacation",          ["vacation", "vacations"]),
    ("Personal Business", ["personal business"]),
    ("Gifts",             ["gifts"]),
    ("Investments",       ["investments", "investment"]),
    ("Taxes",             ["taxes"]),
    ("Debt Repayment",    ["debt repayment"]),
    ("Prudent Reserve",   ["prudent reserve"]),
    ("Subscriptions",     ["subscriptions", "subscription"]),
    ("Dining Out",        ["dining out", "eating out", "dine out"]),
    ("Take Out",          ["take out", "takeout", "takeaway", "to go"]),
]

# Flat lookup: lowercase alias -> canonical name
_KNOWN_CATEGORY_LOOKUP = {
    alias: canonical
    for canonical, aliases in _KNOWN_CATEGORIES
    for alias in aliases
}


# ── Cell helpers ───────────────────────────────────────────────────────────────

def _match_known_category(value):
    """Return the canonical category name if the cell value matches a known category,
    otherwise return None. Matching is case-insensitive and strips whitespace."""
    if not value:
        return None
    normalized = str(value).strip().lower()
    if normalized in _KNOWN_CATEGORY_LOOKUP:
        return _KNOWN_CATEGORY_LOOKUP[normalized]
    # Partial match — cell value starts with a known alias (handles "Groceries/Grocery" etc.)
    for alias, canonical in _KNOWN_CATEGORY_LOOKUP.items():
        if normalized.startswith(alias):
            return canonical
    return None


def _find_totals_col(ws):
    """Return the column index of the TOTALS header in row 1, or None.
    Matches 'Total', 'TOTALS', 'total', etc. case-insensitively."""
    for cell in ws[1]:
        if cell.value and str(cell.value).strip().upper().startswith("TOTAL"):
            return cell.column
    return None


def _is_category_header(cell):
    """True if the cell looks like a top-level category header.
    Matches either:
      1. A known category name (case-insensitive, no formatting required), or
      2. Bold + distinctly colored fill (original formatting-based rule).
    Excludes white and near-white fills used for alternating sub-item rows."""
    if not cell.value:
        return False
    if _match_known_category(cell.value):
        return True
    if not (cell.font and cell.font.bold):
        return False
    fill = cell.fill
    if fill.fill_type != "solid":
        return False
    fg = fill.fgColor
    if fg.type == "theme":
        return True
    return fg.rgb not in _EXCLUDED_COLORS


def _is_bold_colored_header(cell):
    """True if a cell is a bold + distinctly colored header (marks end of a section)."""
    if not cell.value:
        return False
    if not (cell.font and cell.font.bold):
        return False
    fill = cell.fill
    if fill.fill_type != "solid":
        return False
    fg = fill.fgColor
    if fg.type == "theme":
        return True
    return fg.rgb not in _EXCLUDED_COLORS


def _sum_col(ws, row_start, row_end, col):
    """Sum numeric values in a single column over a row range."""
    total = 0.0
    for (val,) in ws.iter_rows(min_row=row_start, max_row=row_end,
                                min_col=col, max_col=col, values_only=True):
        if isinstance(val, (int, float)):
            total += val
    return round(total, 2)


def _read_row_total(ws, row_idx, totals_col):
    """Dual-calculation: sum raw daily cols ourselves, also read the TOTALS col.
    If both match within ±0.02, use the TOTALS col value.
    If they differ, trust our own calculation."""
    own_sum = 0.0
    for col_idx in range(2, totals_col):
        val = ws.cell(row_idx, col_idx).value
        if isinstance(val, (int, float)):
            own_sum += val
    own_sum = round(own_sum, 2)

    totals_val = ws.cell(row_idx, totals_col).value
    sheet_total = round(float(totals_val), 2) if isinstance(totals_val, (int, float)) else 0.0

    if abs(own_sum - sheet_total) <= 0.02:
        return sheet_total
    return own_sum


# ── Template detection ─────────────────────────────────────────────────────────

def _strip_year_prefix(name):
    """Strip a leading year (2020-2035) from a sheet name, e.g. '2025 May' -> 'MAY'.
    Returns the uppercased remainder, or the original uppercased name if no year found."""
    match = re.match(r'^(20[2-3][0-9])\s+(.+)$', name.strip())
    if match:
        return match.group(2).strip().upper()
    return name.strip().upper()


def _is_yyyy_month_template(wb, sheet_name_map):
    """Detect if this workbook uses the 'YYYY Month' style template.
    Requires ALL three conditions to be true to avoid false positives:
      1. At least one sheet name had a year prefix (2020-2035) stripped
      2. Row 11 col A contains 'SPENDING'
      3. Row 11 contains a spend-keyword column header"""
    has_year_prefix = any(
        re.match(r'^(20[2-3][0-9])\s+', actual.strip())
        for actual in sheet_name_map.values()
    )
    if not has_year_prefix:
        return False

    for standard, actual in sheet_name_map.items():
        ws = wb[actual]
        cell_a11 = ws.cell(11, 1).value
        if cell_a11 and str(cell_a11).strip().upper() == "SPENDING":
            for c in range(1, ws.max_column + 1):
                val = ws.cell(11, c).value
                if val and str(val).strip().lower() in _SPEND_KEYWORDS:
                    return True
    return False


def _is_daily_tracking_template(wb, sheet_name_map):
    """Detect if this workbook uses the daily-tracking format.
    Signature: plain month-name sheets (JAN/FEB etc.), row 1 has a date in col C,
    and a TOTALS column header exists in row 1."""
    for standard, actual in sheet_name_map.items():
        ws = wb[actual]
        row1_c = ws.cell(1, 3).value
        has_date = isinstance(row1_c, datetime) or (
            isinstance(row1_c, str) and re.match(r'^\d{4}-\d{2}-\d{2}', row1_c.strip())
        )
        if has_date and _find_totals_col(ws) is not None:
            return True
    return False


# ── Extraction functions ───────────────────────────────────────────────────────

def extract_simple(wb):
    """Parse a simple 2-column budget workbook (no month sheets).
    Expects: col A = category name, col B = amount.
    Treats rows containing 'income' as income; skips bucket headers and subtotals."""
    ws = wb.active
    income = 0.0
    categories = {}
    ordered_cats = []

    for row in ws.iter_rows(min_row=1, values_only=True):
        name = row[0]
        amt  = row[1] if len(row) > 1 else None

        if not name or not isinstance(name, str):
            continue
        name = name.strip()
        if not name:
            continue

        name_lower = name.lower()
        if any(name_lower == w or name_lower.startswith(w) for w in _SKIP_WORDS):
            continue
        if not isinstance(amt, (int, float)):
            continue

        if "income" in name_lower:
            income = float(amt)
        else:
            categories[name] = round(float(amt), 2)
            ordered_cats.append(name)

    monthly_data = {"BUDGET": {"income": income, **categories}}
    return monthly_data, ordered_cats


def _find_spend_col(ws):
    """For the YYYY Month template: find the column in row 11 matching a spend keyword.
    Returns (col_index, data_start_row) or (None, None)."""
    for c in range(1, ws.max_column + 1):
        val = ws.cell(11, c).value
        if val and str(val).strip().lower() in _SPEND_KEYWORDS:
            return c, 12  # data starts at row 12
    return None, None


def _extract_yyyy_month(wb, sheet_name_map, months):
    """Extract data from the YYYY Month style template.
    Reads income from rows 3-10 (Earned col) and category subtotals from rows 12+."""
    results = {}
    ordered_cats = []

    for month in months:
        ws = wb[sheet_name_map[month]]
        spend_col, data_start = _find_spend_col(ws)
        if spend_col is None:
            continue

        month_data = {"income": 0.0}
        cat_order_this_month = []

        # Read income from the header section (rows 3-10)
        # Income rows live above the SPENDING header at row 11
        income_total = 0.0
        for row_idx in range(3, 11):
            name_val = ws.cell(row_idx, 1).value
            if not name_val:
                continue
            name_str = str(name_val).strip().upper()
            if name_str in ("INCOME", "TOTAL INCOME") or "TOTAL" in name_str:
                continue
            amt_val = ws.cell(row_idx, spend_col).value
            if isinstance(amt_val, (int, float)) and amt_val > 0:
                income_total += amt_val
        month_data["income"] = round(income_total, 2)

        for row_idx in range(data_start, ws.max_row + 1):
            cell_a = ws.cell(row_idx, 1)
            if not cell_a.value:
                continue
            name_raw = str(cell_a.value).strip()
            if not name_raw:
                continue

            name_upper = name_raw.upper()
            if "TOTAL" in name_upper or any(name_upper.startswith(p) for p in _SKIP_PREFIXES):
                continue

            # Only read rows that are category group subtotals (bold or known category)
            canonical = _match_known_category(name_raw)
            if not canonical and not (cell_a.font and cell_a.font.bold):
                continue

            name = canonical if canonical else name_raw
            amt_val = ws.cell(row_idx, spend_col).value
            amt = round(float(amt_val), 2) if isinstance(amt_val, (int, float)) else 0.0

            if "INCOME" in name_upper:
                month_data["income"] = amt
            else:
                month_data[name] = amt
                if name not in cat_order_this_month:
                    cat_order_this_month.append(name)

        for name in cat_order_this_month:
            if name not in ordered_cats:
                ordered_cats.append(name)

        results[month] = month_data

    return results, ordered_cats


def _extract_daily_tracking(wb, sheet_name_map, months):
    """Extract data from the daily-tracking format.
    - Category labels are read from the JAN sheet only (other sheets use =JAN!Axx formulas)
    - Income is always collected individually (sub-rows between income header and next bold+colored header)
    - Expenses: prioritize bold+colored header rows. If none found, fall back to known names
    - Dual-calculation: compares own daily sum vs TOTALS col, trusts own sum if mismatch
    - Handles variable month lengths automatically
    """
    jan_ws = wb[sheet_name_map["JAN"]]
    row_labels = {}
    for row_idx in range(1, jan_ws.max_row + 1):
        val = jan_ws.cell(row_idx, 1).value
        if val and isinstance(val, str) and val.strip():
            row_labels[row_idx] = val.strip()

    if _find_totals_col(jan_ws) is None:
        return {}, []

    # ── Identify income rows ──────────────────────────────────────────────────
    income_rows = set()
    in_income_section = False
    for row_idx in sorted(row_labels.keys()):
        label = row_labels[row_idx]
        cell = jan_ws.cell(row_idx, 1)
        if label.strip().upper() == "INCOME":
            in_income_section = True
            continue
        if in_income_section:
            if _is_bold_colored_header(cell):
                in_income_section = False
            else:
                income_rows.add(row_idx)

    # ── Identify expense rows ─────────────────────────────────────────────────
    bold_colored_expense_rows = set()
    for row_idx, label in row_labels.items():
        if row_idx in income_rows:
            continue
        cell = jan_ws.cell(row_idx, 1)
        if _is_bold_colored_header(cell):
            label_upper = label.upper().strip()
            if any(label_upper.startswith(p) for p in _SKIP_PREFIXES):
                continue
            if label_upper in _STRUCTURAL_HEADERS:
                continue
            bold_colored_expense_rows.add(row_idx)

    use_bold_colored = len(bold_colored_expense_rows) > 0

    results = {}
    ordered_cats = []

    for month in months:
        ws = wb[sheet_name_map[month]]
        totals_col = _find_totals_col(ws)
        if totals_col is None:
            continue

        month_data = {"income": 0.0}
        cat_order_this_month = []

        for row_idx, label in row_labels.items():
            label_upper = label.upper().strip()

            if any(label_upper.startswith(p) for p in _SKIP_PREFIXES):
                continue
            if label_upper in _STRUCTURAL_HEADERS:
                continue

            if row_idx in income_rows:
                total = _read_row_total(ws, row_idx, totals_col)
                if total == 0.0:
                    continue
                month_data["income"] = round(month_data["income"] + total, 2)
                continue

            if use_bold_colored:
                if row_idx not in bold_colored_expense_rows:
                    continue
            else:
                if not _match_known_category(label):
                    continue

            total = _read_row_total(ws, row_idx, totals_col)
            if total == 0.0:
                continue

            canonical = _match_known_category(label)
            name = canonical if canonical else label
            month_data[name] = round(month_data.get(name, 0.0) + total, 2)
            if name not in cat_order_this_month:
                cat_order_this_month.append(name)

        for name in cat_order_this_month:
            if name not in ordered_cats:
                ordered_cats.append(name)

        results[month] = month_data

    return results, ordered_cats


def extract_data(filepath):
    """Dynamically detect months and categories from the uploaded workbook."""
    wb = load_workbook(filepath, data_only=True)

    sheet_name_map = {}
    for s in wb.sheetnames:
        standard = _MONTH_ALIASES.get(s.strip().upper())
        if not standard:
            stripped = _strip_year_prefix(s)
            standard = _MONTH_ALIASES.get(stripped)
        if standard and standard not in sheet_name_map:
            sheet_name_map[standard] = s
    months = [m for m in ALL_MONTHS if m in sheet_name_map]

    if not months:
        return extract_simple(wb)

    if _is_yyyy_month_template(wb, sheet_name_map):
        return _extract_yyyy_month(wb, sheet_name_map, months)

    if _is_daily_tracking_template(wb, sheet_name_map):
        return _extract_daily_tracking(wb, sheet_name_map, months)

    # ── Standard extraction ────────────────────────────────────────────────────
    results = {}
    ordered_cats = []

    for month in months:
        ws = wb[sheet_name_map[month]]
        totals_col = _find_totals_col(ws)
        if totals_col is None:
            continue

        headers = []
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=1)
            if _is_category_header(cell):
                canonical = _match_known_category(cell.value)
                name = canonical if canonical else str(cell.value).strip()
                headers.append((row_idx, name))

        month_data = {"income": 0.0}
        cat_order_this_month = []

        for i, (row_idx, name) in enumerate(headers):
            data_start = row_idx + 1
            data_end = (headers[i + 1][0] - 1) if i + 1 < len(headers) else ws.max_row
            name_upper = name.upper()

            if "TOTAL" in name_upper or any(name_upper.startswith(p) for p in _SKIP_PREFIXES):
                continue

            total = _sum_col(ws, data_start, data_end, totals_col)

            if "INCOME" in name_upper:
                month_data["income"] = total
            else:
                month_data[name] = total
                cat_order_this_month.append(name)

        if not ordered_cats:
            ordered_cats = cat_order_this_month

        results[month] = month_data

    return results, ordered_cats
