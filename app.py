import os
import secrets
from flask import Flask, render_template, request, jsonify, send_file, session, after_this_request
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
import tempfile

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

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

# Column-A labels that are structural dividers, not spendable categories.
# Matched with startswith so partial labels like "EXPENSES BY D.A." are caught.
_SKIP_PREFIXES = ("ATM WITHDRAWALS", "EXPENSES")

BUCKET_LABELS = ["Necessities", "Luxuries", "Future Self"]

# Known category keywords — recognized regardless of formatting, bold, or case.
# Each entry is a tuple of (canonical_name, [aliases...]).
# The canonical name is what gets stored; aliases are what we match against.
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


def _match_known_category(value):
    """Return the canonical category name if the cell value matches a known category,
    otherwise return None. Matching is case-insensitive and strips whitespace."""
    if not value:
        return None
    normalized = str(value).strip().lower()
    # Exact match first
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
    # Rule 1: known category name — no formatting required
    if _match_known_category(cell.value):
        return True
    # Rule 2: bold + solid colored fill (original rule)
    if not (cell.font and cell.font.bold):
        return False
    fill = cell.fill
    if fill.fill_type != "solid":
        return False
    fg = fill.fgColor
    _EXCLUDED_COLORS = {"00000000", "FFFFFFFF", "FFF3F3F3"}
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


def extract_simple(wb):
    """Parse a simple 2-column budget workbook (no month sheets).
    Expects: col A = category name, col B = amount.
    Treats rows containing 'income' as income; skips bucket headers and subtotals."""
    ws = wb.active
    _SKIP_WORDS = {"necessities", "luxuries", "future self", "subtotal", "total", "budget template"}

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

        # Skip structural rows
        if any(name_lower == w or name_lower.startswith(w) for w in _SKIP_WORDS):
            continue

        if not isinstance(amt, (int, float)) or amt <= 0:
            continue

        if "income" in name_lower:
            income = float(amt)
        else:
            categories[name] = round(float(amt), 2)
            ordered_cats.append(name)

    monthly_data = {"BUDGET": {"income": income, **categories}}
    return monthly_data, ordered_cats


# Spend-column keywords for the YYYY Month template (Option B)
_SPEND_KEYWORDS = {"total", "totals", "spent", "spend", "monthly spend", "monthly spent"}


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

        # Read income from the header section (rows 3-10) using the Earned column (col E = 5)
        # In this template, income rows live above the SPENDING header at row 11
        income_total = 0.0
        for row_idx in range(3, 11):
            name_val = ws.cell(row_idx, 1).value
            if not name_val:
                continue
            name_str = str(name_val).strip().upper()
            # Skip section headers and total rows
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

            # Skip structural rows
            if "TOTAL" in name_upper or any(name_upper.startswith(p) for p in _SKIP_PREFIXES):
                continue

            # Only read rows that are category group subtotals (e.g. "1. Spiritual")
            # These are bold or match a known category
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

        if not ordered_cats:
            ordered_cats = cat_order_this_month

        results[month] = month_data

    return results, ordered_cats


def extract_data(filepath):
    """Dynamically detect months and categories from the uploaded workbook."""
    wb = load_workbook(filepath, data_only=True)

    # Build a map of standard month key -> actual sheet name.
    # Handles: plain names (JAN, January), and YYYY Month format (2025 May) for years 2020-2035.
    sheet_name_map = {}
    for s in wb.sheetnames:
        # Try direct alias match first
        standard = _MONTH_ALIASES.get(s.strip().upper())
        # If not found, try stripping a year prefix
        if not standard:
            stripped = _strip_year_prefix(s)
            standard = _MONTH_ALIASES.get(stripped)
        if standard and standard not in sheet_name_map:
            sheet_name_map[standard] = s
    months = [m for m in ALL_MONTHS if m in sheet_name_map]

    if not months:
        return extract_simple(wb)

    # ── Option B: Detect and handle the YYYY Month template as a special case ──
    if _is_yyyy_month_template(wb, sheet_name_map):
        return _extract_yyyy_month(wb, sheet_name_map, months)

    # ── Standard extraction ────────────────────────────────────────────────────
    results = {}
    ordered_cats = []   # category names in sheet order, set from the first month

    for month in months:
        ws = wb[sheet_name_map[month]]
        totals_col = _find_totals_col(ws)
        if totals_col is None:
            continue

        # Scan column A for all category headers (known name or bold+colored)
        headers = []
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=1)
            if _is_category_header(cell):
                # Use canonical name if it matches a known category, else use raw value
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


def compute_averages(monthly_data, selected_months=None):
    """Income is summed across all active months.
    All spending categories are averaged across active months.
    If selected_months is None, defaults to all active months (income > 0)."""
    if selected_months:
        active_months = [m for m in selected_months if m in monthly_data]
    else:
        active_months = [m for m, data in monthly_data.items() if data.get("income", 0) > 0]
    n = len(active_months)
    if n == 0:
        return {}

    keys = list(next(iter(monthly_data.values())).keys())
    result = {}
    for k in keys:
        total = sum(monthly_data[m].get(k, 0) for m in active_months)
        if k == "income":
            # Income is the grand total across all months
            result[k] = round(total, 2)
        else:
            # Spending categories are averaged across months
            result[k] = round(total / n, 2)
    return result


def build_output_xlsx(averages, assignments):
    """
    Build the clean 3-bucket output spreadsheet.
    assignments = { category_name: bucket_label, ... }
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Awareness Engine"

    # ── Styles ────────────────────────────────────────────────────────────────
    bucket_font   = Font(name="Arial", bold=True, size=12)
    label_font    = Font(name="Arial", size=11)
    pct_font      = Font(name="Arial", bold=True, size=12)
    total_font    = Font(name="Arial", bold=True, size=11)

    nec_fill  = PatternFill("solid", start_color="D6EAF8")   # light blue
    lux_fill  = PatternFill("solid", start_color="D5F5E3")   # light green
    fs_fill   = PatternFill("solid", start_color="FEF9E7")   # light yellow
    title_fill = PatternFill("solid", start_color="2C3E50")
    title_font = Font(name="Arial", bold=True, size=14, color="FFFFFF")

    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")
    right  = Alignment(horizontal="right",  vertical="center")

    thin = Side(style="thin", color="AAAAAA")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    def style(cell, font=None, fill=None, align=None, num_fmt=None):
        if font:   cell.font      = font
        if fill:   cell.fill      = fill
        if align:  cell.alignment = align
        if num_fmt: cell.number_format = num_fmt
        cell.border = border

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells("A1:C1")
    ws["A1"] = "THE AWARENESS ENGINE — My Financial Reality"
    style(ws["A1"], font=title_font, fill=title_fill, align=center)
    ws.row_dimensions[1].height = 30

    # ── Income row ────────────────────────────────────────────────────────────
    ws.row_dimensions[2].height = 6   # spacer
    ws["A3"] = "Total Income"
    ws["B3"] = averages.get("income", 0)
    ws["C3"] = "100%"
    style(ws["A3"], font=total_font, align=left)
    style(ws["B3"], font=total_font, align=right, num_fmt='"$"#,##0.00')
    style(ws["C3"], font=pct_font,   align=center)
    ws.row_dimensions[3].height = 22

    ws.row_dimensions[4].height = 8   # spacer

    # ── Bucket colours map ────────────────────────────────────────────────────
    bucket_fill = {
        "Necessities":  nec_fill,
        "Luxuries":     lux_fill,
        "Future Self":  fs_fill,
    }

    # ── Group categories by bucket ────────────────────────────────────────────
    buckets = {b: [] for b in BUCKET_LABELS}
    for cat, bucket in assignments.items():
        if bucket in buckets:
            buckets[bucket].append(cat)

    income = averages.get("income", 0)   # guard division with 'if income' checks below
    total_fill = PatternFill("solid", start_color="2C3E50")
    total_white = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    current_row = 5

    for bucket_name in BUCKET_LABELS:
        cats = buckets[bucket_name]
        fill = bucket_fill[bucket_name]

        # Bucket header
        ws.merge_cells(f"A{current_row}:C{current_row}")
        ws[f"A{current_row}"] = f"▶  {bucket_name.upper()}"
        style(ws[f"A{current_row}"], font=bucket_font, fill=fill, align=left)
        ws.row_dimensions[current_row].height = 24
        current_row += 1

        # Column sub-headers
        ws[f"A{current_row}"] = "Category"
        ws[f"B{current_row}"] = "Avg Monthly ($)"
        ws[f"C{current_row}"] = "% of Income"
        for col in ["A", "B", "C"]:
            style(ws[f"{col}{current_row}"], font=Font(name="Arial", bold=True, size=10),
                  fill=fill, align=center)
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        # Category rows — track which rows hold the $ amounts for SUM formula
        data_start = current_row
        if not cats:
            ws[f"A{current_row}"] = "(No categories assigned)"
            ws[f"B{current_row}"] = 0
            ws[f"C{current_row}"] = "0.0%"
            for col in ["A", "B", "C"]:
                style(ws[f"{col}{current_row}"], font=label_font, fill=fill, align=left)
            current_row += 1
        else:
            for cat in cats:
                amt = averages.get(cat, 0)
                pct = (amt / income) * 100 if income else 0
                ws[f"A{current_row}"] = cat
                ws[f"B{current_row}"] = amt
                ws[f"C{current_row}"] = f"{pct:.1f}%"
                style(ws[f"A{current_row}"], font=label_font, fill=fill, align=left)
                style(ws[f"B{current_row}"], font=label_font, fill=fill,
                      align=right, num_fmt='"$"#,##0.00')
                style(ws[f"C{current_row}"], font=label_font, fill=fill, align=center)
                ws.row_dimensions[current_row].height = 20
                current_row += 1

        data_end = current_row - 1

        # Bucket TOTAL row
        bucket_total_amt = sum(averages.get(c, 0) for c in cats)
        bucket_pct = (bucket_total_amt / income) * 100 if income else 0
        ws[f"A{current_row}"] = f"TOTAL — {bucket_name}"
        ws[f"B{current_row}"] = f"=SUM(B{data_start}:B{data_end})"
        ws[f"C{current_row}"] = f"{bucket_pct:.1f}%"
        for col in ["A", "B", "C"]:
            style(ws[f"{col}{current_row}"], font=total_white, fill=total_fill, align=center)
        ws[f"B{current_row}"].number_format = '"$"#,##0.00'
        ws.row_dimensions[current_row].height = 22
        current_row += 2   # spacer after each bucket

    # ── Grand summary rows ────────────────────────────────────────────────────
    ws.column_dimensions["D"].width = 16   # % of Income column

    current_row += 1
    ws.merge_cells(f"A{current_row}:D{current_row}")
    ws[f"A{current_row}"] = "─── GRAND REVEAL ───"
    style(ws[f"A{current_row}"], font=Font(name="Arial", bold=True, size=12,
          color="FFFFFF"), fill=PatternFill("solid", start_color="1A5276"), align=center)
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    # Grand Reveal column headers
    reveal_fill = PatternFill("solid", start_color="1A5276")
    for col, label in [("A", "Bucket"), ("B", "Amount ($)"),
                        ("C", "% of Spending"), ("D", "% of Income")]:
        ws[f"{col}{current_row}"] = label
        style(ws[f"{col}{current_row}"], fill=reveal_fill, align=center)
        ws[f"{col}{current_row}"].font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.row_dimensions[current_row].height = 18
    current_row += 1

    left_out = {cat for cat, b in assignments.items() if b == "Leave Out"}
    all_spent = sum(v for k, v in averages.items() if k != "income" and k not in left_out)

    for bucket_name in BUCKET_LABELS:
        cats = buckets[bucket_name]
        bucket_total = sum(averages.get(c, 0) for c in cats)
        pct_spending = (bucket_total / all_spent) * 100 if all_spent else 0
        pct_income   = (bucket_total / income)    * 100 if income    else 0
        ws[f"A{current_row}"] = bucket_name.upper()
        ws[f"B{current_row}"] = bucket_total
        ws[f"C{current_row}"] = f"{pct_spending:.1f}%"
        ws[f"D{current_row}"] = f"{pct_income:.1f}%"
        for col in ["A", "B", "C", "D"]:
            style(ws[f"{col}{current_row}"],
                  font=Font(name="Arial", bold=True, size=11), align=center)
        ws[f"B{current_row}"].number_format = '"$"#,##0.00'
        ws.row_dimensions[current_row].height = 22
        current_row += 1

    # Total spending summary row
    ws[f"A{current_row}"] = "TOTAL SPENDING"
    ws[f"B{current_row}"] = all_spent
    ws[f"C{current_row}"] = "100.0%"
    ws[f"D{current_row}"] = f"{(all_spent / income * 100):.1f}%" if income else "0.0%"
    dark_fill = PatternFill("solid", start_color="2C3E50")
    white_bold = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    for col in ["A", "B", "C", "D"]:
        style(ws[f"{col}{current_row}"], font=white_bold, fill=dark_fill, align=center)
    ws[f"B{current_row}"].number_format = '"$"#,##0.00'
    ws.row_dimensions[current_row].height = 22

    # ── Save to temp file ─────────────────────────────────────────────────────
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    return tmp.name


def build_table_data(averages, assignments):
    """Build structured table data for frontend rendering."""
    income = averages.get("income", 0)
    left_out = {cat for cat, b in assignments.items() if b == "Leave Out"}

    buckets_out = []
    for bucket_name in BUCKET_LABELS:
        cats = [cat for cat, b in assignments.items() if b == bucket_name]
        categories = [{"name": cat, "amt": round(averages.get(cat, 0), 2)} for cat in cats]
        bucket_total = round(sum(c["amt"] for c in categories), 2)
        pct_income = round((bucket_total / income * 100), 1) if income else 0.0
        buckets_out.append({
            "name": bucket_name,
            "categories": categories,
            "total": bucket_total,
            "pct_income": pct_income
        })

    all_spent = round(sum(v for k, v in averages.items()
                          if k != "income" and k not in left_out), 2)

    grand_reveal = []
    for b in buckets_out:
        grand_reveal.append({
            "name": b["name"],
            "total": b["total"],
            "pct_spending": round((b["total"] / all_spent * 100), 1) if all_spent else 0.0,
            "pct_income": b["pct_income"]
        })

    return {
        "income": round(income, 2),
        "buckets": buckets_out,
        "grand_reveal": grand_reveal,
        "total_spending": all_spent,
        "total_pct_income": round((all_spent / income * 100), 1) if income else 0.0
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Please upload a .xlsx file"}), 400

    # Save upload to temp
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

    try:
        f.save(tmp.name)
        monthly_data, category_names = extract_data(tmp.name)
        session["monthly_data"] = monthly_data
        averages = compute_averages(monthly_data)

        categories = [{"name": n, "amount": averages.get(n, 0)} for n in category_names]

        return jsonify({
            "income": averages.get("income", 0),
            "categories": categories,
            "months_found": list(monthly_data.keys())
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp.name)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    assignments = data.get("assignments", {})  # { cat_name: bucket_label }
    selected_months = data.get("selected_months", None)

    monthly_data = session.get("monthly_data")
    if not monthly_data:
        return jsonify({"error": "No data in session. Please upload your file again."}), 400

    averages = compute_averages(monthly_data, selected_months)
    if not averages:
        return jsonify({"error": "No data found for the selected months."}), 400

    try:
        output_path = build_output_xlsx(averages, assignments)
        session["output_path"] = output_path
        session["assignments"] = assignments

        # Build month-by-month breakdown for trend view
        active_months = selected_months if selected_months else \
            [m for m, d in monthly_data.items() if d.get("income", 0) > 0]
        monthly_breakdown = {
            cat: {m: round(monthly_data[m].get(cat, 0), 2) for m in active_months}
            for cat, bucket in assignments.items()
            if bucket not in ("Leave Out",) and bucket in BUCKET_LABELS
        }

        return jsonify({
            "success": True,
            "table": build_table_data(averages, assignments),
            "monthly_breakdown": monthly_breakdown,
            "trend_months": active_months
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/save", methods=["POST"])
def save():
    data = request.get_json()
    income = float(data.get("income", 0))
    buckets_input = data.get("buckets", {})   # { bucket_name: [{name, amt}, ...] }

    averages = {"income": income}
    assignments = {}
    for bucket_name, cats in buckets_input.items():
        for cat in cats:
            averages[cat["name"]] = float(cat["amt"])
            assignments[cat["name"]] = bucket_name

    try:
        old_path = session.get("output_path")
        if old_path and os.path.exists(old_path):
            os.unlink(old_path)

        output_path = build_output_xlsx(averages, assignments)
        session["output_path"] = output_path
        session["averages"] = averages
        session["assignments"] = assignments
        return jsonify({"success": True, "table": build_table_data(averages, assignments)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download")
def download():
    output_path = session.get("output_path")
    if not output_path or not os.path.exists(output_path):
        return "No file ready. Please generate first.", 400

    @after_this_request
    def cleanup(response):
        try:
            os.unlink(output_path)
            session.pop("output_path", None)
        except Exception:
            pass
        return response

    return send_file(output_path, as_attachment=True,
                     download_name="Awareness_Engine_Results.xlsx")


@app.route("/template/download", methods=["POST"])
def template_download():
    data = request.get_json(force=True)
    income = float(data.get("income", 0))
    buckets = data.get("buckets", {})

    wb = Workbook()
    ws = wb.active
    ws.title = "Budget Template"

    header_font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    bucket_fonts = {
        "Necessities":  Font(name="Calibri", bold=True, size=12, color="1A6380"),
        "Luxuries":     Font(name="Calibri", bold=True, size=12, color="2E7D52"),
        "Future Self":  Font(name="Calibri", bold=True, size=12, color="A0522D"),
    }
    bucket_fills = {
        "Necessities":  PatternFill("solid", fgColor="DBF0F7"),
        "Luxuries":     PatternFill("solid", fgColor="DAF2E5"),
        "Future Self":  PatternFill("solid", fgColor="FEF0E0"),
    }
    header_fill = PatternFill("solid", fgColor="0F1F3D")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    # Header row
    ws.merge_cells("A1:B1")
    ws["A1"] = "Budget Template"
    ws["A1"].font = header_font
    ws["A1"].fill = header_fill
    ws["A1"].alignment = center

    # Income row
    ws["A2"] = "Monthly Income"
    ws["A2"].font = Font(name="Calibri", bold=True, size=12)
    ws["B2"] = income
    ws["B2"].number_format = '"$"#,##0.00'
    ws["B2"].alignment = right_align
    ws["A2"].border = border
    ws["B2"].border = border

    row = 4
    bucket_totals = {}
    for bucket_name in BUCKET_LABELS:
        cats = buckets.get(bucket_name, [])
        # Bucket header
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"] = bucket_name.upper()
        ws[f"A{row}"].font = bucket_fonts[bucket_name]
        ws[f"A{row}"].fill = bucket_fills[bucket_name]
        ws[f"A{row}"].alignment = center
        ws[f"A{row}"].border = border
        row += 1

        total = 0
        for item in cats:
            name = item.get("name", "")
            amt = float(item.get("amount", 0))
            total += amt
            ws[f"A{row}"] = name
            ws[f"B{row}"] = amt
            ws[f"B{row}"].number_format = '"$"#,##0.00'
            ws[f"A{row}"].border = border
            ws[f"B{row}"].border = border
            ws[f"B{row}"].alignment = right_align
            row += 1

        bucket_totals[bucket_name] = total

        # Subtotal
        ws[f"A{row}"] = "Subtotal"
        ws[f"A{row}"].font = Font(name="Calibri", bold=True, size=11)
        ws[f"B{row}"] = total
        ws[f"B{row}"].number_format = '"$"#,##0.00'
        ws[f"B{row}"].font = Font(name="Calibri", bold=True, size=11)
        ws[f"A{row}"].border = border
        ws[f"B{row}"].border = border
        ws[f"B{row}"].alignment = right_align
        row += 2

    # Grand Reveal
    row += 1  # spacer
    reveal_fill  = PatternFill("solid", fgColor="0F1F3D")
    reveal_font  = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    sub_hdr_fill = PatternFill("solid", fgColor="F4F4F1")
    sub_hdr_font = Font(name="Calibri", bold=True, size=11)
    total_fill   = PatternFill("solid", fgColor="2C3E50")
    total_font_w = Font(name="Calibri", bold=True, size=11, color="FFFFFF")

    # Merge across 4 columns for Grand Reveal — widen sheet first
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"] = "⚡ GRAND REVEAL"
    ws[f"A{row}"].font = reveal_font
    ws[f"A{row}"].fill = reveal_fill
    ws[f"A{row}"].alignment = center
    ws[f"A{row}"].border = border
    row += 1

    # Column headers
    for col, label in [("A", "Bucket"), ("B", "Amount"), ("C", "% of Income"), ("D", "% of Spending")]:
        ws[f"{col}{row}"] = label
        ws[f"{col}{row}"].font = sub_hdr_font
        ws[f"{col}{row}"].fill = sub_hdr_fill
        ws[f"{col}{row}"].alignment = center
        ws[f"{col}{row}"].border = border
    row += 1

    total_spending = sum(bucket_totals.values())

    for bucket_name, b_fill in bucket_fills.items():
        amt = bucket_totals[bucket_name]
        pct_inc = round((amt / income * 100), 1) if income else 0.0
        pct_spd = round((amt / total_spending * 100), 1) if total_spending else 0.0
        ws[f"A{row}"] = bucket_name
        ws[f"B{row}"] = amt
        ws[f"C{row}"] = f"{pct_inc:.1f}%"
        ws[f"D{row}"] = f"{pct_spd:.1f}%"
        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{row}"].fill = b_fill
            ws[f"{col}{row}"].font = Font(name="Calibri", bold=True, size=11)
            ws[f"{col}{row}"].alignment = center
            ws[f"{col}{row}"].border = border
        ws[f"B{row}"].number_format = '"$"#,##0.00'
        row += 1

    # Total row
    total_pct_inc = round((total_spending / income * 100), 1) if income else 0.0
    ws[f"A{row}"] = "TOTAL EXPENSES"
    ws[f"B{row}"] = total_spending
    ws[f"C{row}"] = f"{total_pct_inc:.1f}%"
    ws[f"D{row}"] = "100.0%"
    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{row}"].font = total_font_w
        ws[f"{col}{row}"].fill = total_fill
        ws[f"{col}{row}"].alignment = center
        ws[f"{col}{row}"].border = border
    ws[f"B{row}"].number_format = '"$"#,##0.00'

    # Remaining Income row
    row += 1
    remaining = income - total_spending
    remaining_pct = round((remaining / income * 100), 1) if income else 0.0
    if remaining >= 0:
        rem_fill = PatternFill("solid", fgColor="D5F5E3")
        rem_font = Font(name="Calibri", bold=True, size=11, color="145A32")
    else:
        rem_fill = PatternFill("solid", fgColor="FDECEA")
        rem_font = Font(name="Calibri", bold=True, size=11, color="C0392B")
    ws[f"A{row}"] = "Remaining Income"
    ws[f"B{row}"] = remaining
    ws[f"C{row}"] = f"{remaining_pct:.1f}%"
    ws[f"D{row}"] = "—"
    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{row}"].fill = rem_fill
        ws[f"{col}{row}"].font = rem_font
        ws[f"{col}{row}"].alignment = center
        ws[f"{col}{row}"].border = border
    ws[f"B{row}"].number_format = '"$"#,##0.00'

    # Column widths
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()
    tmp_path = tmp.name

    @after_this_request
    def cleanup_tmpl(response):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return response

    return send_file(tmp_path, as_attachment=True,
                     download_name="Budget_Template.xlsx")


if __name__ == "__main__":
    app.run(debug=True, port=5050)
