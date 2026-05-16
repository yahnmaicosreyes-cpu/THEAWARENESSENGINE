"""
excel_builders.py — Excel file generation.

Builds all downloadable .xlsx files:
  - build_output_xlsx()        → the Awareness Engine results report
  - build_single_month_template() → single-month budget template
  - build_three_month_template()  → 3-month average budget template
"""

import tempfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from extractors import BUCKET_LABELS


# ── Shared style helper ────────────────────────────────────────────────────────

def _make_budget_template_styles():
    """Return shared styles used by both single-month and 3-month budget template builders."""
    thin = Side(style="thin", color="CCCCCC")
    return {
        "header_font": Font(name="Calibri", bold=True, size=13, color="FFFFFF"),
        "bucket_fonts": {
            "Necessities":  Font(name="Calibri", bold=True, size=12, color="1A6380"),
            "Luxuries":     Font(name="Calibri", bold=True, size=12, color="2E7D52"),
            "Future Self":  Font(name="Calibri", bold=True, size=12, color="A0522D"),
        },
        "bucket_fills": {
            "Necessities":  PatternFill("solid", fgColor="DBF0F7"),
            "Luxuries":     PatternFill("solid", fgColor="DAF2E5"),
            "Future Self":  PatternFill("solid", fgColor="FEF0E0"),
        },
        "header_fill":  PatternFill("solid", fgColor="0F1F3D"),
        "border":       Border(left=thin, right=thin, top=thin, bottom=thin),
        "center":       Alignment(horizontal="center", vertical="center"),
        "right_align":  Alignment(horizontal="right",  vertical="center"),
    }


# ── Awareness Engine output report ─────────────────────────────────────────────

def _render_bucket_sections(ws, averages, assignments, totals, buckets, income,
                             style_fn, bucket_fill, label_font, bucket_font,
                             center, right, left, border, start_row, all_buckets=None):
    """Render all bucket sections into ws. Returns the next available row.
    all_buckets controls order; defaults to BUCKET_LABELS for backwards compat."""
    if all_buckets is None:
        all_buckets = BUCKET_LABELS
    total_fill  = PatternFill("solid", start_color="2C3E50")
    total_white = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    current_row = start_row

    for bucket_name in all_buckets:
        cats = buckets[bucket_name]
        fill = bucket_fill[bucket_name]

        # Bucket header
        ws.merge_cells(f"A{current_row}:D{current_row}")
        ws[f"A{current_row}"] = f"▶  {bucket_name.upper()}"
        style_fn(ws[f"A{current_row}"], font=bucket_font, fill=fill, align=left)
        ws.row_dimensions[current_row].height = 24
        current_row += 1

        # Column sub-headers
        ws[f"A{current_row}"] = "Category"
        ws[f"B{current_row}"] = "Total Spending (all months)"
        ws[f"C{current_row}"] = "Avg Monthly Spending"
        ws[f"D{current_row}"] = "% of Income"
        for col in ["A", "B", "C", "D"]:
            style_fn(ws[f"{col}{current_row}"], font=Font(name="Arial", bold=True, size=10),
                     fill=fill, align=center)
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        # Category rows
        data_start = current_row
        if not cats:
            ws[f"A{current_row}"] = "(No categories assigned)"
            ws[f"B{current_row}"] = 0
            ws[f"C{current_row}"] = 0
            ws[f"D{current_row}"] = "0.0%"
            for col in ["A", "B", "C", "D"]:
                style_fn(ws[f"{col}{current_row}"], font=label_font, fill=fill, align=left)
            current_row += 1
        else:
            for cat in cats:
                amt = averages.get(cat, 0)
                raw = totals.get(cat, 0)
                pct = (amt / income) * 100 if income else 0
                ws[f"A{current_row}"] = cat
                ws[f"B{current_row}"] = raw
                ws[f"C{current_row}"] = amt
                ws[f"D{current_row}"] = f"{pct:.1f}%"
                style_fn(ws[f"A{current_row}"], font=label_font, fill=fill, align=left)
                style_fn(ws[f"B{current_row}"], font=label_font, fill=fill,
                         align=right, num_fmt='"$"#,##0.00')
                style_fn(ws[f"C{current_row}"], font=label_font, fill=fill,
                         align=right, num_fmt='"$"#,##0.00')
                style_fn(ws[f"D{current_row}"], font=label_font, fill=fill, align=center)
                ws.row_dimensions[current_row].height = 20
                current_row += 1

        data_end = current_row - 1

        # Bucket TOTAL row
        bucket_total_amt = sum(averages.get(c, 0) for c in cats)
        bucket_pct = (bucket_total_amt / income) * 100 if income else 0
        ws[f"A{current_row}"] = f"TOTAL — {bucket_name}"
        ws[f"B{current_row}"] = f"=SUM(B{data_start}:B{data_end})"
        ws[f"C{current_row}"] = f"=SUM(C{data_start}:C{data_end})"
        ws[f"D{current_row}"] = f"{bucket_pct:.1f}%"
        for col in ["A", "B", "C", "D"]:
            style_fn(ws[f"{col}{current_row}"], font=total_white, fill=total_fill, align=center)
        ws[f"B{current_row}"].number_format = '"$"#,##0.00'
        ws[f"C{current_row}"].number_format = '"$"#,##0.00'
        ws.row_dimensions[current_row].height = 22
        current_row += 2   # spacer after each bucket

    return current_row


def _render_grand_reveal(ws, averages, assignments, totals, buckets, income,
                          style_fn, center, right, border, start_row, all_buckets=None):
    """Render the Grand Reveal summary section into ws.
    all_buckets controls order; defaults to BUCKET_LABELS for backwards compat."""
    if all_buckets is None:
        all_buckets = BUCKET_LABELS
    current_row = start_row + 1   # one spacer before the header

    reveal_fill = PatternFill("solid", start_color="1A5276")
    dark_fill   = PatternFill("solid", start_color="2C3E50")
    white_bold  = Font(name="Arial", bold=True, size=11, color="FFFFFF")

    # Grand Reveal header
    ws.merge_cells(f"A{current_row}:E{current_row}")
    ws[f"A{current_row}"] = "─── GRAND REVEAL ───"
    style_fn(ws[f"A{current_row}"], font=Font(name="Arial", bold=True, size=12, color="FFFFFF"),
             fill=reveal_fill, align=center)
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    # Column headers
    for col, label in [("A", "Bucket"), ("B", "Total Spending (all months)"),
                        ("C", "Avg Monthly Spending"), ("D", "% of Spending"), ("E", "% of Income")]:
        ws[f"{col}{current_row}"] = label
        style_fn(ws[f"{col}{current_row}"], fill=reveal_fill, align=center)
        ws[f"{col}{current_row}"].font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.row_dimensions[current_row].height = 18
    current_row += 1

    left_out      = {cat for cat, b in assignments.items() if b == "Leave Out"}
    all_spent     = sum(v for k, v in averages.items() if k != "income" and k not in left_out)
    all_raw_spent = sum(totals.get(k, 0) for k in averages if k != "income" and k not in left_out)

    for bucket_name in all_buckets:
        cats = buckets[bucket_name]
        bucket_total     = sum(averages.get(c, 0) for c in cats)
        bucket_raw_total = sum(totals.get(c, 0) for c in cats)
        pct_spending = (bucket_total / all_spent) * 100 if all_spent else 0
        pct_income   = (bucket_raw_total / income) * 100 if income    else 0
        ws[f"A{current_row}"] = bucket_name.upper()
        ws[f"B{current_row}"] = bucket_raw_total
        ws[f"C{current_row}"] = bucket_total
        ws[f"D{current_row}"] = f"{pct_spending:.1f}%"
        ws[f"E{current_row}"] = f"{pct_income:.1f}%"
        for col in ["A", "B", "C", "D", "E"]:
            style_fn(ws[f"{col}{current_row}"],
                     font=Font(name="Arial", bold=True, size=11), align=center)
        ws[f"B{current_row}"].number_format = '"$"#,##0.00'
        ws[f"C{current_row}"].number_format = '"$"#,##0.00'
        ws.row_dimensions[current_row].height = 22
        current_row += 1

    # Total spending row
    ws[f"A{current_row}"] = "TOTAL SPENDING"
    ws[f"B{current_row}"] = all_raw_spent
    ws[f"C{current_row}"] = all_spent
    ws[f"D{current_row}"] = "100.0%"
    ws[f"E{current_row}"] = f"{(all_spent / income * 100):.1f}%" if income else "0.0%"
    for col in ["A", "B", "C", "D", "E"]:
        style_fn(ws[f"{col}{current_row}"], font=white_bold, fill=dark_fill, align=center)
    ws[f"B{current_row}"].number_format = '"$"#,##0.00'
    ws[f"C{current_row}"].number_format = '"$"#,##0.00'
    ws.row_dimensions[current_row].height = 22

    # Remaining Income row
    current_row += 1
    remaining_income = round(income - all_raw_spent, 2)
    if remaining_income >= 0:
        rem_fill = PatternFill("solid", start_color="D5F5E3")
        rem_font = Font(name="Arial", bold=True, size=11, color="145A32")
    else:
        rem_fill = PatternFill("solid", start_color="FDECEA")
        rem_font = Font(name="Arial", bold=True, size=11, color="C0392B")
    ws[f"A{current_row}"] = "Remaining Income"
    ws[f"B{current_row}"] = remaining_income
    ws[f"C{current_row}"] = ""
    ws[f"D{current_row}"] = "—"
    ws[f"E{current_row}"] = "—"
    for col in ["A", "B", "C", "D", "E"]:
        ws[f"{col}{current_row}"].fill      = rem_fill
        ws[f"{col}{current_row}"].font      = rem_font
        ws[f"{col}{current_row}"].alignment = center
        ws[f"{col}{current_row}"].border    = border
    ws[f"B{current_row}"].number_format = '"$"#,##0.00'
    ws.row_dimensions[current_row].height = 22


def build_output_xlsx(averages, assignments, totals=None):
    """
    Build the clean 3-bucket output spreadsheet.
    assignments = { category_name: bucket_label, ... }
    totals = { category_name: raw_sum_across_months, ... }
    Returns the path to a temp .xlsx file.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Awareness Engine"
    totals = totals or {}

    # ── Styles ────────────────────────────────────────────────────────────────
    bucket_font = Font(name="Arial", bold=True, size=12)
    label_font  = Font(name="Arial", size=11)
    pct_font    = Font(name="Arial", bold=True, size=12)
    total_font  = Font(name="Arial", bold=True, size=11)
    title_fill  = PatternFill("solid", start_color="2C3E50")
    title_font  = Font(name="Arial", bold=True, size=14, color="FFFFFF")

    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")
    right  = Alignment(horizontal="right",  vertical="center")

    thin   = Side(style="thin", color="AAAAAA")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    def style(cell, font=None, fill=None, align=None, num_fmt=None):
        if font:    cell.font      = font
        if fill:    cell.fill      = fill
        if align:   cell.alignment = align
        if num_fmt: cell.number_format = num_fmt
        cell.border = border

    bucket_fill = {
        "Necessities": PatternFill("solid", start_color="D6EAF8"),
        "Luxuries":    PatternFill("solid", start_color="D5F5E3"),
        "Future Self": PatternFill("solid", start_color="FEF9E7"),
    }

    # ── Derive full bucket order (built-ins first, then custom) ───────────────
    _custom_palette = ["EDE7F6", "E0F2F1", "FCE4EC", "FFF8E1", "F3E5F5"]
    known_buckets  = set(BUCKET_LABELS)
    custom_buckets = list(dict.fromkeys(
        b for b in assignments.values() if b != "Leave Out" and b not in known_buckets
    ))
    all_buckets = BUCKET_LABELS + custom_buckets
    # Assign a fill color to each custom bucket (cycling through the palette)
    for i, b in enumerate(custom_buckets):
        bucket_fill[b] = PatternFill("solid", start_color=_custom_palette[i % len(_custom_palette)])

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells("A1:E1")
    ws["A1"] = "THE AWARENESS ENGINE — My Financial Reality"
    style(ws["A1"], font=title_font, fill=title_fill, align=center)
    ws.row_dimensions[1].height = 30

    # ── Income row ────────────────────────────────────────────────────────────
    ws.row_dimensions[2].height = 6   # spacer
    income = averages.get("income", 0)
    ws["A3"] = "Total Income"
    ws["B3"] = income
    ws["E3"] = "100%"
    style(ws["A3"], font=total_font, align=left)
    style(ws["B3"], font=total_font, align=right, num_fmt='"$"#,##0.00')
    style(ws["C3"], font=total_font, align=right)
    style(ws["D3"], font=total_font, align=right)
    style(ws["E3"], font=pct_font,   align=center)
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 8   # spacer

    # ── Group categories by bucket (all buckets including custom) ─────────────
    buckets = {b: [] for b in all_buckets}
    for cat, bucket in assignments.items():
        if bucket in buckets:
            buckets[bucket].append(cat)

    # ── Bucket sections ───────────────────────────────────────────────────────
    next_row = _render_bucket_sections(
        ws, averages, assignments, totals, buckets, income,
        style, bucket_fill, label_font, bucket_font, center, right, left, border,
        start_row=5, all_buckets=all_buckets
    )

    # ── Grand Reveal ──────────────────────────────────────────────────────────
    _render_grand_reveal(
        ws, averages, assignments, totals, buckets, income,
        style, center, right, border, start_row=next_row, all_buckets=all_buckets
    )

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    return tmp.name


# ── Budget template builders ───────────────────────────────────────────────────

def _write_grand_reveal_section(ws, row, bucket_totals, bucket_fills, income,
                                 border, center, reveal_label):
    """Write the Grand Reveal block into the template sheet. Returns nothing."""
    reveal_fill  = PatternFill("solid", fgColor="0F1F3D")
    reveal_font  = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    sub_hdr_fill = PatternFill("solid", fgColor="F4F4F1")
    sub_hdr_font = Font(name="Calibri", bold=True, size=11)
    total_fill   = PatternFill("solid", fgColor="2C3E50")
    total_font_w = Font(name="Calibri", bold=True, size=11, color="FFFFFF")

    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"]           = reveal_label
    ws[f"A{row}"].font      = reveal_font
    ws[f"A{row}"].fill      = reveal_fill
    ws[f"A{row}"].alignment = center
    ws[f"A{row}"].border    = border
    row += 1

    for col, label in [("A", "Bucket"), ("B", "Amount"),
                        ("C", "% of Income"), ("D", "% of Spending")]:
        ws[f"{col}{row}"]           = label
        ws[f"{col}{row}"].font      = sub_hdr_font
        ws[f"{col}{row}"].fill      = sub_hdr_fill
        ws[f"{col}{row}"].alignment = center
        ws[f"{col}{row}"].border    = border
    row += 1

    # Custom fallback fill for any bucket not in the built-in palette
    _fallback_fill = PatternFill("solid", fgColor="EDE7F6")

    total_spending = sum(bucket_totals.values())
    for bucket_name, amt in bucket_totals.items():
        b_fill  = bucket_fills.get(bucket_name, _fallback_fill)
        pct_inc = round((amt / income * 100), 1) if income else 0.0
        pct_spd = round((amt / total_spending * 100), 1) if total_spending else 0.0
        ws[f"A{row}"] = bucket_name
        ws[f"B{row}"] = amt
        ws[f"C{row}"] = f"{pct_inc:.1f}%"
        ws[f"D{row}"] = f"{pct_spd:.1f}%"
        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{row}"].fill      = b_fill
            ws[f"{col}{row}"].font      = Font(name="Calibri", bold=True, size=11)
            ws[f"{col}{row}"].alignment = center
            ws[f"{col}{row}"].border    = border
        ws[f"B{row}"].number_format = '"$"#,##0.00'
        row += 1

    total_pct_inc = round((total_spending / income * 100), 1) if income else 0.0
    ws[f"A{row}"] = "TOTAL EXPENSES"
    ws[f"B{row}"] = total_spending
    ws[f"C{row}"] = f"{total_pct_inc:.1f}%"
    ws[f"D{row}"] = "100.0%"
    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{row}"].font      = total_font_w
        ws[f"{col}{row}"].fill      = total_fill
        ws[f"{col}{row}"].alignment = center
        ws[f"{col}{row}"].border    = border
    ws[f"B{row}"].number_format = '"$"#,##0.00'
    row += 1

    remaining     = income - total_spending
    remaining_pct = round((remaining / income * 100), 1) if income else 0.0
    rem_fill = PatternFill("solid", fgColor="D5F5E3") if remaining >= 0 \
               else PatternFill("solid", fgColor="FDECEA")
    rem_font = Font(name="Calibri", bold=True, size=11, color="145A32") if remaining >= 0 \
               else Font(name="Calibri", bold=True, size=11, color="C0392B")
    ws[f"A{row}"] = "Remaining Income"
    ws[f"B{row}"] = remaining
    ws[f"C{row}"] = f"{remaining_pct:.1f}%"
    ws[f"D{row}"] = "—"
    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{row}"].fill      = rem_fill
        ws[f"{col}{row}"].font      = rem_font
        ws[f"{col}{row}"].alignment = center
        ws[f"{col}{row}"].border    = border
    ws[f"B{row}"].number_format = '"$"#,##0.00'


def build_single_month_template(income, buckets_input):
    """
    Build a single-month budget template .xlsx.
    buckets_input = { bucket_name: [{name, amount}, ...] }
    Returns path to a temp .xlsx file.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Budget Template"

    st          = _make_budget_template_styles()
    header_font = st["header_font"]
    bucket_fonts = st["bucket_fonts"]
    bucket_fills = st["bucket_fills"]
    header_fill  = st["header_fill"]
    border       = st["border"]
    center       = st["center"]
    right_align  = st["right_align"]

    ws.merge_cells("A1:B1")
    ws["A1"]           = "Budget Template"
    ws["A1"].font      = header_font
    ws["A1"].fill      = header_fill
    ws["A1"].alignment = center

    ws["A2"]              = "Monthly Income"
    ws["A2"].font         = Font(name="Calibri", bold=True, size=12)
    ws["B2"]              = income
    ws["B2"].number_format = '"$"#,##0.00'
    ws["B2"].alignment    = right_align
    ws["A2"].border       = border
    ws["B2"].border       = border

    # Ordered bucket list: built-ins first, then any custom buckets from input
    _known = set(BUCKET_LABELS)
    ordered_buckets = [b for b in BUCKET_LABELS if b in buckets_input] + \
                      [b for b in buckets_input if b not in _known]
    # Extend font/fill dicts with fallback styles for custom buckets
    _custom_colors_tpl = ["EDE7F6", "E0F2F1", "FCE4EC"]
    _custom_idx = 0
    for b in ordered_buckets:
        if b not in bucket_fills:
            bucket_fills[b] = PatternFill("solid", fgColor=_custom_colors_tpl[_custom_idx % len(_custom_colors_tpl)])
            bucket_fonts[b]  = Font(name="Calibri", bold=True, size=12, color="4A235A")
            _custom_idx += 1

    row = 4
    bucket_totals = {}
    for bucket_name in ordered_buckets:
        cats = buckets_input.get(bucket_name, [])
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"]           = bucket_name.upper()
        ws[f"A{row}"].font      = bucket_fonts[bucket_name]
        ws[f"A{row}"].fill      = bucket_fills[bucket_name]
        ws[f"A{row}"].alignment = center
        ws[f"A{row}"].border    = border
        row += 1

        total = 0.0
        for item in cats:
            name = item.get("name", "")
            amt  = float(item.get("amount", 0))
            total += amt
            ws[f"A{row}"]              = name
            ws[f"B{row}"]              = amt
            ws[f"B{row}"].number_format = '"$"#,##0.00'
            ws[f"A{row}"].border       = border
            ws[f"B{row}"].border       = border
            ws[f"B{row}"].alignment    = right_align
            row += 1

        bucket_totals[bucket_name] = total
        ws[f"A{row}"]              = "Subtotal"
        ws[f"A{row}"].font         = Font(name="Calibri", bold=True, size=11)
        ws[f"B{row}"]              = total
        ws[f"B{row}"].number_format = '"$"#,##0.00'
        ws[f"B{row}"].font         = Font(name="Calibri", bold=True, size=11)
        ws[f"A{row}"].border       = border
        ws[f"B{row}"].border       = border
        ws[f"B{row}"].alignment    = right_align
        row += 2

    row += 1  # spacer before Grand Reveal
    _write_grand_reveal_section(ws, row, bucket_totals, bucket_fills, income,
                                 border, center, "⚡ GRAND REVEAL")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


def build_three_month_template(incomes, month_names, buckets_input):
    """
    Build a 3-month average budget template .xlsx.
    incomes = [float, float, float]
    month_names = [str, str, str]
    buckets_input = { bucket_name: [{name, amounts: [m1, m2, m3]}, ...] }
    Returns path to a temp .xlsx file.
    """
    avg_income = sum(incomes) / 3

    wb = Workbook()
    ws = wb.active
    ws.title = "Budget Template"

    st           = _make_budget_template_styles()
    header_font  = st["header_font"]
    bucket_fonts = st["bucket_fonts"]
    bucket_fills = st["bucket_fills"]
    header_fill  = st["header_fill"]
    border       = st["border"]
    center       = st["center"]
    right_align  = st["right_align"]

    ws.merge_cells("A1:B1")
    ws["A1"]           = f"Last 3 Months Budget — {' / '.join(month_names)}"
    ws["A1"].font      = header_font
    ws["A1"].fill      = header_fill
    ws["A1"].alignment = center

    ws["A2"]              = "Avg Monthly Income"
    ws["A2"].font         = Font(name="Calibri", bold=True, size=12)
    ws["B2"]              = round(avg_income, 2)
    ws["B2"].number_format = '"$"#,##0.00'
    ws["B2"].alignment    = right_align
    ws["A2"].border       = border
    ws["B2"].border       = border

    # Ordered bucket list: built-ins first, then any custom buckets from input
    _known = set(BUCKET_LABELS)
    ordered_buckets = [b for b in BUCKET_LABELS if b in buckets_input] + \
                      [b for b in buckets_input if b not in _known]
    # Extend font/fill dicts with fallback styles for custom buckets
    _custom_colors_tpl = ["EDE7F6", "E0F2F1", "FCE4EC"]
    _custom_idx = 0
    for b in ordered_buckets:
        if b not in bucket_fills:
            bucket_fills[b] = PatternFill("solid", fgColor=_custom_colors_tpl[_custom_idx % len(_custom_colors_tpl)])
            bucket_fonts[b]  = Font(name="Calibri", bold=True, size=12, color="4A235A")
            _custom_idx += 1

    row = 4
    bucket_totals = {}
    for bucket_name in ordered_buckets:
        cats = buckets_input.get(bucket_name, [])
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"]           = bucket_name.upper()
        ws[f"A{row}"].font      = bucket_fonts[bucket_name]
        ws[f"A{row}"].fill      = bucket_fills[bucket_name]
        ws[f"A{row}"].alignment = center
        ws[f"A{row}"].border    = border
        row += 1

        total = 0.0
        for item in cats:
            name    = item.get("name", "")
            amounts = item.get("amounts", [0, 0, 0])
            avg_amt = sum(amounts) / 3
            total  += avg_amt
            ws[f"A{row}"]              = name
            ws[f"B{row}"]              = round(avg_amt, 2)
            ws[f"B{row}"].number_format = '"$"#,##0.00'
            ws[f"A{row}"].border       = border
            ws[f"B{row}"].border       = border
            ws[f"B{row}"].alignment    = right_align
            row += 1

        bucket_totals[bucket_name] = total
        ws[f"A{row}"]              = "Subtotal"
        ws[f"A{row}"].font         = Font(name="Calibri", bold=True, size=11)
        ws[f"B{row}"]              = round(total, 2)
        ws[f"B{row}"].number_format = '"$"#,##0.00'
        ws[f"B{row}"].font         = Font(name="Calibri", bold=True, size=11)
        ws[f"A{row}"].border       = border
        ws[f"B{row}"].border       = border
        ws[f"B{row}"].alignment    = right_align
        row += 2

    row += 1  # spacer before Grand Reveal
    _write_grand_reveal_section(ws, row, bucket_totals, bucket_fills, avg_income,
                                 border, center, "⚡ GRAND REVEAL (3-Month Average)")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()
    return tmp.name
