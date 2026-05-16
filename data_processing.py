"""
data_processing.py — Computation layer.

Transforms raw monthly data (from extractors.py) into averages, totals,
and the structured table payload the frontend renders.
"""

from extractors import BUCKET_LABELS


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

    keys = list(dict.fromkeys(k for m in active_months for k in monthly_data[m].keys()))
    result = {}
    for k in keys:
        total = sum(monthly_data[m].get(k, 0) for m in active_months)
        if k == "income":
            result[k] = round(total, 2)          # grand total across all months
        else:
            result[k] = round(total / n, 2)      # average across months
    return result


def compute_totals(monthly_data, selected_months=None):
    """Return raw sum (not average) for each category across active months."""
    if selected_months:
        active_months = [m for m in selected_months if m in monthly_data]
    else:
        active_months = [m for m, data in monthly_data.items() if data.get("income", 0) > 0]
    if not active_months:
        return {}
    keys = list(dict.fromkeys(k for m in active_months for k in monthly_data[m].keys()))
    return {k: round(sum(monthly_data[m].get(k, 0) for m in active_months), 2) for k in keys}


def build_table_data(averages, assignments, totals=None):
    """Build structured table data for frontend rendering."""
    income   = averages.get("income", 0)
    left_out = {cat for cat, b in assignments.items() if b == "Leave Out"}
    totals   = totals or {}

    # Derive full ordered bucket list — built-ins first, then any custom buckets the
    # user created (preserved in insertion order via dict.fromkeys).
    known = set(BUCKET_LABELS)
    custom = [b for b in dict.fromkeys(assignments.values())
              if b != "Leave Out" and b not in known]
    all_buckets = BUCKET_LABELS + custom

    buckets_out = []
    for bucket_name in all_buckets:
        cats = [cat for cat, b in assignments.items() if b == bucket_name]
        categories = [
            {"name": cat,
             "amt": round(averages.get(cat, 0), 2),
             "raw_total": round(totals.get(cat, 0), 2)}
            for cat in cats
        ]
        bucket_total     = round(sum(c["amt"] for c in categories), 2)
        bucket_raw_total = round(sum(c["raw_total"] for c in categories), 2)
        pct_income       = round((bucket_raw_total / income * 100), 1) if income else 0.0
        buckets_out.append({
            "name":       bucket_name,
            "categories": categories,
            "total":      bucket_total,
            "raw_total":  bucket_raw_total,
            "pct_income": pct_income,
        })

    all_spent     = round(sum(v for k, v in averages.items()
                              if k != "income" and k not in left_out), 2)
    all_raw_spent = round(sum(totals.get(k, 0) for k in averages
                               if k != "income" and k not in left_out), 2)

    grand_reveal = [
        {
            "name":        b["name"],
            "total":       b["total"],
            "raw_total":   b["raw_total"],
            "pct_spending": round((b["total"] / all_spent * 100), 1) if all_spent else 0.0,
            "pct_income":  b["pct_income"],
        }
        for b in buckets_out
    ]

    return {
        "income":            round(income, 2),
        "buckets":           buckets_out,
        "grand_reveal":      grand_reveal,
        "total_spending":    all_spent,
        "total_raw_spending": all_raw_spent,
        "remaining_income":  round(income - all_raw_spent, 2),
        "total_pct_income":  round((all_raw_spent / income * 100), 1) if income else 0.0,
    }
