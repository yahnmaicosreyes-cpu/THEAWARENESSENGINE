"""
app.py — Flask application entry point and routes.

All business logic lives in separate modules:
  extractors.py      — spreadsheet parsing
  data_processing.py — averages, totals, table data
  excel_builders.py  — Excel file generation
"""

import os
import secrets
import tempfile
from flask import Flask, render_template, request, jsonify, send_file, session, after_this_request

from extractors import extract_data, BUCKET_LABELS
from data_processing import compute_averages, compute_totals, build_table_data
from excel_builders import build_output_xlsx, build_single_month_template, build_three_month_template

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config["SESSION_COOKIE_SECURE"]   = True   # only send cookie over HTTPS
app.config["SESSION_COOKIE_HTTPONLY"] = True   # block JavaScript from reading the cookie
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"  # block cross-site request forgery
app.config["MAX_CONTENT_LENGTH"]      = 10 * 1024 * 1024  # 10 MB upload limit


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

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    try:
        f.save(tmp.name)
        monthly_data, category_names = extract_data(tmp.name)

        # Deduplicate category names (preserving order) in case the spreadsheet
        # lists the same category more than once under the same month.
        seen = set()
        category_names = [n for n in category_names if not (n in seen or seen.add(n))]

        session["monthly_data"] = monthly_data
        averages = compute_averages(monthly_data)
        categories = [{"name": n, "amount": averages.get(n, 0)} for n in category_names]

        return jsonify({
            "income":       averages.get("income", 0),
            "categories":   categories,
            "months_found": list(monthly_data.keys()),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp.name)


@app.route("/generate", methods=["POST"])
def generate():
    data            = request.get_json()
    assignments     = data.get("assignments", {})
    selected_months = data.get("selected_months", None)

    monthly_data = session.get("monthly_data")
    if not monthly_data:
        return jsonify({"error": "No data in session. Please upload your file again."}), 400

    averages = compute_averages(monthly_data, selected_months)
    if not averages:
        return jsonify({"error": "No data found for the selected months."}), 400

    totals = compute_totals(monthly_data, selected_months)

    try:
        output_path = build_output_xlsx(averages, assignments, totals)
        session["output_path"] = output_path
        session["assignments"] = assignments

        active_months = selected_months if selected_months else \
            [m for m, d in monthly_data.items() if d.get("income", 0) > 0]
        session["active_months"] = active_months

        monthly_breakdown = {
            cat: {m: round(monthly_data[m].get(cat, 0), 2) for m in active_months}
            for cat, bucket in assignments.items()
            if bucket not in ("Leave Out",) and bucket in BUCKET_LABELS
        }

        return jsonify({
            "success":           True,
            "table":             build_table_data(averages, assignments, totals),
            "monthly_breakdown": monthly_breakdown,
            "trend_months":      active_months,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/save", methods=["POST"])
def save():
    data          = request.get_json()
    income        = float(data.get("income", 0))
    buckets_input = data.get("buckets", {})

    averages    = {"income": income}
    assignments = {}
    for bucket_name, cats in buckets_input.items():
        for cat in cats:
            averages[cat["name"]]    = float(cat["amt"])
            assignments[cat["name"]] = bucket_name

    monthly_data  = session.get("monthly_data")
    active_months = session.get("active_months")
    totals        = compute_totals(monthly_data, active_months) if monthly_data else {}

    try:
        old_path = session.get("output_path")
        if old_path and os.path.exists(old_path):
            os.unlink(old_path)

        output_path = build_output_xlsx(averages, assignments, totals)
        session["output_path"]  = output_path
        session["averages"]     = averages
        session["assignments"]  = assignments
        return jsonify({"success": True, "table": build_table_data(averages, assignments, totals)})
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
    data   = request.get_json(force=True)
    income = float(data.get("income", 0))
    buckets = data.get("buckets", {})

    tmp_path = build_single_month_template(income, buckets)

    @after_this_request
    def cleanup_tmpl(response):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return response

    return send_file(tmp_path, as_attachment=True, download_name="Budget_Template.xlsx")


@app.route("/three-month/download", methods=["POST"])
def three_month_download():
    data         = request.get_json(force=True)
    incomes      = [float(v) for v in data.get("incomes", [0, 0, 0])]
    month_names  = data.get("month_names", ["Month 1", "Month 2", "Month 3"])
    buckets      = data.get("buckets", {})

    tmp_path = build_three_month_template(incomes, month_names, buckets)

    @after_this_request
    def cleanup_three(response):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return response

    return send_file(tmp_path, as_attachment=True, download_name="Last_3_Months_Budget.xlsx")


if __name__ == "__main__":
    app.run(debug=True, port=5050)
