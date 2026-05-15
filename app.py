"""
app.py — Flask application entry point and routes.

All business logic lives in separate modules:
  extractors.py      — spreadsheet parsing
  data_processing.py — averages, totals, table data
  excel_builders.py  — Excel file generation
"""

import base64
import io
import json
import os
import secrets
import tempfile
import warnings
from flask import Flask, render_template, request, jsonify, send_file, session, Response
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

from extractors import extract_data, BUCKET_LABELS
from data_processing import compute_averages, compute_totals, build_table_data
from excel_builders import build_output_xlsx, build_single_month_template, build_three_month_template

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
if not os.environ.get("SECRET_KEY"):
    warnings.warn(
        "SECRET_KEY env var is not set — using a random key. "
        "All sessions will be invalidated on every server restart.",
        RuntimeWarning, stacklevel=1,
    )
app.config["SESSION_COOKIE_SECURE"]   = True   # only send cookie over HTTPS
app.config["SESSION_COOKIE_HTTPONLY"] = True   # block JavaScript from reading the cookie
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"  # block cross-site request forgery
app.config["MAX_CONTENT_LENGTH"]      = 10 * 1024 * 1024  # 10 MB upload limit

limiter = Limiter(
    key_func=get_remote_address,
    app=app,
    storage_uri="memory://",
    default_limits=[],
)


# ── HTTP Basic Auth ───────────────────────────────────────────────────────────

def _check_basic_auth():
    """Validate the Authorization header against env-var credentials.

    Credentials are never stored in code — set BASIC_AUTH_USERNAME and
    BASIC_AUTH_PASSWORD as environment variables (e.g. in Render's dashboard).
    Returns True only when both env vars are set AND the header matches.
    """
    expected_user = os.environ.get("BASIC_AUTH_USERNAME", "")
    expected_pass = os.environ.get("BASIC_AUTH_PASSWORD", "")

    # If credentials aren't configured, block all access to prevent
    # accidentally running an unprotected instance in production.
    if not expected_user or not expected_pass:
        return False

    header = request.headers.get("Authorization", "")
    if not header.startswith("Basic "):
        return False

    try:
        # Decode the base64 "username:password" payload from the header.
        decoded  = base64.b64decode(header[6:]).decode("utf-8")
        username, _, password = decoded.partition(":")
        return username == expected_user and password == expected_pass
    except Exception:
        return False


@app.before_request
def require_auth():
    """Gate every request behind HTTP Basic Auth.

    The browser caches credentials for the session, so the user is only
    prompted once. To sign out the user must close the browser or manually
    clear saved passwords — there is no server-side logout for Basic Auth.
    """
    if _check_basic_auth():
        return  # credentials valid — let the request through

    # Return 401 with WWW-Authenticate to trigger the browser's login dialog.
    return Response(
        "Authentication required.",
        401,
        {"WWW-Authenticate": 'Basic realm="The Awareness Engine"'},
    )


# ── Helpers ───────────────────────────────────────────────────────────────────

def _clear_monthly_data():
    """Delete the server-side JSON file for the current session's monthly data."""
    path = session.get("monthly_data_path")
    if path and os.path.exists(path):
        try:
            os.unlink(path)
        except Exception:
            pass


def _load_monthly_data():
    """Load monthly data from the server-side JSON file."""
    path = session.get("monthly_data_path")
    if not path or not os.path.exists(path):
        return None
    with open(path, "r") as f:
        return json.load(f)


def _serve_and_delete(file_path, download_name):
    """Read file into memory buffer, delete from disk, then serve the buffer.

    Avoids Windows file-lock errors that occur when trying to unlink a file
    that send_file still has open.
    """
    buf = io.BytesIO()
    with open(file_path, "rb") as f:
        buf.write(f.read())
    os.unlink(file_path)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    _clear_monthly_data()
    session.clear()
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
@limiter.limit("20 per minute")
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Please upload a .xlsx file"}), 400

    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        f.save(tmp_path)
        monthly_data, category_names = extract_data(tmp_path)

        # Deduplicate category names (preserving order) in case the spreadsheet
        # lists the same category more than once under the same month.
        seen = set()
        category_names = [n for n in category_names if not (n in seen or seen.add(n))]

        # Store monthly_data server-side as JSON to avoid the ~4 KB cookie limit.
        _clear_monthly_data()
        tmp_json = tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode="w")
        json.dump(monthly_data, tmp_json)
        tmp_json.close()
        session["monthly_data_path"] = tmp_json.name

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
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


@app.route("/generate", methods=["POST"])
@limiter.limit("30 per minute")
def generate():
    data            = request.get_json()
    assignments     = data.get("assignments", {})
    selected_months = data.get("selected_months", None)

    monthly_data = _load_monthly_data()
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

    monthly_data  = _load_monthly_data()
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
@limiter.limit("20 per minute")
def download():
    output_path = session.get("output_path")
    if not output_path or not os.path.exists(output_path):
        return "No file ready. Please generate first.", 400

    session.pop("output_path", None)
    return _serve_and_delete(output_path, "Awareness_Engine_Results.xlsx")


@app.route("/template/download", methods=["POST"])
def template_download():
    data    = request.get_json(force=True)
    income  = float(data.get("income", 0))
    buckets = data.get("buckets", {})

    tmp_path = build_single_month_template(income, buckets)
    return _serve_and_delete(tmp_path, "Budget_Template.xlsx")


@app.route("/three-month/download", methods=["POST"])
def three_month_download():
    data        = request.get_json(force=True)
    incomes     = [float(v) for v in data.get("incomes", [0, 0, 0])]
    month_names = data.get("month_names", ["Month 1", "Month 2", "Month 3"])
    buckets     = data.get("buckets", {})

    tmp_path = build_three_month_template(incomes, month_names, buckets)
    return _serve_and_delete(tmp_path, "Last_3_Months_Budget.xlsx")


@app.route("/download/inspiration-template")
def inspiration_template():
    path = os.path.join(os.path.dirname(__file__), "SpreadSheet Budget Template.xlsx")
    return send_file(path, as_attachment=True,
                     download_name="Awareness_Engine_Budget_Template.xlsx")


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "false").lower() == "true", port=5050)
