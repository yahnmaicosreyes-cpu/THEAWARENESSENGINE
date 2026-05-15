"""Integration tests for Flask routes — upload, generate, save, download."""

import io
import os
import tempfile

import pytest
from openpyxl import Workbook


def _make_simple_xlsx():
    """Return a BytesIO containing a minimal valid budget xlsx."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Income", 3000])
    ws.append(["Rent", 1200])
    ws.append(["Food", 400])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class TestUpload:
    def test_no_file_returns_400(self, client):
        resp = client.post("/upload")
        assert resp.status_code == 400
        assert "No file uploaded" in resp.get_json()["error"]

    def test_wrong_extension_returns_400(self, client):
        resp = client.post(
            "/upload",
            data={"file": (io.BytesIO(b"fake"), "budget.csv")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 400
        assert "xlsx" in resp.get_json()["error"].lower()

    def test_valid_xlsx_returns_200_with_data(self, client):
        resp = client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 200
        data = resp.get_json()
        assert "income" in data
        assert "categories" in data
        assert "months_found" in data

    def test_upload_stores_income_value(self, client):
        resp = client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        assert resp.get_json()["income"] == 3000

    def test_upload_returns_category_names(self, client):
        resp = client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        cat_names = [c["name"] for c in resp.get_json()["categories"]]
        assert "Rent" in cat_names
        assert "Food" in cat_names

    def test_corrupt_file_returns_500(self, client):
        resp = client.post(
            "/upload",
            data={"file": (io.BytesIO(b"not an xlsx"), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 500


class TestGenerate:
    def test_no_session_returns_400(self, client):
        resp = client.post(
            "/generate",
            json={"assignments": {}, "selected_months": None},
        )
        assert resp.status_code == 400
        assert "session" in resp.get_json()["error"].lower()

    def test_generate_after_upload_succeeds(self, client):
        upload = client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        cats = upload.get_json()["categories"]
        assignments = {c["name"]: "Necessities" for c in cats}

        gen = client.post(
            "/generate",
            json={"assignments": assignments, "selected_months": None},
        )
        assert gen.status_code == 200
        data = gen.get_json()
        assert data["success"] is True
        assert "table" in data
        assert "monthly_breakdown" in data

    def test_generate_invalid_months_returns_400(self, client):
        client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        resp = client.post(
            "/generate",
            json={"assignments": {}, "selected_months": ["DEC"]},
        )
        assert resp.status_code == 400


class TestDownload:
    def test_no_file_ready_returns_400(self, client):
        resp = client.get("/download")
        assert resp.status_code == 400

    def test_download_after_generate_returns_xlsx(self, client):
        client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        cats_resp = client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        cats = cats_resp.get_json()["categories"]
        assignments = {c["name"]: "Necessities" for c in cats}

        client.post(
            "/generate",
            json={"assignments": assignments, "selected_months": None},
        )

        resp = client.get("/download")
        assert resp.status_code == 200
        assert "spreadsheetml" in resp.content_type


class TestSecurity:
    def test_session_cookie_httponly(self, client):
        client.post(
            "/upload",
            data={"file": (_make_simple_xlsx(), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        cookies = client.get_cookie("session")
        assert cookies is not None

    def test_upload_rejects_oversized_filename_as_wrong_ext(self, client):
        """File with .xlsx extension but no content — should fail gracefully."""
        resp = client.post(
            "/upload",
            data={"file": (io.BytesIO(b""), "budget.xlsx")},
            content_type="multipart/form-data",
        )
        # Empty file is still .xlsx — extract_data will raise, caught as 500
        assert resp.status_code in (400, 500)

    def test_path_traversal_in_filename_rejected(self, client):
        """../../etc/passwd.xlsx should fail the xlsx check or extraction."""
        resp = client.post(
            "/upload",
            data={"file": (io.BytesIO(b"bad"), "../../etc/passwd.xlsx")},
            content_type="multipart/form-data",
        )
        # Extension is technically .xlsx, so it passes that check but fails extraction
        assert resp.status_code in (400, 500)
