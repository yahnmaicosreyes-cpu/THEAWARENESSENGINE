"""Tests for data_processing.py — averages, totals, and table structure."""

import pytest


MONTHLY_DATA = {
    "JAN": {"income": 3000, "Food": 400, "Rent": 1200},
    "FEB": {"income": 3500, "Food": 350, "Rent": 1200},
}


class TestComputeAverages:
    def test_income_is_summed_not_averaged(self):
        from data_processing import compute_averages
        result = compute_averages(MONTHLY_DATA)
        assert result["income"] == 6500

    def test_expenses_are_averaged(self):
        from data_processing import compute_averages
        result = compute_averages(MONTHLY_DATA)
        assert result["Food"] == 375.0
        assert result["Rent"] == 1200.0

    def test_selected_months_subset(self):
        from data_processing import compute_averages
        result = compute_averages(MONTHLY_DATA, selected_months=["JAN"])
        assert result["income"] == 3000
        assert result["Food"] == 400

    def test_zero_income_months_excluded_intentionally(self):
        """$0 income months are excluded by design — living off savings scenario
        is handled by the user selecting specific months instead."""
        from data_processing import compute_averages
        data = {
            "JAN": {"income": 3000, "Food": 400},
            "FEB": {"income": 0, "Food": 300},
        }
        result = compute_averages(data)
        assert result["income"] == 3000
        assert result["Food"] == 400  # only JAN active

    def test_no_active_months_returns_empty(self):
        from data_processing import compute_averages
        result = compute_averages({"JAN": {"income": 0}})
        assert result == {}

    def test_selected_month_not_in_data_ignored(self):
        from data_processing import compute_averages
        result = compute_averages(MONTHLY_DATA, selected_months=["JAN", "DEC"])
        assert result["income"] == 3000   # only JAN found
        assert result["Food"] == 400


class TestComputeTotals:
    def test_sums_across_active_months(self):
        from data_processing import compute_totals
        result = compute_totals(MONTHLY_DATA)
        assert result["income"] == 6500
        assert result["Food"] == 750
        assert result["Rent"] == 2400

    def test_selected_months_only(self):
        from data_processing import compute_totals
        result = compute_totals(MONTHLY_DATA, selected_months=["FEB"])
        assert result["income"] == 3500
        assert result["Food"] == 350

    def test_no_active_months_returns_empty(self):
        from data_processing import compute_totals
        result = compute_totals({"JAN": {"income": 0}})
        assert result == {}


class TestBuildTableData:
    def _averages(self):
        return {"income": 3000, "Rent": 1200, "Groceries": 400, "Entertainment": 100}

    def _assignments(self):
        return {
            "Rent": "Necessities",
            "Groceries": "Necessities",
            "Entertainment": "Leave Out",
        }

    def test_top_level_keys_present(self):
        from data_processing import build_table_data
        result = build_table_data(self._averages(), self._assignments())
        for key in ("income", "buckets", "grand_reveal", "total_spending",
                    "total_raw_spending", "remaining_income", "total_pct_income"):
            assert key in result

    def test_income_value_correct(self):
        from data_processing import build_table_data
        result = build_table_data(self._averages(), self._assignments())
        assert result["income"] == 3000

    def test_leave_out_excluded_from_total_spending(self):
        from data_processing import build_table_data
        result = build_table_data(self._averages(), self._assignments())
        # Entertainment (100) is "Leave Out" — should not be in total
        assert result["total_spending"] == 1600  # Rent 1200 + Groceries 400

    def test_remaining_income_calculation(self):
        from data_processing import build_table_data
        averages = self._averages()
        # Pass averages as totals (single-month equivalent) so remaining_income is meaningful
        result = build_table_data(averages, self._assignments(), totals=averages)
        # Rent (1200) + Groceries (400) = 1600; Entertainment is Leave Out
        assert result["remaining_income"] == 3000 - 1600

    def test_bucket_categories_grouped_correctly(self):
        from data_processing import build_table_data
        result = build_table_data(self._averages(), self._assignments())
        necessities = next(b for b in result["buckets"] if b["name"] == "Necessities")
        cat_names = [c["name"] for c in necessities["categories"]]
        assert "Rent" in cat_names
        assert "Groceries" in cat_names
        assert "Entertainment" not in cat_names
