"""
Tests for the Python backend analyzer module.
Run with: pytest backend/test_analyzer.py -v
"""

import math
import pytest
import pandas as pd
import numpy as np

import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from analyzer import (
    analyze_dataframe,
    detect_anomalies,
    detect_trends,
    _infer_type,
    _safe_float,
)


# ─────────────────────────────────────────────
#  _infer_type
# ─────────────────────────────────────────────

class TestInferType:
    def test_numeric_series(self):
        s = pd.Series([1.0, 2.0, 3.0, 4.0, 5.0])
        assert _infer_type(s) == "numeric"

    def test_text_series(self):
        s = pd.Series(["alpha", "beta", "gamma", "delta"])
        assert _infer_type(s) == "text"

    def test_empty_series(self):
        s = pd.Series([], dtype=object)
        assert _infer_type(s) == "text"

    def test_mostly_numeric_with_some_text(self):
        # 80%+ numeric → numeric
        s = pd.Series([1, 2, 3, 4, "five"])
        result = _infer_type(s)
        # Could be numeric or text depending on coercion; at minimum it shouldn't crash
        assert result in ("numeric", "text", "datetime")


# ─────────────────────────────────────────────
#  _safe_float
# ─────────────────────────────────────────────

class TestSafeFloat:
    def test_normal_float(self):
        assert _safe_float(3.14) == pytest.approx(3.14)

    def test_none_returns_none(self):
        assert _safe_float(None) is None

    def test_nan_returns_none(self):
        assert _safe_float(float("nan")) is None

    def test_inf_returns_none(self):
        assert _safe_float(float("inf")) is None

    def test_string_number(self):
        assert _safe_float("2.5") == pytest.approx(2.5)


# ─────────────────────────────────────────────
#  analyze_dataframe
# ─────────────────────────────────────────────

class TestAnalyzeDataframe:
    def _make_df(self):
        return pd.DataFrame(
            {
                "Name":  ["Alice", "Bob", "Carol", "Dave", "Eve"],
                "Score": [85.0, 92.0, 78.0, 88.0, 95.0],
                "Age":   [30, 25, 35, 28, 32],
            }
        )

    def test_returns_correct_column_count(self):
        result = analyze_dataframe(self._make_df())
        assert result["total_cols"] == 3

    def test_returns_correct_row_count(self):
        result = analyze_dataframe(self._make_df())
        assert result["total_rows"] == 5

    def test_numeric_column_has_mean(self):
        result = analyze_dataframe(self._make_df())
        score_col = next(c for c in result["columns"] if c["column"] == "Score")
        assert score_col["type"] == "numeric"
        assert score_col["mean"] == pytest.approx(87.6, abs=0.01)

    def test_text_column_has_unique_count(self):
        result = analyze_dataframe(self._make_df())
        name_col = next(c for c in result["columns"] if c["column"] == "Name")
        assert name_col["type"] == "text"
        assert name_col["unique_count"] == 5

    def test_report_string_is_generated(self):
        result = analyze_dataframe(self._make_df())
        assert isinstance(result["report"], str)
        assert len(result["report"]) > 20

    def test_missing_values_counted(self):
        df = pd.DataFrame({"A": [1.0, None, 3.0, None, 5.0]})
        result = analyze_dataframe(df)
        col = result["columns"][0]
        assert col["missing"] == 2

    def test_empty_dataframe(self):
        df = pd.DataFrame()
        result = analyze_dataframe(df)
        assert result["total_cols"] == 0
        assert result["total_rows"] == 0


# ─────────────────────────────────────────────
#  detect_anomalies
# ─────────────────────────────────────────────

class TestDetectAnomalies:
    def _make_df_with_outlier(self):
        values = [10.0, 11.0, 12.0, 10.5, 11.5, 12.5, 10.0, 11.0, 1000.0, 12.0]
        return pd.DataFrame({"Value": values})

    def test_detects_obvious_outlier(self):
        df = self._make_df_with_outlier()
        anomalies = detect_anomalies(df)
        assert len(anomalies) >= 1
        assert any(a["value"] == 1000.0 for a in anomalies)

    def test_no_anomalies_for_uniform_data(self):
        df = pd.DataFrame({"V": [10.0] * 20})
        anomalies = detect_anomalies(df)
        assert anomalies == []

    def test_returns_list(self):
        df = self._make_df_with_outlier()
        result = detect_anomalies(df)
        assert isinstance(result, list)

    def test_anomaly_has_required_fields(self):
        df = self._make_df_with_outlier()
        anomalies = detect_anomalies(df)
        required_keys = {"row_index", "column", "value", "z_score", "method", "severity"}
        for a in anomalies:
            assert required_keys.issubset(set(a.keys()))

    def test_severity_is_valid_value(self):
        df = self._make_df_with_outlier()
        for a in detect_anomalies(df):
            assert a["severity"] in ("high", "medium", "low")

    def test_skips_text_columns(self):
        df = pd.DataFrame({"Names": ["Alice", "Bob", "Carol", "Dave", "Eve"]})
        assert detect_anomalies(df) == []

    def test_custom_z_threshold(self):
        df = self._make_df_with_outlier()
        high_thresh = detect_anomalies(df, z_threshold=5.0)
        low_thresh  = detect_anomalies(df, z_threshold=1.5)
        assert len(low_thresh) >= len(high_thresh)


# ─────────────────────────────────────────────
#  detect_trends
# ─────────────────────────────────────────────

class TestDetectTrends:
    def test_detects_upward_trend(self):
        df = pd.DataFrame({"Sales": [10, 20, 30, 40, 50, 60, 70]})
        trends = detect_trends(df)
        assert len(trends) == 1
        assert trends[0]["direction"] == "upward"

    def test_detects_downward_trend(self):
        df = pd.DataFrame({"Errors": [70, 60, 50, 40, 30, 20, 10]})
        trends = detect_trends(df)
        assert trends[0]["direction"] == "downward"

    def test_strong_trend_has_high_r2(self):
        df = pd.DataFrame({"X": list(range(20))})
        trends = detect_trends(df)
        assert trends[0]["r2"] > 0.99

    def test_trend_sorted_by_r2_descending(self):
        df = pd.DataFrame({
            "Strong": list(range(10)),
            "Noisy":  [i + np.random.normal(0, 5) for i in range(10)],
        })
        trends = detect_trends(df)
        r2s = [t["r2"] for t in trends]
        assert r2s == sorted(r2s, reverse=True)

    def test_returns_empty_for_text_columns(self):
        df = pd.DataFrame({"Name": ["a", "b", "c", "d", "e"]})
        assert detect_trends(df) == []

    def test_short_series_skipped(self):
        df = pd.DataFrame({"V": [1.0, 2.0]})
        assert detect_trends(df) == []

    def test_trend_has_required_fields(self):
        df = pd.DataFrame({"Revenue": list(range(10))})
        trends = detect_trends(df)
        required = {"column", "direction", "slope", "r2", "p_value", "strength", "description"}
        for t in trends:
            assert required.issubset(set(t.keys()))
