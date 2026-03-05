"""
analyzer.py – Advanced data analysis, anomaly detection and trend detection
===========================================================================
Used by the FastAPI backend.  All computation is local.
"""

from __future__ import annotations

import math
from typing import Any

import numpy as np
import pandas as pd
from scipy import stats


# ──────────────────────────────────────────────
#  Descriptive analysis
# ──────────────────────────────────────────────


def analyze_dataframe(df: pd.DataFrame) -> dict[str, Any]:
    """
    Return per-column descriptive statistics for a DataFrame.
    """
    results: list[dict] = []

    for col in df.columns:
        series = df[col].dropna()
        col_type = _infer_type(series)

        info: dict[str, Any] = {
            "column": str(col),
            "type": col_type,
            "count": int(len(df[col])),
            "missing": int(df[col].isna().sum()),
        }

        if col_type == "numeric":
            numeric = pd.to_numeric(series, errors="coerce").dropna()
            desc = numeric.describe()
            q1  = float(numeric.quantile(0.25))
            q3  = float(numeric.quantile(0.75))
            info.update(
                {
                    "mean":     _safe_float(desc["mean"]),
                    "median":   _safe_float(numeric.median()),
                    "std":      _safe_float(desc["std"]),
                    "min":      _safe_float(desc["min"]),
                    "max":      _safe_float(desc["max"]),
                    "q1":       _safe_float(q1),
                    "q3":       _safe_float(q3),
                    "iqr":      _safe_float(q3 - q1),
                    "sum":      _safe_float(numeric.sum()),
                    "skewness": _safe_float(float(numeric.skew())),
                    "kurtosis": _safe_float(float(numeric.kurtosis())),
                }
            )
        elif col_type == "text":
            vc = series.value_counts()
            info.update(
                {
                    "unique_count": int(series.nunique()),
                    "top_values": [
                        {"value": str(v), "count": int(c)}
                        for v, c in vc.head(5).items()
                    ],
                }
            )
        elif col_type == "datetime":
            dt = pd.to_datetime(series, errors="coerce").dropna()
            info.update(
                {
                    "min_date": str(dt.min()),
                    "max_date": str(dt.max()),
                    "range_days": int((dt.max() - dt.min()).days),
                }
            )

        results.append(info)

    return {
        "columns": results,
        "total_rows": int(len(df)),
        "total_cols": int(len(df.columns)),
        "report": _build_text_report(results),
    }


def _build_text_report(columns: list[dict]) -> str:
    lines = ["📊 Data Analysis Report", "=" * 40]
    for c in columns:
        lines.append(f"\n🔹 {c['column']}  [{c['type']}]")
        lines.append(f"   Count: {c['count']}  |  Missing: {c['missing']}")
        if c["type"] == "numeric":
            lines.append(f"   Min:  {c['min']}   Max:  {c['max']}")
            lines.append(f"   Mean: {c['mean']}   Median: {c['median']}")
            lines.append(f"   Std:  {c['std']}   IQR:  {c['iqr']}")
            lines.append(f"   Skewness: {c['skewness']}   Kurtosis: {c['kurtosis']}")
        elif c["type"] == "text":
            tops = ", ".join(f"{t['value']} ({t['count']})" for t in c.get("top_values", []))
            lines.append(f"   Unique: {c['unique_count']}  Top: {tops}")
        elif c["type"] == "datetime":
            lines.append(f"   Range: {c['min_date']} → {c['max_date']} ({c['range_days']} days)")
    return "\n".join(lines)


# ──────────────────────────────────────────────
#  Anomaly detection
# ──────────────────────────────────────────────


def detect_anomalies(
    df: pd.DataFrame,
    z_threshold: float = 3.0,
    use_isolation_forest: bool = False,
) -> list[dict[str, Any]]:
    """
    Detect anomalies in numeric columns using IQR and Z-score.
    Optionally also uses Isolation Forest (scikit-learn).
    """
    anomalies: list[dict] = []

    for col in df.columns:
        series = df[col]
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.isna().all():
            continue

        clean = numeric.dropna()
        if len(clean) < 4:
            continue

        mean = float(clean.mean())
        std  = float(clean.std(ddof=1))
        q1   = float(clean.quantile(0.25))
        q3   = float(clean.quantile(0.75))
        iqr  = q3 - q1
        lower_fence = q1 - 1.5 * iqr
        upper_fence = q3 + 1.5 * iqr

        for idx, val in numeric.items():
            if pd.isna(val):
                continue
            v = float(val)
            z = abs(v - mean) / std if std > 0 else 0.0
            is_z   = z > z_threshold
            is_iqr = v < lower_fence or v > upper_fence

            if is_z or is_iqr:
                severity = (
                    "high"   if z > 4.0 else
                    "medium" if z > z_threshold else
                    "low"
                )
                anomalies.append(
                    {
                        "row_index": int(idx),
                        "column":    str(col),
                        "value":     v,
                        "z_score":   round(z, 4),
                        "mean":      round(mean, 4),
                        "std":       round(std, 4),
                        "method":    ("Z-score + IQR" if is_z and is_iqr
                                      else "Z-score" if is_z else "IQR"),
                        "severity":  severity,
                    }
                )

    # Optional: Isolation Forest for multivariate detection
    if use_isolation_forest:
        anomalies = _merge_isolation_forest(df, anomalies)

    return sorted(anomalies, key=lambda a: -abs(a["z_score"]))


def _merge_isolation_forest(df: pd.DataFrame, existing: list[dict]) -> list[dict]:
    """Add Isolation Forest anomalies (requires scikit-learn)."""
    try:
        from sklearn.ensemble import IsolationForest

        num_df = df.select_dtypes(include=[np.number]).dropna()
        if len(num_df) < 10:
            return existing

        clf = IsolationForest(contamination=0.05, random_state=42)
        preds = clf.fit_predict(num_df)

        existing_rows = {a["row_index"] for a in existing}
        for i, (idx, pred) in enumerate(zip(num_df.index, preds)):
            if pred == -1 and int(idx) not in existing_rows:
                existing.append(
                    {
                        "row_index": int(idx),
                        "column":    "multi-column",
                        "value":     None,
                        "z_score":   0.0,
                        "mean":      0.0,
                        "std":       0.0,
                        "method":    "Isolation Forest",
                        "severity":  "medium",
                    }
                )
    except ImportError:
        pass  # scikit-learn not installed
    return existing


# ──────────────────────────────────────────────
#  Trend detection
# ──────────────────────────────────────────────


def detect_trends(df: pd.DataFrame) -> list[dict[str, Any]]:
    """
    Detect linear trends in numeric columns using scipy.stats.linregress.
    """
    trends: list[dict] = []

    for col in df.columns:
        numeric = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(numeric) < 3:
            continue

        x = np.arange(len(numeric))
        y = numeric.values.astype(float)

        slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
        r2 = r_value ** 2

        direction = "upward" if slope > 0 else "downward" if slope < 0 else "flat"
        strength  = "strong" if r2 > 0.7 else "moderate" if r2 > 0.4 else "weak"

        trends.append(
            {
                "column":     str(col),
                "direction":  direction,
                "slope":      round(float(slope), 6),
                "intercept":  round(float(intercept), 6),
                "r2":         round(float(r2), 4),
                "p_value":    round(float(p_value), 6),
                "std_err":    round(float(std_err), 6),
                "strength":   strength,
                "description": (
                    f"{col}: {direction} trend (slope={slope:.4f}, R²={r2:.3f}) – {strength}"
                ),
            }
        )

    return sorted(trends, key=lambda t: -t["r2"])


# ──────────────────────────────────────────────
#  Helpers
# ──────────────────────────────────────────────


def _infer_type(series: pd.Series) -> str:
    if series.empty:
        return "text"

    # Try numeric
    numeric = pd.to_numeric(series, errors="coerce")
    if numeric.notna().mean() > 0.8:
        return "numeric"

    # Try datetime
    try:
        pd.to_datetime(series, errors="raise", format="mixed")
        return "datetime"
    except Exception:
        pass

    return "text"


def _safe_float(val: Any) -> float | None:
    try:
        f = float(val)
        return round(f, 6) if math.isfinite(f) else None
    except (TypeError, ValueError):
        return None
