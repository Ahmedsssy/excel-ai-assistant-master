/**
 * data-analyzer.js – Client-side data analysis utilities
 *
 * Pure JS functions for descriptive statistics, anomaly detection, and
 * trend analysis.  These run entirely in the browser with no external
 * dependencies, keeping all data private.
 */

"use strict";

/* ─────────────────────────────────────────────
   1.  Column type inference
───────────────────────────────────────────── */

export const COL_TYPES = Object.freeze({
  NUMERIC: "numeric",
  DATE: "date",
  TEXT: "text",
  BOOLEAN: "boolean",
  MIXED: "mixed",
});

/**
 * Infer the type of a column from its values.
 * @param {any[]} values  Flat array of cell values (one column).
 * @returns {string}      One of COL_TYPES.
 */
export function inferColumnType(values) {
  const nonEmpty = values.filter((v) => v !== null && v !== "");
  if (nonEmpty.length === 0) return COL_TYPES.TEXT;

  let numCount = 0;
  let dateCount = 0;
  let boolCount = 0;
  let textCount = 0;

  for (const v of nonEmpty) {
    if (typeof v === "boolean") { boolCount++; continue; }
    if (typeof v === "number" || (typeof v === "string" && !isNaN(parseFloat(v)) && isFinite(v))) {
      numCount++;
      continue;
    }
    if (typeof v === "string" && !isNaN(Date.parse(v))) { dateCount++; continue; }
    textCount++;
  }

  const total = nonEmpty.length;
  if (numCount / total > 0.8)  return COL_TYPES.NUMERIC;
  if (dateCount / total > 0.8) return COL_TYPES.DATE;
  if (boolCount / total > 0.8) return COL_TYPES.BOOLEAN;
  if (textCount / total > 0.5) return COL_TYPES.TEXT;
  return COL_TYPES.MIXED;
}

/* ─────────────────────────────────────────────
   2.  Descriptive statistics
───────────────────────────────────────────── */

/**
 * Compute descriptive statistics for an array of numbers.
 */
export function describeNumeric(values) {
  const nums = values
    .filter((v) => v !== null && v !== undefined && v !== "")
    .map(Number)
    .filter((n) => !isNaN(n));
  if (nums.length === 0) return null;

  nums.sort((a, b) => a - b);

  const n   = nums.length;
  const sum = nums.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  const variance = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0) / n;
  const std = Math.sqrt(variance);
  const min = nums[0];
  const max = nums[n - 1];
  const q1  = percentile(nums, 25);
  const median = percentile(nums, 50);
  const q3  = percentile(nums, 75);
  const iqr = q3 - q1;

  return { n, sum, mean, median, std, variance, min, max, q1, q3, iqr, range: max - min };
}

function percentile(sorted, p) {
  const idx = (p / 100) * (sorted.length - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return sorted[lo];
  return sorted[lo] * (hi - idx) + sorted[hi] * (idx - lo);
}

/**
 * Analyse every column in the dataset and return an array of column summaries.
 */
export function analyseDataset(data) {
  if (!data || !data.rows || data.rows.length === 0) return [];

  const numCols = data.rows[0].length;
  const results = [];

  for (let c = 0; c < numCols; c++) {
    const colValues = data.rows.map((r) => r[c]);
    const type = inferColumnType(colValues);
    const header = data.headers ? data.headers[c] : `Column ${c + 1}`;

    const summary = { header, type, count: colValues.length };

    if (type === COL_TYPES.NUMERIC) {
      const stats = describeNumeric(colValues);
      Object.assign(summary, stats);

      // Missing / empty
      summary.missing = colValues.filter((v) => v === null || v === "").length;
    } else if (type === COL_TYPES.TEXT) {
      const nonEmpty = colValues.filter((v) => v !== null && v !== "");
      const freq = {};
      nonEmpty.forEach((v) => { freq[String(v)] = (freq[String(v)] || 0) + 1; });
      const sorted = Object.entries(freq).sort((a, b) => b[1] - a[1]);
      summary.uniqueCount = sorted.length;
      summary.topValues   = sorted.slice(0, 5);
      summary.missing = colValues.length - nonEmpty.length;
    }

    results.push(summary);
  }

  return results;
}

/**
 * Format a column summary array into a human-readable report string.
 */
export function formatAnalysisReport(summaries) {
  if (!summaries.length) return "No data to analyse.";

  const lines = ["📊 Data Analysis Report", "═".repeat(40)];

  for (const s of summaries) {
    lines.push(`\n🔹 ${s.header}  [${s.type}]`);
    lines.push(`   Count: ${s.count}  |  Missing: ${s.missing ?? 0}`);

    if (s.type === COL_TYPES.NUMERIC) {
      lines.push(`   Min:    ${fmt(s.min)}   Max:    ${fmt(s.max)}`);
      lines.push(`   Mean:   ${fmt(s.mean)}   Median: ${fmt(s.median)}`);
      lines.push(`   Std:    ${fmt(s.std)}   IQR:    ${fmt(s.iqr)}`);
      lines.push(`   Q1:     ${fmt(s.q1)}   Q3:     ${fmt(s.q3)}`);
    } else if (s.type === COL_TYPES.TEXT) {
      lines.push(`   Unique values: ${s.uniqueCount}`);
      if (s.topValues?.length) {
        lines.push(`   Top values: ${s.topValues.map(([v, c]) => `${v} (${c})`).join(", ")}`);
      }
    }
  }

  return lines.join("\n");
}

function fmt(n) {
  if (n === undefined || n === null) return "N/A";
  return Number(n.toFixed(4)).toLocaleString();
}

/* ─────────────────────────────────────────────
   3.  Anomaly detection  (IQR & Z-score)
───────────────────────────────────────────── */

/**
 * Detect outliers in numeric columns using IQR and Z-score methods.
 *
 * Returns an array of anomaly objects:
 * { rowIndex, colIndex, header, value, method, severity }
 */
export function detectAnomalies(data, zThreshold = 3.0) {
  if (!data || !data.rows || data.rows.length === 0) return [];

  const numCols = data.rows[0].length;
  const anomalies = [];

  for (let c = 0; c < numCols; c++) {
    const colValues = data.rows.map((r) => r[c]);
    const type = inferColumnType(colValues);
    if (type !== COL_TYPES.NUMERIC) continue;

    const header = data.headers ? data.headers[c] : `Column ${c + 1}`;
    const nums   = colValues.map(Number).filter((n) => !isNaN(n));
    const stats  = describeNumeric(nums);
    if (!stats) continue;

    const lowerFence = stats.q1 - 1.5 * stats.iqr;
    const upperFence = stats.q3 + 1.5 * stats.iqr;

    for (let r = 0; r < data.rows.length; r++) {
      const raw = data.rows[r][c];
      if (raw === null || raw === "") continue;
      const v = Number(raw);
      if (isNaN(v)) continue;

      const zScore = stats.std > 0 ? Math.abs((v - stats.mean) / stats.std) : 0;
      const isIQR  = v < lowerFence || v > upperFence;
      const isZ    = zScore > zThreshold;

      if (isIQR || isZ) {
        // Use both z-score and IQR multiplier to determine severity
        const iqrMultiplier = stats.iqr > 0
          ? Math.abs(v - (v < lowerFence ? lowerFence : upperFence)) / stats.iqr
          : 0;
        const severity =
          zScore > 4 || iqrMultiplier > 3
            ? "high"
            : zScore > zThreshold || iqrMultiplier > 1.5
            ? "medium"
            : "low";

        anomalies.push({
          rowIndex: r,            // 0-based index in data.rows
          colIndex: c,
          header,
          value: v,
          mean: stats.mean,
          std: stats.std,
          zScore,
          method: isZ && isIQR ? "Z-score + IQR" : isZ ? "Z-score" : "IQR",
          severity,
        });
      }
    }
  }

  return anomalies;
}

/**
 * Build a text report from anomaly results.
 */
export function formatAnomalyReport(anomalies, headers = null) {
  if (!anomalies.length) return "✅ No anomalies detected in the numeric columns.";

  const lines = [`🔍 Anomaly Detection Report – ${anomalies.length} anomaly(ies) found`, "═".repeat(45)];

  for (const a of anomalies) {
    const rowLabel = `Row ${a.rowIndex + 2}`; // +1 for header row, +1 for 1-based
    lines.push(
      `\n⚠️  [${a.severity.toUpperCase()}] ${a.header} – ${rowLabel}` +
        `\n   Value: ${fmt(a.value)}  |  Mean: ${fmt(a.mean)}  |  Std: ${fmt(a.std)}` +
        `\n   Z-score: ${a.zScore.toFixed(2)}  |  Method: ${a.method}`
    );
  }

  return lines.join("\n");
}

/* ─────────────────────────────────────────────
   4.  Trend detection
───────────────────────────────────────────── */

/**
 * Simple linear regression: returns { slope, intercept, r2 }
 */
export function linearRegression(xArr, yArr) {
  const n = Math.min(xArr.length, yArr.length);
  if (n < 2) return null;

  const xs = xArr.slice(0, n).map(Number);
  const ys = yArr.slice(0, n).map(Number);

  const meanX = xs.reduce((a, b) => a + b, 0) / n;
  const meanY = ys.reduce((a, b) => a + b, 0) / n;

  let ssXY = 0, ssXX = 0, ssTot = 0;
  for (let i = 0; i < n; i++) {
    ssXY += (xs[i] - meanX) * (ys[i] - meanY);
    ssXX += (xs[i] - meanX) ** 2;
    ssTot += (ys[i] - meanY) ** 2;
  }

  const slope     = ssXX !== 0 ? ssXY / ssXX : 0;
  const intercept = meanY - slope * meanX;
  const ssRes     = ys.reduce((acc, y, i) => acc + (y - (slope * xs[i] + intercept)) ** 2, 0);
  const r2        = ssTot !== 0 ? 1 - ssRes / ssTot : 0;

  return { slope, intercept, r2 };
}

/**
 * Detect monotonic trends across all numeric columns.
 * Returns array of { header, direction, slope, r2, description }
 */
export function detectTrends(data) {
  if (!data || !data.rows || data.rows.length < 3) return [];

  const numCols = data.rows[0].length;
  const trends  = [];
  const xArr    = data.rows.map((_, i) => i);

  for (let c = 0; c < numCols; c++) {
    const header = data.headers ? data.headers[c] : `Column ${c + 1}`;
    const colVals = data.rows.map((r) => r[c]);
    if (inferColumnType(colVals) !== COL_TYPES.NUMERIC) continue;

    const yArr = colVals.map(Number).filter((v) => !isNaN(v));
    if (yArr.length < 3) continue;

    const reg = linearRegression(xArr.slice(0, yArr.length), yArr);
    if (!reg) continue;

    const direction =
      reg.slope > 0 ? "upward" : reg.slope < 0 ? "downward" : "flat";

    trends.push({
      header,
      direction,
      slope: reg.slope,
      r2: reg.r2,
      description:
        `${header}: ${direction} trend (slope=${fmt(reg.slope)}, R²=${reg.r2.toFixed(3)})` +
        (reg.r2 > 0.7 ? " – strong" : reg.r2 > 0.4 ? " – moderate" : " – weak"),
    });
  }

  return trends;
}

/**
 * Format trend results into a readable report.
 */
export function formatTrendReport(trends) {
  if (!trends.length) return "No numeric columns found for trend analysis.";

  const lines = ["📈 Trend Analysis Report", "═".repeat(40)];
  for (const t of trends) {
    const arrow = t.direction === "upward" ? "⬆️" : t.direction === "downward" ? "⬇️" : "➡️";
    lines.push(`\n${arrow}  ${t.description}`);
  }
  return lines.join("\n");
}
