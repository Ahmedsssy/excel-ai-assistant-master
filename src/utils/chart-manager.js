/**
 * chart-manager.js – Smart chart type recommendation and creation
 *
 * Analyses column types to recommend the most appropriate chart, then
 * delegates to the Excel API layer to insert it.
 */

"use strict";

import { COL_TYPES, inferColumnType } from "./data-analyzer.js";
import { createChart } from "./excel-api.js";

/* ─────────────────────────────────────────────
   Chart type map (Excel.ChartType enum values)
───────────────────────────────────────────── */

export const CHART_TYPES = Object.freeze({
  ColumnClustered: "ColumnClustered",
  Line:            "Line",
  Pie:             "Pie",
  BarClustered:    "BarClustered",
  Area:            "Area",
  Scatter:         "Scatter",
  Histogram:       "Histogram",
  BoxWhisker:      "BoxWhisker",
});

/* ─────────────────────────────────────────────
   Recommendation logic
───────────────────────────────────────────── */

/**
 * Recommend the best chart type for the given data.
 *
 * Rules (in priority order):
 *  1. 1 text col + 1 numeric col + ≤ 7 rows  → Pie
 *  2. 1 text col + 1 numeric col              → BarClustered
 *  3. 1 date col + ≥1 numeric col             → Line (time-series)
 *  4. 2+ numeric cols, one looks like X-axis  → Scatter
 *  5. 1 numeric col (distribution)            → Histogram
 *  6. Multiple numeric cols                   → ColumnClustered (default)
 *
 * @param {object} data  Shape: { headers, rows, rowCount, colCount }
 * @returns {{ type: string, reason: string }}
 */
export function recommendChartType(data) {
  if (!data || !data.rows || data.rows.length === 0) {
    return { type: CHART_TYPES.ColumnClustered, reason: "Default chart (no data detected)." };
  }

  const numCols = data.rows[0].length;
  const colTypes = Array.from({ length: numCols }, (_, c) =>
    inferColumnType(data.rows.map((r) => r[c]))
  );

  const textCols    = colTypes.filter((t) => t === COL_TYPES.TEXT).length;
  const numericCols = colTypes.filter((t) => t === COL_TYPES.NUMERIC).length;
  const dateCols    = colTypes.filter((t) => t === COL_TYPES.DATE).length;
  const rows        = data.rows.length;

  // Pie: single category + single numeric + few rows
  if (textCols === 1 && numericCols === 1 && rows <= 7) {
    return {
      type: CHART_TYPES.Pie,
      reason: "Pie chart — categorical labels with a single numeric series and few items.",
    };
  }

  // Bar: single category + single numeric + many rows
  if (textCols === 1 && numericCols === 1) {
    return {
      type: CHART_TYPES.BarClustered,
      reason: "Horizontal bar chart — comparing values across many categories.",
    };
  }

  // Line: date/time column present → time series
  if (dateCols >= 1 && numericCols >= 1) {
    return {
      type: CHART_TYPES.Line,
      reason: "Line chart — date/time column detected, ideal for time-series trends.",
    };
  }

  // Scatter: 2+ numeric cols (x-y relationship)
  if (numericCols >= 2 && textCols === 0) {
    return {
      type: CHART_TYPES.Scatter,
      reason: "Scatter plot — multiple numeric columns, useful for correlation analysis.",
    };
  }

  // Histogram: single numeric column (distribution)
  if (numericCols === 1 && textCols === 0) {
    return {
      type: CHART_TYPES.Histogram,
      reason: "Histogram — single numeric column, shows value distribution.",
    };
  }

  // Default: clustered column
  return {
    type: CHART_TYPES.ColumnClustered,
    reason: "Clustered column chart — comparing multiple series across categories.",
  };
}

/* ─────────────────────────────────────────────
   Chart creation helpers
───────────────────────────────────────────── */

/**
 * Auto-recommend and create the best chart for the selection.
 */
export async function autoChart(data, rangeAddress) {
  const { type, reason } = recommendChartType(data);
  const title = buildChartTitle(data, type);
  await createChart(rangeAddress, type, title);
  return { type, reason, title };
}

/**
 * Create a chart of a specific type for the selection.
 */
export async function insertChart(rangeAddress, chartType, data) {
  const title = buildChartTitle(data, chartType);
  await createChart(rangeAddress, chartType, title);
  return { type: chartType, title };
}

/**
 * Build a descriptive chart title from the data headers.
 */
function buildChartTitle(data, chartType) {
  if (!data?.headers) return `${chartType} Chart`;
  const numericHeaders = data.headers.filter(Boolean).slice(0, 3).join(", ");
  return numericHeaders ? `${numericHeaders}` : `${chartType} Chart`;
}
