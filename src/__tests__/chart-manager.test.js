/**
 * Tests for chart-manager.js recommendation logic.
 * Uses mock for excel-api.js and data-analyzer.js so no Office.js needed.
 */

import { recommendChartType, CHART_TYPES } from "../utils/chart-manager.js";
import { COL_TYPES } from "../utils/data-analyzer.js";

// ─────────────────────────────────────────────
//  Helpers to build synthetic datasets
// ─────────────────────────────────────────────

function buildData(headers, rows) {
  return { headers, rows, rowCount: rows.length, colCount: rows[0]?.length || 0 };
}

// ─────────────────────────────────────────────
//  recommendChartType
// ─────────────────────────────────────────────

describe("recommendChartType", () => {
  test("returns default for null data", () => {
    const rec = recommendChartType(null);
    expect(rec.type).toBe(CHART_TYPES.ColumnClustered);
  });

  test("returns Pie for 1 text + 1 numeric column with ≤7 rows", () => {
    const data = buildData(
      ["Category", "Value"],
      [
        ["A", 10], ["B", 20], ["C", 15], ["D", 25], ["E", 30],
      ]
    );
    const rec = recommendChartType(data);
    expect(rec.type).toBe(CHART_TYPES.Pie);
  });

  test("returns BarClustered for 1 text + 1 numeric column with >7 rows", () => {
    const rows = Array.from({ length: 10 }, (_, i) => [`Cat${i}`, i * 10]);
    const data = buildData(["Category", "Value"], rows);
    const rec = recommendChartType(data);
    expect(rec.type).toBe(CHART_TYPES.BarClustered);
  });

  test("returns Histogram for single numeric column", () => {
    const data = buildData(
      ["Score"],
      [[72], [85], [90], [60], [78], [88], [65], [92]]
    );
    const rec = recommendChartType(data);
    expect(rec.type).toBe(CHART_TYPES.Histogram);
  });

  test("returns Scatter for 2+ numeric columns", () => {
    const data = buildData(
      ["X", "Y"],
      [[1, 2], [3, 4], [5, 6], [7, 8], [9, 10]]
    );
    const rec = recommendChartType(data);
    expect(rec.type).toBe(CHART_TYPES.Scatter);
  });

  test("returns ColumnClustered as fallback for mixed data", () => {
    const data = buildData(
      ["Name", "Q1", "Q2", "Q3"],
      [
        ["Alice", 100, 200, 300],
        ["Bob",   150, 250, 350],
        ["Carol", 120, 220, 320],
      ]
    );
    const rec = recommendChartType(data);
    expect(rec.type).toBe(CHART_TYPES.ColumnClustered);
  });

  test("recommendation includes a reason string", () => {
    const data = buildData(["Val"], [[1], [2], [3]]);
    const rec = recommendChartType(data);
    expect(typeof rec.reason).toBe("string");
    expect(rec.reason.length).toBeGreaterThan(5);
  });
});
