/**
 * Tests for data-analyzer.js
 *
 * These tests exercise the pure-JS analysis functions that run in the browser
 * without any Office.js or network dependencies.
 */

// The module uses ES module syntax; babel-jest handles the transform.
import {
  inferColumnType,
  COL_TYPES,
  describeNumeric,
  analyseDataset,
  detectAnomalies,
  detectTrends,
  linearRegression,
  formatAnalysisReport,
  formatAnomalyReport,
  formatTrendReport,
} from "../utils/data-analyzer.js";

// ─────────────────────────────────────────────
//  inferColumnType
// ─────────────────────────────────────────────

describe("inferColumnType", () => {
  test("detects numeric column", () => {
    expect(inferColumnType([1, 2, 3, 4, 5])).toBe(COL_TYPES.NUMERIC);
  });

  test("detects numeric strings as numeric", () => {
    expect(inferColumnType(["1.5", "2.0", "3.14", "0"])).toBe(COL_TYPES.NUMERIC);
  });

  test("detects text column", () => {
    expect(inferColumnType(["apple", "banana", "cherry", "date"])).toBe(COL_TYPES.TEXT);
  });

  test("detects boolean column", () => {
    expect(inferColumnType([true, false, true, true, false])).toBe(COL_TYPES.BOOLEAN);
  });

  test("returns text for empty column", () => {
    expect(inferColumnType([])).toBe(COL_TYPES.TEXT);
    expect(inferColumnType([null, null, ""])).toBe(COL_TYPES.TEXT);
  });

  test("mixed numeric and text returns mixed or text", () => {
    const result = inferColumnType([1, 2, "hello", "world", "foo"]);
    expect([COL_TYPES.MIXED, COL_TYPES.TEXT]).toContain(result);
  });
});

// ─────────────────────────────────────────────
//  describeNumeric
// ─────────────────────────────────────────────

describe("describeNumeric", () => {
  const values = [10, 20, 30, 40, 50];

  test("returns correct mean", () => {
    const stats = describeNumeric(values);
    expect(stats.mean).toBe(30);
  });

  test("returns correct median", () => {
    const stats = describeNumeric(values);
    expect(stats.median).toBe(30);
  });

  test("returns correct min/max", () => {
    const stats = describeNumeric(values);
    expect(stats.min).toBe(10);
    expect(stats.max).toBe(50);
  });

  test("returns correct n", () => {
    const stats = describeNumeric(values);
    expect(stats.n).toBe(5);
  });

  test("returns correct sum", () => {
    const stats = describeNumeric(values);
    expect(stats.sum).toBe(150);
  });

  test("returns null for empty / all-NaN array", () => {
    expect(describeNumeric([])).toBeNull();
    expect(describeNumeric(["abc", "def"])).toBeNull();
  });

  test("filters out NaN / non-numeric values", () => {
    const stats = describeNumeric([10, null, 20, "x", 30]);
    expect(stats.n).toBe(3);
    expect(stats.mean).toBe(20);
  });

  test("returns non-negative std", () => {
    const stats = describeNumeric([5, 5, 5, 5, 5]);
    expect(stats.std).toBe(0);
  });
});

// ─────────────────────────────────────────────
//  analyseDataset
// ─────────────────────────────────────────────

describe("analyseDataset", () => {
  const data = {
    headers: ["Name", "Score", "Age"],
    rows: [
      ["Alice", 85, 30],
      ["Bob",   92, 25],
      ["Carol", 78, 35],
      ["Dave",  88, 28],
    ],
  };

  test("returns a summary per column", () => {
    const summaries = analyseDataset(data);
    expect(summaries).toHaveLength(3);
  });

  test("first column is text type", () => {
    const summaries = analyseDataset(data);
    expect(summaries[0].type).toBe(COL_TYPES.TEXT);
  });

  test("second column (Score) is numeric type", () => {
    const summaries = analyseDataset(data);
    expect(summaries[1].type).toBe(COL_TYPES.NUMERIC);
  });

  test("numeric column has mean", () => {
    const summaries = analyseDataset(data);
    const scoreSummary = summaries[1];
    expect(scoreSummary.mean).toBeCloseTo(85.75, 1);
  });

  test("handles empty data gracefully", () => {
    expect(analyseDataset(null)).toEqual([]);
    expect(analyseDataset({ rows: [] })).toEqual([]);
  });
});

// ─────────────────────────────────────────────
//  detectAnomalies
// ─────────────────────────────────────────────

describe("detectAnomalies", () => {
  // Clear outlier at index 4 (value=1000 among 10–30 range)
  const dataWithOutlier = {
    headers: ["Value"],
    rows: [[10], [15], [12], [14], [1000], [11], [13], [16], [12], [14]],
  };

  test("detects obvious outlier", () => {
    const anomalies = detectAnomalies(dataWithOutlier);
    expect(anomalies.length).toBeGreaterThan(0);
    const outlierRow = anomalies.find((a) => a.value === 1000);
    expect(outlierRow).toBeDefined();
  });

  test("outlier has high severity", () => {
    const anomalies = detectAnomalies(dataWithOutlier);
    const outlier = anomalies.find((a) => a.value === 1000);
    expect(outlier.severity).toBe("high");
  });

  test("returns empty array for clean data", () => {
    const cleanData = {
      headers: ["Value"],
      rows: [[10], [11], [12], [10], [11], [12], [10], [11], [12], [11]],
    };
    const anomalies = detectAnomalies(cleanData);
    expect(anomalies).toHaveLength(0);
  });

  test("skips text columns", () => {
    const textData = {
      headers: ["Name"],
      rows: [["Alice"], ["Bob"], ["Carol"], ["Dave"], ["Eve"]],
    };
    expect(detectAnomalies(textData)).toHaveLength(0);
  });

  test("returns empty array for null/empty data", () => {
    expect(detectAnomalies(null)).toEqual([]);
    expect(detectAnomalies({ rows: [] })).toEqual([]);
  });
});

// ─────────────────────────────────────────────
//  linearRegression
// ─────────────────────────────────────────────

describe("linearRegression", () => {
  test("perfect upward linear trend gives slope=1, r2=1", () => {
    const x = [0, 1, 2, 3, 4];
    const y = [0, 1, 2, 3, 4];
    const reg = linearRegression(x, y);
    expect(reg.slope).toBeCloseTo(1, 5);
    expect(reg.r2).toBeCloseTo(1, 5);
  });

  test("perfect downward trend gives negative slope", () => {
    const x = [0, 1, 2, 3, 4];
    const y = [10, 8, 6, 4, 2];
    const reg = linearRegression(x, y);
    expect(reg.slope).toBeLessThan(0);
    expect(reg.r2).toBeCloseTo(1, 5);
  });

  test("flat data gives slope near 0", () => {
    const x = [0, 1, 2, 3];
    const y = [5, 5, 5, 5];
    const reg = linearRegression(x, y);
    expect(reg.slope).toBeCloseTo(0, 5);
  });

  test("returns null for fewer than 2 points", () => {
    expect(linearRegression([1], [1])).toBeNull();
    expect(linearRegression([], [])).toBeNull();
  });
});

// ─────────────────────────────────────────────
//  detectTrends
// ─────────────────────────────────────────────

describe("detectTrends", () => {
  const risingData = {
    headers: ["Sales"],
    rows: [[100], [110], [120], [130], [140], [150], [160]],
  };

  const fallingData = {
    headers: ["Errors"],
    rows: [[50], [45], [40], [35], [30], [25], [20]],
  };

  test("detects upward trend", () => {
    const trends = detectTrends(risingData);
    expect(trends.length).toBe(1);
    expect(trends[0].direction).toBe("upward");
  });

  test("detects downward trend", () => {
    const trends = detectTrends(fallingData);
    expect(trends.length).toBe(1);
    expect(trends[0].direction).toBe("downward");
  });

  test("upward trend has positive slope", () => {
    const trends = detectTrends(risingData);
    expect(trends[0].slope).toBeGreaterThan(0);
  });

  test("strong trend has high R²", () => {
    const trends = detectTrends(risingData);
    expect(trends[0].r2).toBeGreaterThan(0.9);
  });

  test("skips text columns", () => {
    const textData = {
      headers: ["Label"],
      rows: [["a"], ["b"], ["c"], ["d"]],
    };
    expect(detectTrends(textData)).toHaveLength(0);
  });
});

// ─────────────────────────────────────────────
//  Report formatters
// ─────────────────────────────────────────────

describe("formatAnalysisReport", () => {
  test("returns no-data message for empty summaries", () => {
    expect(formatAnalysisReport([])).toBe("No data to analyse.");
  });

  test("includes column header in report", () => {
    const data = {
      headers: ["Revenue"],
      rows: [[100], [200], [300]],
    };
    const summaries = analyseDataset(data);
    const report = formatAnalysisReport(summaries);
    expect(report).toContain("Revenue");
  });
});

describe("formatAnomalyReport", () => {
  test("returns no anomalies message for empty list", () => {
    const report = formatAnomalyReport([]);
    expect(report).toContain("No anomalies");
  });

  test("includes anomaly info in report", () => {
    const data = {
      headers: ["Value"],
      rows: [[10], [12], [11], [500], [10], [11], [12], [11], [10], [12]],
    };
    const anomalies = detectAnomalies(data);
    const report = formatAnomalyReport(anomalies);
    expect(report).toContain("Anomaly");
  });
});

describe("formatTrendReport", () => {
  test("returns message when no numeric columns", () => {
    const report = formatTrendReport([]);
    expect(report).toContain("No numeric");
  });

  test("includes direction in trend report", () => {
    const data = {
      headers: ["Price"],
      rows: [[10], [20], [30], [40], [50]],
    };
    const trends = detectTrends(data);
    const report = formatTrendReport(trends);
    expect(report.toLowerCase()).toContain("upward");
  });
});
