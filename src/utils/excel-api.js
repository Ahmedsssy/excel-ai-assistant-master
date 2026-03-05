/**
 * excel-api.js – Excel API utility layer
 *
 * Provides helpers that wrap Office.js Excel API calls used by the add-in.
 * All functions return plain JS objects so they can be serialised to JSON
 * and sent to the AI or the Python backend without Office context.
 */

"use strict";

/* ─────────────────────────────────────────────
   1.  Reading data from the workbook
───────────────────────────────────────────── */

/**
 * Read the currently selected range.
 * Returns { headers, rows, address, rowCount, colCount }
 */
export async function readSelection() {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load(["values", "address", "rowCount", "columnCount"]);
    await ctx.sync();

    const values = range.values;
    if (!values || values.length === 0) return null;

    const headers = values[0].map((h) => (h !== null && h !== "" ? String(h) : null));
    const hasHeaders = headers.some((h) => h && isNaN(h));

    return {
      address: range.address,
      rowCount: range.rowCount,
      colCount: range.columnCount,
      headers: hasHeaders ? headers : null,
      rows: hasHeaders ? values.slice(1) : values,
      rawValues: values,
    };
  });
}

/**
 * Read all used-range data from the active sheet.
 */
export async function readUsedRange() {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load(["values", "address", "rowCount", "columnCount"]);
    await ctx.sync();

    if (!range.values || range.values.length === 0) return null;

    const values = range.values;
    const headers = values[0].map((h) => (h !== null && h !== "" ? String(h) : null));
    const hasHeaders = headers.some((h) => h && isNaN(h));

    return {
      address: range.address,
      rowCount: range.rowCount,
      colCount: range.columnCount,
      headers: hasHeaders ? headers : null,
      rows: hasHeaders ? values.slice(1) : values,
      rawValues: values,
    };
  });
}

/* ─────────────────────────────────────────────
   2.  Writing data / formatting
───────────────────────────────────────────── */

/**
 * Highlight specific cells in the active sheet with a given fill colour.
 * @param {string[][]} addresses  Array of A1-style cell addresses
 * @param {string}     color      Hex colour, e.g. "#FF4444"
 */
export async function highlightCells(addresses, color) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    for (const addr of addresses) {
      const cell = sheet.getRange(addr);
      cell.format.fill.color = color;
    }
    await ctx.sync();
  });
}

/**
 * Clear fill colour from an array of addresses.
 */
export async function clearHighlights(addresses) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    for (const addr of addresses) {
      const cell = sheet.getRange(addr);
      cell.format.fill.clear();
    }
    await ctx.sync();
  });
}

/**
 * Write a 2-D array of values starting at a given address.
 */
export async function writeValues(address, values2d) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.values = values2d;
    await ctx.sync();
  });
}

/* ─────────────────────────────────────────────
   3.  Charting
───────────────────────────────────────────── */

/**
 * Create a chart in the active sheet.
 * @param {string}  rangeAddress   Source data range (A1 notation)
 * @param {string}  chartType      Excel.ChartType string value, e.g. "ColumnClustered"
 * @param {string}  title          Chart title
 */
export async function createChart(rangeAddress, chartType, title = "Chart") {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const dataRange = sheet.getRange(rangeAddress);

    const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
    chart.title.text = title;
    chart.legend.position = Excel.ChartLegendPosition.bottom;
    chart.legend.visible = true;

    // Position chart below the data (chart height ≈ 18 rows)
    const dataEndRow = parseInt((rangeAddress.match(/\d+/) || ["1"])[0], 10);
    const chartTopRow = dataEndRow + 2;
    chart.setPosition(`A${chartTopRow}`, `N${chartTopRow + 18}`);

    await ctx.sync();
    return { success: true, chartType, title };
  });
}

/* ─────────────────────────────────────────────
   4.  Helpers
───────────────────────────────────────────── */

/**
 * Convert a column index (0-based) to Excel column letters.
 */
export function colIndexToLetter(idx) {
  let letter = "";
  let n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

/**
 * Build an A1 cell address from row / col (0-based).
 */
export function toA1(row, col) {
  return `${colIndexToLetter(col)}${row + 1}`;
}

/**
 * Summarise data for context injection into the AI prompt.
 * Keeps it compact so we don't exceed token limits.
 */
export function summariseData(data, maxRows = 50) {
  if (!data) return "No data available.";

  const lines = [];
  if (data.address) lines.push(`Range: ${data.address}`);
  lines.push(`Dimensions: ${data.rowCount} rows × ${data.colCount} columns`);

  if (data.headers) {
    lines.push(`Headers: ${data.headers.join(", ")}`);
  }

  const preview = (data.rows || data.rawValues || []).slice(0, maxRows);
  if (preview.length > 0) {
    lines.push(`\nFirst ${Math.min(preview.length, maxRows)} data rows:`);
    preview.forEach((row, i) => {
      lines.push(`  Row ${i + 1}: ${row.join("\t")}`);
    });
    if ((data.rows?.length || 0) > maxRows) {
      lines.push(`  … (${data.rows.length - maxRows} more rows not shown)`);
    }
  }

  return lines.join("\n");
}
