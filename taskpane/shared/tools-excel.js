/* global Excel */
//
// Excel tool implementations + Excel-specific helpers.
//
// Each `toolExcel*` function is invoked from the dispatcher in taskpane.js
// when a matching `excel_*` tool_call arrives from the daemon. All operations
// go through Excel.run for proper context lifecycle. Address strings use
// A1 notation, optionally sheet-qualified (e.g. "Sheet1!A1:C5").

function _splitSheetAddress(address, fallbackSheetName) {
  // Returns { sheetName, a1 } from "Sheet1!A1:B2" or "A1:B2" (uses fallback).
  if (!address) return { sheetName: fallbackSheetName, a1: null };
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(address);
  if (m) return { sheetName: m[1] || m[2], a1: m[3] };
  return { sheetName: fallbackSheetName, a1: address };
}

// Parse the top-left cell of an A1-notation range into 1-based row/col.
// "C5" → { row: 5, col: 3 }; "AA10" → { row: 10, col: 27 }; "C5:F10" → uses C5.
// Used so excel_find_value reports absolute spreadsheet coordinates, not
// 0-based offsets within whatever used-range happens to start at.
function _topLeftCell(a1OrRange) {
  const a1 = String(a1OrRange).split(":")[0];
  const m = /^([A-Za-z]+)(\d+)$/.exec(a1);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10), col };
}

// Spreadsheet column number → A1 letters. 1→"A", 27→"AA", etc.
function _colNumberToLetters(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function _activeSheetName(context) {
  const ws = context.workbook.worksheets.getActiveWorksheet();
  ws.load("name");
  await context.sync();
  return ws.name;
}

export async function toolExcelGetSelectedRange() {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, values, rowCount, columnCount, worksheet/name");
    await context.sync();
    return {
      address: range.address,
      sheet: range.worksheet.name,
      row_count: range.rowCount,
      column_count: range.columnCount,
      values: range.values,
    };
  });
}

export async function toolExcelListSheets() {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name, items/position");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await context.sync();
    // Load used range per sheet (separate sync because we needed names first).
    const used = sheets.items.map((s) => {
      const r = s.getUsedRangeOrNullObject(true);
      r.load("address, rowCount, columnCount, isNullObject");
      return { sheet: s, used: r };
    });
    await context.sync();
    return {
      active_sheet: active.name,
      sheets: used.map(({ sheet, used: u }) => ({
        name: sheet.name,
        position: sheet.position,
        used_range: u.isNullObject ? null : u.address,
        row_count: u.isNullObject ? 0 : u.rowCount,
        column_count: u.isNullObject ? 0 : u.columnCount,
      })),
    };
  });
}

export async function toolExcelReadRange({ address = null, sheet = null }) {
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    // For whole-sheet reads (no explicit address), use the null-object
    // variant so an empty sheet returns gracefully instead of throwing
    // an InvalidArgument from Excel.js.
    const range = a1 ? ws.getRange(a1) : ws.getUsedRangeOrNullObject(true);
    range.load("address, values, rowCount, columnCount, isNullObject");
    await context.sync();
    if (!a1 && range.isNullObject) {
      return {
        sheet: sheetName,
        address: null,
        row_count: 0,
        column_count: 0,
        values: [],
        empty: true,
      };
    }
    return {
      sheet: sheetName,
      address: range.address,
      row_count: range.rowCount,
      column_count: range.columnCount,
      values: range.values,
    };
  });
}

export async function toolExcelWriteRange({ address, values, sheet = null }) {
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) {
    throw new Error("`values` must be a non-empty 2D array.");
  }
  // Ragged-array guard: every row must have the same length as values[0],
  // otherwise Excel.js fails with a confusing internal error (or silently
  // pads with undefined on some builds). Surface the offending row index.
  const expectedCols = values[0].length;
  for (let i = 1; i < values.length; i++) {
    if (!Array.isArray(values[i]) || values[i].length !== expectedCols) {
      const got = Array.isArray(values[i]) ? `${values[i].length} elements` : "not an array";
      throw new Error(
        `Ragged values: row ${i} is ${got} but row 0 has ${expectedCols}. ` +
          `All rows must be the same length.`,
      );
    }
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    if (!a1) throw new Error("`address` is required for write_range.");
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("rowCount, columnCount, address");
    await context.sync();
    if (range.rowCount !== values.length || range.columnCount !== expectedCols) {
      throw new Error(
        `Shape mismatch: range ${range.address} is ${range.rowCount}×${range.columnCount} ` +
          `but values is ${values.length}×${expectedCols}.`,
      );
    }
    range.values = values;
    await context.sync();
    return { sheet: sheetName, address: range.address, written: values.length * expectedCols };
  });
}

export async function toolExcelFindValue({
  query,
  sheet = null,
  match_case = false,
  whole_cell = false,
}) {
  if (typeof query !== "string" || !query.length) throw new Error("`query` is required.");
  return await Excel.run(async (context) => {
    const targetSheets = sheet
      ? [context.workbook.worksheets.getItem(sheet)]
      : (() => {
          const all = context.workbook.worksheets;
          all.load("items/name");
          return all;
        })();
    if (Array.isArray(targetSheets)) {
      // single sheet case — load name for the result
      targetSheets[0].load("name");
    }
    await context.sync();
    const sheetsToScan = Array.isArray(targetSheets) ? targetSheets : targetSheets.items;
    const matches = [];
    for (const ws of sheetsToScan) {
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("address, values, rowCount, columnCount, isNullObject");
      await context.sync();
      if (used.isNullObject) continue;
      // The used range may start anywhere on the sheet — e.g. "Sheet1!C5:F10".
      // Translate per-cell offsets within it into absolute spreadsheet rows
      // and columns so the agent's row/column numbers match what the user
      // sees in Excel's row/column headers.
      const origin = _topLeftCell(used.address) || { row: 1, col: 1 };
      const q = match_case ? query : query.toLowerCase();
      for (let r = 0; r < used.rowCount; r++) {
        for (let c = 0; c < used.columnCount; c++) {
          const v = used.values[r][c];
          if (v === null || v === undefined || v === "") continue;
          const s = match_case ? String(v) : String(v).toLowerCase();
          const hit = whole_cell ? s === q : s.includes(q);
          if (hit) {
            const absRow = origin.row + r;
            const absCol = origin.col + c;
            matches.push({
              sheet: ws.name,
              row: absRow,
              column: absCol,
              address: `${_colNumberToLetters(absCol)}${absRow}`,
              value: v,
            });
          }
        }
      }
    }
    return { query, match_count: matches.length, matches };
  });
}

export async function toolExcelInsertRows({ sheet = null, at, count = 1 }) {
  if (!Number.isInteger(at) || at < 1) throw new Error("`at` must be a positive integer.");
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    const target = ws.getRange(`${at}:${at + count - 1}`);
    target.insert(Excel.InsertShiftDirection.down);
    await context.sync();
    return { sheet: activeName, inserted_at: at, count };
  });
}

export async function toolExcelDeleteRows({ sheet = null, at, count = 1 }) {
  if (!Number.isInteger(at) || at < 1) throw new Error("`at` must be a positive integer.");
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    const target = ws.getRange(`${at}:${at + count - 1}`);
    target.delete(Excel.DeleteShiftDirection.up);
    await context.sync();
    return { sheet: activeName, deleted_at: at, count };
  });
}

export async function toolExcelSelectRange({ address, sheet = null }) {
  if (!address || typeof address !== "string") {
    throw new Error("`address` (an A1 range like 'B4' or 'A1:D9') is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    ws.activate(); // make sure the selection is visible on the right sheet
    const range = ws.getRange(address);
    range.select();
    range.load("address");
    await context.sync();
    return { sheet: activeName, address: range.address, selected: true };
  });
}

// ---------------------------------------------------------------------------
// Tier 1 editing tools — formulas / formatting / columns / sheets / clear
// ---------------------------------------------------------------------------

// Tool: excel_write_formula — like write_range but sets formulas (write_range
// is values-only). Each cell gets an A1/R1C1 formula string, e.g. "=SUM(A1:A9)".
export async function toolExcelWriteFormula({ address, formulas, sheet = null }) {
  if (!Array.isArray(formulas) || !formulas.length || !Array.isArray(formulas[0])) {
    throw new Error("`formulas` must be a non-empty 2D array.");
  }
  const cols = formulas[0].length;
  for (let i = 1; i < formulas.length; i++) {
    if (!Array.isArray(formulas[i]) || formulas[i].length !== cols) {
      throw new Error(
        `Ragged formulas: row ${i} has ${formulas[i]?.length} but row 0 has ${cols}.`,
      );
    }
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    if (!a1) throw new Error("`address` is required for write_formula.");
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("rowCount, columnCount, address");
    await context.sync();
    if (range.rowCount !== formulas.length || range.columnCount !== cols) {
      throw new Error(
        `Shape mismatch: range ${range.address} is ${range.rowCount}×${range.columnCount} ` +
          `but formulas is ${formulas.length}×${cols}.`,
      );
    }
    range.formulas = formulas;
    await context.sync();
    return { sheet: sheetName, address: range.address, written: formulas.length * cols };
  });
}

// Tool: excel_set_format — number format / font / fill / borders on a range.
export async function toolExcelSetFormat({
  address,
  sheet = null,
  number_format,
  bold,
  italic,
  font_size,
  font_name,
  font_color,
  fill_color,
  border,
}) {
  if (!address || typeof address !== "string") {
    throw new Error("`address` (an A1 range) is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("rowCount, columnCount, address");
    await context.sync();
    if (number_format !== undefined) {
      // numberFormat expects a 2D array matching the range shape.
      range.numberFormat = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => number_format),
      );
    }
    if (bold !== undefined) range.format.font.bold = bold;
    if (italic !== undefined) range.format.font.italic = italic;
    if (font_size !== undefined) range.format.font.size = font_size;
    if (font_name !== undefined) range.format.font.name = font_name;
    if (font_color !== undefined) range.format.font.color = font_color;
    if (fill_color !== undefined) range.format.fill.color = fill_color;
    if (border) {
      for (const edge of [
        "EdgeTop",
        "EdgeBottom",
        "EdgeLeft",
        "EdgeRight",
        "InsideHorizontal",
        "InsideVertical",
      ]) {
        const b = range.format.borders.getItem(edge);
        b.style = "Continuous";
        b.weight = "Thin";
      }
    }
    await context.sync();
    return { sheet: sheetName, address: range.address, formatted: true };
  });
}

// Column-letter range helper, e.g. ("C", 2) → "C:D".
function _colSpan(at, count) {
  const start = String(at)
    .toUpperCase()
    .replace(/[^A-Z]/g, "");
  if (!start) throw new Error("`at` must be a column letter, e.g. 'C'.");
  let n = 0;
  for (const ch of start) n = n * 26 + (ch.charCodeAt(0) - 64);
  return `${start}:${_colNumberToLetters(n + count - 1)}`;
}

// Tool: excel_insert_columns — insert blank columns, shifting right.
export async function toolExcelInsertColumns({ sheet = null, at, count = 1 }) {
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    ws.getRange(_colSpan(at, count)).insert(Excel.InsertShiftDirection.right);
    await context.sync();
    return { sheet: activeName, inserted_at: String(at).toUpperCase(), count };
  });
}

// Tool: excel_delete_columns — delete columns, shifting left.
export async function toolExcelDeleteColumns({ sheet = null, at, count = 1 }) {
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    ws.getRange(_colSpan(at, count)).delete(Excel.DeleteShiftDirection.left);
    await context.sync();
    return { sheet: activeName, deleted_at: String(at).toUpperCase(), count };
  });
}

// Tool: excel_add_sheet — add a worksheet (optionally at a 0-based position).
export async function toolExcelAddSheet({ name, position = null }) {
  if (!name || typeof name !== "string") throw new Error("`name` is required.");
  return await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.add(name);
    if (Number.isInteger(position)) ws.position = position;
    ws.load("name, position");
    await context.sync();
    return { name: ws.name, position: ws.position, added: true };
  });
}

// Tool: excel_delete_sheet — delete a worksheet by name.
export async function toolExcelDeleteSheet({ name }) {
  if (!name || typeof name !== "string") throw new Error("`name` is required.");
  return await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(name);
    ws.delete();
    await context.sync();
    return { name, deleted: true };
  });
}

// Tool: excel_rename_sheet — rename a worksheet.
export async function toolExcelRenameSheet({ name, new_name }) {
  if (!name || !new_name) throw new Error("`name` and `new_name` are required.");
  return await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(name);
    ws.name = new_name;
    await context.sync();
    return { from: name, to: new_name, renamed: true };
  });
}

// Tool: excel_clear_range — clear contents, formats, or both.
export async function toolExcelClearRange({ address, sheet = null, what = "contents" }) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  const APPLY = { contents: "Contents", formats: "Formats", all: "All" };
  if (!APPLY[what]) throw new Error(`\`what\` must be one of: ${Object.keys(APPLY).join(", ")}`);
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.clear(Excel.ClearApplyTo[what]);
    range.load("address");
    await context.sync();
    return { sheet: sheetName, address: range.address, cleared: what };
  });
}

// ---------------------------------------------------------------------------
// Tier 2 editing tools — sort / filter / tables / charts / dimensions
// ---------------------------------------------------------------------------

// Tool: excel_sort_range — sort a range by one column.
export async function toolExcelSortRange({
  address,
  sheet = null,
  key = 0,
  ascending = true,
  has_headers = false,
}) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  if (!Number.isInteger(key) || key < 0) throw new Error("`key` must be a 0-based column index.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.sort.apply([{ key, ascending: !!ascending }], false, !!has_headers);
    range.load("address");
    await context.sync();
    return {
      sheet: sheetName,
      address: range.address,
      sorted_by_column: key,
      ascending: !!ascending,
    };
  });
}

// Tool: excel_autofilter — apply an AutoFilter to a range, or clear it.
export async function toolExcelAutoFilter({ address, sheet = null, clear = false }) {
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    if (clear) {
      ws.autoFilter.remove();
      await context.sync();
      return { sheet: activeName, autofilter: "removed" };
    }
    if (!address || typeof address !== "string") {
      throw new Error("`address` is required (or pass clear: true to remove the filter).");
    }
    const { a1 } = _splitSheetAddress(address, activeName);
    ws.autoFilter.apply(ws.getRange(a1));
    await context.sync();
    return { sheet: activeName, autofilter: "applied", address: a1 };
  });
}

// Tool: excel_create_table — turn a range into a named Excel table.
export async function toolExcelCreateTable({
  address,
  sheet = null,
  has_headers = true,
  name = null,
}) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const table = ws.tables.add(a1, !!has_headers);
    if (name) table.name = name;
    table.load("name");
    await context.sync();
    return { sheet: sheetName, table: table.name, range: a1 };
  });
}

// Tool: excel_add_table_rows — append rows to an existing table.
export async function toolExcelAddTableRows({ table, values, index = null }) {
  if (!table || typeof table !== "string") throw new Error("`table` (table name) is required.");
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) {
    throw new Error("`values` must be a non-empty 2D array.");
  }
  return await Excel.run(async (context) => {
    const t = context.workbook.tables.getItem(table);
    t.rows.add(Number.isInteger(index) ? index : null, values);
    await context.sync();
    return { table, added_rows: values.length };
  });
}

const CHART_TYPE = {
  column: "ColumnClustered",
  bar: "BarClustered",
  line: "Line",
  pie: "Pie",
  scatter: "XYScatter",
  area: "Area",
};

// Tool: excel_create_chart — add a chart from a data range.
export async function toolExcelCreateChart({
  data_address,
  sheet = null,
  chart_type = "column",
  title = null,
}) {
  if (!data_address || typeof data_address !== "string") {
    throw new Error("`data_address` is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(data_address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const type = CHART_TYPE[String(chart_type).toLowerCase()] || chart_type;
    const chart = ws.charts.add(type, ws.getRange(a1), Excel.ChartSeriesBy.auto);
    if (title) {
      chart.title.text = title;
      chart.title.visible = true;
    }
    chart.load("name");
    await context.sync();
    return { sheet: sheetName, chart: chart.name, type, data: a1 };
  });
}

// Tool: excel_set_column_width — set a fixed width, or autofit.
export async function toolExcelSetColumnWidth({
  address,
  sheet = null,
  width = null,
  autofit = false,
}) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const range = context.workbook.worksheets.getItem(sheetName).getRange(a1);
    if (autofit) range.format.autofitColumns();
    else if (typeof width === "number") range.format.columnWidth = width;
    else throw new Error("Provide `width` (points) or `autofit: true`.");
    await context.sync();
    return { sheet: sheetName, address: a1, autofit: !!autofit, width: autofit ? null : width };
  });
}

// Tool: excel_set_row_height — set a fixed height, or autofit.
export async function toolExcelSetRowHeight({
  address,
  sheet = null,
  height = null,
  autofit = false,
}) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const range = context.workbook.worksheets.getItem(sheetName).getRange(a1);
    if (autofit) range.format.autofitRows();
    else if (typeof height === "number") range.format.rowHeight = height;
    else throw new Error("Provide `height` (points) or `autofit: true`.");
    await context.sync();
    return { sheet: sheetName, address: a1, autofit: !!autofit, height: autofit ? null : height };
  });
}
