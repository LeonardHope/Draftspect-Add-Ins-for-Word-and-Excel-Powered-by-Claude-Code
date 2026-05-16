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
