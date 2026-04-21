/**
 * XLSX/CSV/ODS parser using ExcelJS (MIT, actively maintained).
 * Handles .xlsx, .xlsm, .csv files. Legacy .xls and .ods are bridged via LibreOffice.
 *
 * ExcelJS uses 1-based row/col indexing internally; CellRange output uses 0-based for
 * API compatibility with existing tool consumers.
 */
import { createRequire } from 'node:module';
import type { SheetInfo, CellRange } from '../types.js';

// ExcelJS ships as CommonJS — use createRequire for reliable ESM interop in Node.js 22+
const _require = createRequire(import.meta.url);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const ExcelJSLib = _require('exceljs') as any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type ExcelJSWorkbook = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type ExcelJSWorksheet = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type ExcelJSCell = any;

export interface XlsxWorkbook {
  sheetNames: string[];
  raw: ExcelJSWorkbook;
}

// ── Cell address helpers ──────────────────────────────────────────────────────

/**
 * Parse an A1-style cell address into 1-based {row, col}.
 * Example: "A1" → {row:1, col:1}, "B3" → {row:3, col:2}
 */
function parseCellAddress(addr: string): { row: number; col: number } {
  const match = addr.toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell address: "${addr}". Use A1 format.`);
  const colStr = match[1]!;
  let col = 0;
  for (const ch of colStr) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(match[2]!, 10), col };
}

/** Encode 1-based {row, col} to "A1"-style address. */
function encodeCell(row: number, col: number): string {
  let colStr = '';
  let c = col;
  while (c > 0) {
    colStr = String.fromCharCode(64 + (c % 26 || 26)) + colStr;
    c = Math.floor((c - 1) / 26);
  }
  return `${colStr}${row}`;
}

/**
 * Parse "A1:C10" into 1-based row/col bounds.
 * If only a single cell address is given, start === end.
 */
function parseRange(rangeStr: string): { sr: number; er: number; sc: number; ec: number } {
  const parts = rangeStr.split(':');
  if (parts.length === 1) {
    const { row, col } = parseCellAddress(rangeStr);
    return { sr: row, er: row, sc: col, ec: col };
  }
  const start = parseCellAddress(parts[0]!);
  const end = parseCellAddress(parts[1]!);
  return { sr: start.row, er: end.row, sc: start.col, ec: end.col };
}

/** Extract a scalar value from an ExcelJS cell (handles formula/rich-text/dates). */
function getCellValue(cell: ExcelJSCell): string | number | boolean | null {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    // Formula cell: { formula, result }
    if ('result' in v) {
      const r = v.result;
      if (r instanceof Date) return r.toISOString();
      if (typeof r === 'string' || typeof r === 'number' || typeof r === 'boolean') return r;
      return null;
    }
    // Hyperlink: { text, hyperlink }
    if ('text' in v && typeof v.text === 'string') return v.text;
    // Rich text: { richText: [{text}] }
    if ('richText' in v && Array.isArray(v.richText)) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return v.richText.map((r: any) => String(r.text ?? '')).join('');
    }
    // Error value: { error }
    if ('error' in v) return null;
  }
  return String(v);
}

// ── Public API ────────────────────────────────────────────────────────────────

export async function parseXlsx(filePath: string): Promise<XlsxWorkbook> {
  const wb: ExcelJSWorkbook = new ExcelJSLib.Workbook();
  if (filePath.toLowerCase().endsWith('.csv')) {
    await wb.csv.readFile(filePath);
  } else {
    await wb.xlsx.readFile(filePath);
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheetNames: string[] = wb.worksheets.map((ws: any) => ws.name as string);
  return { sheetNames, raw: wb };
}

export function getSheetInfo(workbook: XlsxWorkbook): SheetInfo[] {
  return workbook.sheetNames.map((name, index) => {
    const ws: ExcelJSWorksheet = workbook.raw.getWorksheet(name);
    if (!ws) return { name, index, rowCount: 0, colCount: 0 };
    return {
      name,
      index,
      rowCount: ws.actualRowCount as number,
      colCount: ws.actualColumnCount as number,
    };
  });
}

export function getRange(
  workbook: XlsxWorkbook,
  sheetName: string,
  rangeStr?: string,
  limit = 100,
  offset = 0,
): CellRange {
  const ws: ExcelJSWorksheet = workbook.raw.getWorksheet(sheetName);
  if (!ws) {
    throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${workbook.sheetNames.join(', ')}`);
  }

  const totalRows: number = ws.actualRowCount;
  const totalCols: number = ws.actualColumnCount;
  if (totalRows === 0 || totalCols === 0) {
    return { rows: [], startRow: 0, startCol: 0, endRow: 0, endCol: 0, sheetName };
  }

  // All bounds are 1-based for ExcelJS
  let sr: number, er: number, sc: number, ec: number;
  if (rangeStr) {
    try {
      ({ sr, er, sc, ec } = parseRange(rangeStr));
    } catch {
      throw new Error(`Invalid range: "${rangeStr}". Use A1:C10 format.`);
    }
  } else {
    // Paginate across all rows
    sr = 1 + offset;
    er = Math.min(totalRows, sr + limit - 1);
    sc = 1;
    ec = totalCols;
  }

  const rows: (string | number | boolean | null)[][] = [];

  for (let r = sr; r <= er; r++) {
    const row: (string | number | boolean | null)[] = [];
    for (let c = sc; c <= ec; c++) {
      row.push(getCellValue(ws.getCell(r, c)));
    }
    rows.push(row);
  }

  // Detect header row (first row of sheet, all strings or null)
  let headers: string[] | undefined;
  if (rows.length > 0 && sr === 1) {
    const firstRow = rows[0]!;
    const isHeaderRow = firstRow.every(cell => typeof cell === 'string' || cell === null);
    if (isHeaderRow) {
      headers = firstRow.map(cell => String(cell ?? ''));
    }
  }

  return {
    headers,
    rows,
    startRow: sr - 1, // convert to 0-based for API compatibility
    startCol: sc - 1,
    endRow: er - 1,
    endCol: ec - 1,
    sheetName,
  };
}

export function getFormulas(
  workbook: XlsxWorkbook,
  sheetName: string,
  rangeStr?: string,
): Array<{ cell: string; formula: string; value: unknown }> {
  const ws: ExcelJSWorksheet = workbook.raw.getWorksheet(sheetName);
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);

  const totalRows: number = ws.actualRowCount;
  const totalCols: number = ws.actualColumnCount;
  if (totalRows === 0) return [];

  const { sr, er, sc, ec } = rangeStr
    ? parseRange(rangeStr)
    : { sr: 1, er: totalRows, sc: 1, ec: totalCols };

  const results: Array<{ cell: string; formula: string; value: unknown }> = [];

  for (let r = sr; r <= er; r++) {
    for (let c = sc; c <= ec; c++) {
      const cell: ExcelJSCell = ws.getCell(r, c);
      const v = cell.value;
      // ExcelJS formula cell: { formula, result }
      if (v !== null && typeof v === 'object' && 'formula' in v && typeof v.formula === 'string') {
        results.push({ cell: encodeCell(r, c), formula: `=${v.formula}`, value: v.result ?? null });
      }
    }
  }

  return results;
}

export function setCellValue(
  workbook: XlsxWorkbook,
  sheetName: string,
  cellAddr: string,
  value: string | number | boolean | null,
  formula?: string,
): void {
  const ws: ExcelJSWorksheet = workbook.raw.getWorksheet(sheetName);
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);

  const cell: ExcelJSCell = ws.getCell(cellAddr);
  if (formula) {
    cell.value = { formula: formula.replace(/^=/, ''), result: value };
  } else {
    cell.value = value;
  }
}

export async function saveXlsx(workbook: XlsxWorkbook, filePath: string): Promise<void> {
  if (filePath.toLowerCase().endsWith('.csv')) {
    await workbook.raw.csv.writeFile(filePath);
  } else {
    await workbook.raw.xlsx.writeFile(filePath);
  }
}

/**
 * Set multiple cells starting from `startCell` (A1 format) with the given 2D values array.
 */
export function setRangeValues(
  workbook: XlsxWorkbook,
  sheetName: string,
  startCell: string,
  values: (string | number | boolean | null)[][],
): void {
  const ws: ExcelJSWorksheet = workbook.raw.getWorksheet(sheetName);
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);

  const { row: startRow, col: startCol } = parseCellAddress(startCell);
  for (let r = 0; r < values.length; r++) {
    const rowData = values[r]!;
    for (let c = 0; c < rowData.length; c++) {
      ws.getCell(startRow + r, startCol + c).value = rowData[c] ?? null;
    }
  }
}

/**
 * Add a new worksheet to the workbook.
 */
export function addNewSheet(workbook: XlsxWorkbook, sheetName: string, headers?: string[]): void {
  const ws: ExcelJSWorksheet = workbook.raw.addWorksheet(sheetName);
  if (headers && headers.length > 0) {
    headers.forEach((h, i) => ws.getCell(1, i + 1).value = h);
  }
  workbook.sheetNames.push(sheetName);
}

/**
 * Create a new empty XLSX workbook file with one blank sheet.
 */
export async function createEmptyXlsx(filePath: string): Promise<void> {
  const wb: ExcelJSWorkbook = new ExcelJSLib.Workbook();
  wb.addWorksheet('Sheet1');
  await wb.xlsx.writeFile(filePath);
}

/**
 * Search all cells across all sheets for a query string.
 */
export function searchCells(
  workbook: XlsxWorkbook,
  query: string,
  limit: number,
): Array<{ sheet: string; cell: string; value: string }> {
  const lowerQuery = query.toLowerCase();
  const results: Array<{ sheet: string; cell: string; value: string }> = [];

  for (const sheetName of workbook.sheetNames) {
    if (results.length >= limit) break;
    const ws: ExcelJSWorksheet = workbook.raw.getWorksheet(sheetName);
    if (!ws) continue;

    const totalRows: number = ws.actualRowCount;
    const totalCols: number = ws.actualColumnCount;

    for (let r = 1; r <= totalRows && results.length < limit; r++) {
      for (let c = 1; c <= totalCols && results.length < limit; c++) {
        const v = getCellValue(ws.getCell(r, c));
        if (v !== null && String(v).toLowerCase().includes(lowerQuery)) {
          results.push({ sheet: sheetName, cell: encodeCell(r, c), value: String(v) });
        }
      }
    }
  }

  return results;
}

export function rangeToMarkdownTable(range: CellRange): string {
  if (range.rows.length === 0) return '*(empty)*';

  const allRows = range.rows;
  const headers = range.headers ?? allRows[0]?.map((_, i) => `Col${i + 1}`) ?? [];
  const dataRows = range.headers ? allRows.slice(1) : allRows;

  const colWidths = headers.map((h, i) => {
    const maxDataWidth = Math.max(...dataRows.map(r => String(r[i] ?? '').length));
    return Math.max(String(h).length, maxDataWidth, 3);
  });

  const headerRow = '| ' + headers.map((h, i) => String(h).padEnd(colWidths[i]!)).join(' | ') + ' |';
  const separator = '| ' + colWidths.map(w => '-'.repeat(w)).join(' | ') + ' |';
  const dataRowStrings = dataRows.map(
    row => '| ' + headers.map((_, i) => String(row[i] ?? '').padEnd(colWidths[i]!)).join(' | ') + ' |',
  );

  return [headerRow, separator, ...dataRowStrings].join('\n');
}
