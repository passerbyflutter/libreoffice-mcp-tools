import type { CellRange } from '../types.js';

/**
 * Convert a cell range to a markdown table string.
 */
export function rangeToMarkdownTable(range: CellRange): string {
  if (range.rows.length === 0) return '*(empty range)*';

  const allRows = range.rows;
  const headers = range.headers ?? allRows[0]?.map((_, i) => columnIndexToLetter(range.startCol + i)) ?? [];
  const dataRows = range.headers ? allRows.slice(1) : allRows;

  const colWidths = headers.map((h, i) => {
    const maxDataWidth = Math.max(0, ...dataRows.map(r => String(r[i] ?? '').length));
    return Math.max(String(h).length, maxDataWidth, 3);
  });

  const headerRow = '| ' + headers.map((h, i) => String(h).padEnd(colWidths[i]!)).join(' | ') + ' |';
  const separator = '| ' + colWidths.map(w => '-'.repeat(w)).join(' | ') + ' |';
  const dataRowStrings = dataRows.map(
    row => '| ' + headers.map((_, i) => String(row[i] ?? '').padEnd(colWidths[i]!)).join(' | ') + ' |',
  );

  return [headerRow, separator, ...dataRowStrings].join('\n');
}

function columnIndexToLetter(index: number): string {
  let result = '';
  let n = index;
  do {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return result;
}
