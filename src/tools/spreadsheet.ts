import { z } from 'zod';
import { defineTool, docIdSchema, paginationSchema, sheetNameSchema } from './ToolDefinition.js';
import { DocumentCategory } from './categories.js';
import { parseXlsx, getSheetInfo, getRange, getFormulas, setCellValue, saveXlsx, setRangeValues, addNewSheet } from '../parsers/XlsxParser.js';
import { rangeToMarkdownTable } from '../formatters/TableFormatter.js';

export const spreadsheetListSheets = defineTool({
  name: 'spreadsheet_list_sheets',
  description: 'List all sheets in a spreadsheet with their names, row counts, and column counts.',
  annotations: {
    category: DocumentCategory.SPREADSHEET,
    readOnlyHint: true,
    title: 'List Sheets',
  },
  schema: {
    ...docIdSchema,
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'calc') {
      throw new Error('spreadsheet_list_sheets only works with spreadsheet documents.');
    }
    const wb = await session.getOrParseXlsx(() => parseXlsx(session.parsedPath));
    const sheets = getSheetInfo(wb);
    response.appendText(`${sheets.length} sheet(s):`);
    response.attachJson(sheets.map(s => ({ name: s.name, rows: s.rowCount, cols: s.colCount })));
  },
});

export const spreadsheetGetRange = defineTool({
  name: 'spreadsheet_get_range',
  description: `Read a range of cells from a spreadsheet sheet.
Returns data as structured JSON and a markdown table.
Use range like "A1:D10" for specific cells, or omit for paginated full sheet.
Token-efficient: specify a range rather than reading the whole sheet.`,
  annotations: {
    category: DocumentCategory.SPREADSHEET,
    readOnlyHint: true,
    title: 'Get Cell Range',
  },
  schema: {
    ...docIdSchema,
    ...sheetNameSchema,
    range: z.string().optional().describe('Cell range in A1:C10 format. If omitted, returns paginated rows.'),
    ...paginationSchema,
    format: z.enum(['json', 'table', 'both']).optional().default('both').describe('Output format: json, table (markdown), or both'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'calc') {
      throw new Error('spreadsheet_get_range only works with spreadsheet documents.');
    }
    const wb = await session.getOrParseXlsx(() => parseXlsx(session.parsedPath));
    const sheetName = request.params.sheetName ?? wb.sheetNames[0]!;
    const cellRange = getRange(wb, sheetName, request.params.range, request.params.limit, request.params.offset);

    response.appendText(`Sheet "${sheetName}" — ${cellRange.rows.length} row(s), ${(cellRange.rows[0] ?? []).length} col(s):`);

    const fmt = request.params.format ?? 'both';
    if (fmt === 'json' || fmt === 'both') {
      response.attachJson({ sheetName, range: request.params.range, headers: cellRange.headers, rows: cellRange.rows });
    }
    if (fmt === 'table' || fmt === 'both') {
      response.attachMarkdown(rangeToMarkdownTable(cellRange));
    }
  },
});

export const spreadsheetSetCell = defineTool({
  name: 'spreadsheet_set_cell',
  description: 'Set the value or formula of a single cell in a spreadsheet. Changes are saved immediately.',
  annotations: {
    category: DocumentCategory.SPREADSHEET,
    readOnlyHint: false,
    title: 'Set Cell',
  },
  schema: {
    ...docIdSchema,
    ...sheetNameSchema,
    cell: z.string().describe('Cell address (e.g., "A1", "B3")'),
    value: z.union([z.string(), z.number(), z.boolean(), z.null()]).describe('Cell value'),
    formula: z.string().optional().describe('Excel formula (e.g., "=SUM(A1:A10)"). If provided, takes precedence over value.'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'calc') {
      throw new Error('spreadsheet_set_cell only works with spreadsheet documents.');
    }
    // Parse fresh for write, then invalidate cache
    const wb = await parseXlsx(session.parsedPath);
    const sheetName = request.params.sheetName ?? wb.sheetNames[0]!;
    setCellValue(wb, sheetName, request.params.cell, request.params.value, request.params.formula);
    await saveXlsx(wb, session.parsedPath);
    session.invalidateCache();
    response.appendText(`Cell ${request.params.cell} in "${sheetName}" updated to: ${request.params.formula ?? request.params.value}`);
  },
});

export const spreadsheetSetRange = defineTool({
  name: 'spreadsheet_set_range',
  description: 'Set multiple cell values in a spreadsheet. Provide a 2D array of values matching the range.',
  annotations: {
    category: DocumentCategory.SPREADSHEET,
    readOnlyHint: false,
    title: 'Set Cell Range',
  },
  schema: {
    ...docIdSchema,
    ...sheetNameSchema,
    startCell: z.string().describe('Top-left cell address (e.g., "A1")'),
    values: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).describe('2D array of values (rows × columns)'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'calc') {
      throw new Error('spreadsheet_set_range only works with spreadsheet documents.');
    }
    const wb = await parseXlsx(session.parsedPath);
    const sheetName = request.params.sheetName ?? wb.sheetNames[0]!;
    setRangeValues(wb, sheetName, request.params.startCell, request.params.values);
    await saveXlsx(wb, session.parsedPath);
    session.invalidateCache();

    response.appendText(`Updated ${request.params.values.length} row(s) × ${request.params.values[0]?.length ?? 0} col(s) starting at ${request.params.startCell}.`);
  },
});

export const spreadsheetAddSheet = defineTool({
  name: 'spreadsheet_add_sheet',
  description: 'Add a new sheet to a spreadsheet workbook.',
  annotations: {
    category: DocumentCategory.SPREADSHEET,
    readOnlyHint: false,
    title: 'Add Sheet',
  },
  schema: {
    ...docIdSchema,
    sheetName: z.string().describe('Name for the new sheet'),
    headers: z.array(z.string()).optional().describe('Optional column headers for the first row'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'calc') {
      throw new Error('spreadsheet_add_sheet only works with spreadsheet documents.');
    }
    const wb = await parseXlsx(session.parsedPath);
    if (wb.sheetNames.includes(request.params.sheetName)) {
      throw new Error(`Sheet "${request.params.sheetName}" already exists.`);
    }
    addNewSheet(wb, request.params.sheetName, request.params.headers);
    await saveXlsx(wb, session.parsedPath);
    session.invalidateCache();
    response.appendText(`Sheet "${request.params.sheetName}" added successfully.`);
  },
});

export const spreadsheetGetFormulas = defineTool({
  name: 'spreadsheet_get_formulas',
  description: 'Get all formulas in a spreadsheet range (returns formula expressions, not computed values).',
  annotations: {
    category: DocumentCategory.SPREADSHEET,
    readOnlyHint: true,
    title: 'Get Formulas',
  },
  schema: {
    ...docIdSchema,
    ...sheetNameSchema,
    range: z.string().optional().describe('Cell range (e.g., "A1:Z100"). Defaults to entire sheet.'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'calc') {
      throw new Error('spreadsheet_get_formulas only works with spreadsheet documents.');
    }
    const wb = await session.getOrParseXlsx(() => parseXlsx(session.parsedPath));
    const sheetName = request.params.sheetName ?? wb.sheetNames[0]!;
    const formulas = getFormulas(wb, sheetName, request.params.range);
    if (formulas.length === 0) {
      response.appendText(`No formulas found in "${sheetName}"${request.params.range ? ` (range: ${request.params.range})` : ''}.`);
    } else {
      response.appendText(`${formulas.length} formula(s) in "${sheetName}":`);
      response.attachJson(formulas);
    }
  },
});
