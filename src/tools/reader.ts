import { z } from 'zod';
import { defineTool, docIdSchema, paginationSchema } from './ToolDefinition.js';
import { DocumentCategory } from './categories.js';
import { parseDocx, paragraphsToMarkdown, searchDocx } from '../parsers/DocxParser.js';
import { parseXlsx, getSheetInfo, searchCells } from '../parsers/XlsxParser.js';
import { parsePptx } from '../parsers/PptxParser.js';
import { truncateMarkdown } from '../formatters/MarkdownFormatter.js';
import { formatOutline, formatSheetList, formatSlideList } from '../formatters/JsonFormatter.js';

const MAX_TEXT_CHARS = 4000;

async function getDocxData(session: import('../DocumentSession.js').DocumentSession) {
  return session.getOrParseDocx(() => parseDocx(session.parsedPath));
}

export const documentGetMetadata = defineTool({
  name: 'document_get_metadata',
  description: `Get metadata for an open document: title, author, page count, word count, format, file size, dates.
Token-efficient: use this before reading document content to understand its size and structure.`,
  annotations: {
    category: DocumentCategory.READING,
    readOnlyHint: true,
    title: 'Get Document Metadata',
  },
  schema: {
    ...docIdSchema,
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    const docType = session.getDocumentType();

    let metadata;
    if (docType === 'writer') {
      const data = await getDocxData(session);
      metadata = data.metadata;
    } else if (docType === 'calc') {
      const wb = await session.getOrParseXlsx(() => parseXlsx(session.parsedPath));
      const sheets = getSheetInfo(wb);
      metadata = {
        format: session.originalExt.replace('.', '').toUpperCase(),
        filePath: session.originalPath,
        documentType: 'calc' as const,
        sheetCount: sheets.length,
        totalRows: sheets.reduce((a, s) => a + s.rowCount, 0),
      };
    } else if (docType === 'impress') {
      const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
      metadata = {
        format: session.originalExt.replace('.', '').toUpperCase(),
        filePath: session.originalPath,
        documentType: 'impress' as const,
        slideCount: pptx.slideCount,
      };
    } else {
      metadata = {
        format: session.originalExt.replace('.', '').toUpperCase(),
        filePath: session.originalPath,
        documentType: 'unknown' as const,
      };
    }

    response.appendText(`Metadata for ${session.originalPath}:`);
    response.attachJson(metadata);
  },
});

export const documentGetOutline = defineTool({
  name: 'document_get_outline',
  description: `Get the structural outline of a document:
- Writer documents: headings hierarchy (H1, H2, H3...)
- Calc spreadsheets: list of sheet names with row/column counts
- Impress presentations: slide titles with index numbers
Token-efficient: always call this before document_read_text to understand structure.`,
  annotations: {
    category: DocumentCategory.READING,
    readOnlyHint: true,
    title: 'Get Document Outline',
  },
  schema: {
    ...docIdSchema,
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    const docType = session.getDocumentType();

    if (docType === 'writer') {
      const data = await getDocxData(session);
      if (data.outline.length === 0) {
        response.appendText('Document has no headings. Use document_read_text to read full content.');
      } else {
        response.appendText(`Document outline (${data.outline.length} headings):`);
        response.attachJson(formatOutline(data.outline));
        response.appendText(`Use document_read_range with offset/limit to read specific sections.`);
      }
    } else if (docType === 'calc') {
      const wb = await session.getOrParseXlsx(() => parseXlsx(session.parsedPath));
      const sheets = getSheetInfo(wb);
      response.appendText(`Spreadsheet sheets (${sheets.length}):`);
      response.attachJson(formatSheetList(sheets));
      response.appendText(`Use spreadsheet_get_range to read cell data.`);
    } else if (docType === 'impress') {
      const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
      response.appendText(`Presentation slides (${pptx.slideCount}):`);
      response.attachJson(formatSlideList(pptx.slides));
      response.appendText(`Use presentation_get_slide to read individual slide content.`);
    } else {
      response.appendText('Outline not available for this document type. Use document_read_text.');
    }
  },
});

export const documentReadText = defineTool({
  name: 'document_read_text',
  description: `Read document content as Markdown text. For large documents, use limit/offset for pagination.
Default returns up to ${MAX_TEXT_CHARS} characters. Use document_get_outline first to understand structure.`,
  annotations: {
    category: DocumentCategory.READING,
    readOnlyHint: true,
    title: 'Read Document Text',
  },
  schema: {
    ...docIdSchema,
    ...paginationSchema,
    maxChars: z.number().int().optional().default(MAX_TEXT_CHARS).describe(`Maximum characters to return. Default: ${MAX_TEXT_CHARS}`),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    const { offset = 0, limit = 50, maxChars = MAX_TEXT_CHARS } = request.params;
    const docType = session.getDocumentType();

    if (docType === 'writer') {
      const data = await getDocxData(session);
      const slicedParagraphs = data.paragraphs.slice(offset, offset + limit);
      const markdown = paragraphsToMarkdown(slicedParagraphs);
      const truncated = truncateMarkdown(markdown, maxChars);
      response.appendText(`Content (paragraphs ${offset}–${offset + slicedParagraphs.length - 1} of ${data.paragraphs.length}):`);
      response.attachMarkdown(truncated);
      if (offset + limit < data.paragraphs.length) {
        response.appendText(`More content available. Use offset=${offset + limit} for next page.`);
      }
    } else if (docType === 'calc') {
      response.appendText('For spreadsheets, use spreadsheet_get_range instead of document_read_text.');
    } else if (docType === 'impress') {
      const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
      const slides = pptx.slides.slice(offset, offset + limit);
      const text = slides.map(s =>
        `## Slide ${s.index + 1}: ${s.title ?? '(no title)'}\n\n${s.body ?? '(no content)'}`
      ).join('\n\n---\n\n');
      response.attachMarkdown(truncateMarkdown(text, maxChars));
    } else {
      const { readFile } = await import('node:fs/promises');
      const text = await readFile(session.parsedPath, 'utf-8');
      response.attachMarkdown(truncateMarkdown(text, maxChars));
    }
  },
});

export const documentReadRange = defineTool({
  name: 'document_read_range',
  description: `Read a specific range of a document by paragraph index (for Writer) or slide index (for Impress).
More token-efficient than document_read_text when you only need a specific section.`,
  annotations: {
    category: DocumentCategory.READING,
    readOnlyHint: true,
    title: 'Read Document Range',
  },
  schema: {
    ...docIdSchema,
    startIndex: z.number().int().describe('Start paragraph/slide index (0-based)'),
    endIndex: z.number().int().optional().describe('End paragraph/slide index (exclusive). Defaults to startIndex + 10'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    const { startIndex, endIndex } = request.params;
    const end = endIndex ?? startIndex + 10;
    const docType = session.getDocumentType();

    if (docType === 'writer') {
      const data = await getDocxData(session);
      const slice = data.paragraphs.slice(startIndex, end);
      const markdown = paragraphsToMarkdown(slice);
      response.appendText(`Paragraphs ${startIndex}–${Math.min(end, data.paragraphs.length) - 1}:`);
      response.attachMarkdown(markdown);
    } else if (docType === 'impress') {
      const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
      const slides = pptx.slides.slice(startIndex, end);
      const text = slides.map(s =>
        `## Slide ${s.index + 1}: ${s.title ?? '(no title)'}\n\n${s.body ?? '(no content)'}`
      ).join('\n\n---\n\n');
      response.attachMarkdown(text);
    } else {
      response.appendText('document_read_range is for Writer and Impress. For Calc, use spreadsheet_get_range.');
    }
  },
});

export const documentSearch = defineTool({
  name: 'document_search',
  description: 'Search for text within a document. Returns matching paragraphs/cells/slides with surrounding context.',
  annotations: {
    category: DocumentCategory.READING,
    readOnlyHint: true,
    title: 'Search Document',
  },
  schema: {
    ...docIdSchema,
    query: z.string().describe('Text to search for (case-insensitive)'),
    limit: z.number().int().optional().default(10).describe('Max results to return. Default: 10'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    const { query, limit = 10 } = request.params;
    const docType = session.getDocumentType();

    if (docType === 'writer') {
      const data = await getDocxData(session);
      const results = searchDocx(data.paragraphs, query, limit);
      if (results.length === 0) {
        response.appendText(`No results found for "${query}".`);
      } else {
        response.appendText(`Found ${results.length} result(s) for "${query}":`);
        response.attachJson(results.map(r => ({
          paragraph: r.paragraphIndex,
          context: r.context,
        })));
      }
    } else if (docType === 'calc') {
      const wb = await session.getOrParseXlsx(() => parseXlsx(session.parsedPath));
      const results = searchCells(wb, query, limit);
      response.appendText(`Found ${results.length} cell(s) matching "${query}":`);
      response.attachJson(results);
    } else if (docType === 'impress') {
      const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
      const lowerQuery = query.toLowerCase();
      const results = pptx.slides.filter(
        s => (s.title ?? '').toLowerCase().includes(lowerQuery) ||
             (s.body ?? '').toLowerCase().includes(lowerQuery)
      ).slice(0, limit);
      response.appendText(`Found ${results.length} slide(s) matching "${query}":`);
      response.attachJson(results.map(s => ({ index: s.index, title: s.title, bodyPreview: s.body?.slice(0, 100) })));
    }
  },
});
