# LibreOffice MCP Tools — Copilot Instructions

## Project Overview
A TypeScript MCP (Model Context Protocol) server that lets AI/LLM agents read, write, and edit Office documents (DOCX, DOC, XLSX, XLS, PPTX, PPT, ODF, PDF, CSV, TXT) via LibreOffice. Designed to minimize token usage through structured, range-based, outline-first document access.

**Reference architecture**: [chrome-devtools-mcp](https://github.com/ChromeDevTools/chrome-devtools-mcp)

## Architecture Principles

### Token Efficiency (Core Design Goal)
1. **Outline-first** — `document_get_outline` returns headings/sheet names/slide titles BEFORE content
2. **Range-based access** — `document_read_range` with paragraph/row/slide offsets, never dump whole doc
3. **Pagination** — ALL list tools support `limit` + `offset` params
4. **Structured JSON** — spreadsheet data as `{headers: string[], rows: any[][]}` not raw CSV
5. **Markdown output** — text documents rendered as compact markdown (headings → #, bold → **, etc.)
6. **Metadata-first** — word count, page count, sheet names BEFORE reading content

### Technology Stack
- **Language**: TypeScript, ESM modules (`"type": "module"` in package.json)
- **MCP SDK**: `@modelcontextprotocol/sdk` 
- **Schema validation**: `zod`
- **Build**: `tsc` (TypeScript compiler)
- **Native document parsers** (fast, no LibreOffice needed for reads):
  - DOCX/DOTX: `mammoth` for text extraction, `docx` for writing
  - XLSX/XLS/ODS/CSV: `xlsx` (SheetJS)
  - PPTX: `jszip` + custom XML parser
- **LibreOffice CLI** (`soffice --headless`) for:
  - Converting legacy binary formats (DOC→DOCX, XLS→XLSX, PPT→PPTX) before parsing
  - Exporting to PDF, HTML, etc.
  - Writing to ODF formats
  - Complex write operations on any format

### File Format Support
| Format | Extensions | Read | Write | Method |
|---|---|---|---|---|
| Word 2007+ | `.docx`, `.dotx` | ✅ | ✅ | Native (mammoth/docx) |
| Word 97-2003 | `.doc`, `.dot` | ✅ | ✅ | LibreOffice → DOCX bridge |
| Excel 2007+ | `.xlsx`, `.xlsm` | ✅ | ✅ | Native (SheetJS) |
| Excel 97-2003 | `.xls` | ✅ | ✅ | LibreOffice → XLSX bridge |
| PowerPoint 2007+ | `.pptx` | ✅ | ✅ | Native (JSZip XML) |
| PowerPoint 97-2003 | `.ppt` | ✅ | ✅ | LibreOffice → PPTX bridge |
| OpenDocument Text | `.odt` | ✅ | ✅ | LibreOffice bridge |
| OpenDocument Spreadsheet | `.ods` | ✅ | ✅ | LibreOffice bridge |
| OpenDocument Presentation | `.odp` | ✅ | ✅ | LibreOffice bridge |
| Rich Text Format | `.rtf` | ✅ | ✅ | LibreOffice bridge |
| CSV | `.csv` | ✅ | ✅ | Native (SheetJS/built-in) |
| PDF | `.pdf` | ✅ (text) | ❌ | LibreOffice CLI |
| Plain text | `.txt` | ✅ | ✅ | Native |

**Binary format bridge**: DOC/XLS/PPT are auto-converted via `soffice --headless --convert-to docx/xlsx/pptx` on `document_open`, then handled by the same native parsers as modern formats. The conversion happens to a temp file; the original is never modified.

## Project Structure

```
src/
├── index.ts                    # createMcpServer() factory, tool registration
├── LibreOfficeAdapter.ts       # soffice subprocess, convert, detect path
├── DocumentContext.ts          # Open document registry / session state
├── DocumentSession.ts          # Per-document handle, metadata cache, format bridge
├── McpResponse.ts              # Response builder (appendText, attachJson, etc.)
├── Mutex.ts                    # Serial execution guard (soffice doesn't support concurrency)
├── logger.ts                   # DEBUG=lo-mcp:* debug logging
├── types.ts                    # Shared TS types
├── parsers/
│   ├── DocxParser.ts           # DOCX → {metadata, outline, paragraphs, text}
│   ├── XlsxParser.ts           # XLSX/CSV → {sheets, ranges, cells}
│   └── PptxParser.ts           # PPTX → {slides: {title, body, notes}[]}
├── formatters/
│   ├── MarkdownFormatter.ts    # Document paragraphs → Markdown string
│   ├── JsonFormatter.ts        # Document → compact JSON
│   └── TableFormatter.ts       # Spreadsheet range → Markdown table
└── tools/
    ├── ToolDefinition.ts       # defineTool(), Context/Response/Request interfaces
    ├── categories.ts           # DocumentCategory enum
    ├── tools.ts                # createTools() aggregator
    ├── documents.ts            # document_open, _close, _list, _create
    ├── reader.ts               # document_get_metadata, _get_outline, _read_text, _read_range, _search
    ├── writer.ts               # document_insert_text, _replace_text, _insert_paragraph, _apply_style
    ├── spreadsheet.ts          # spreadsheet_list_sheets, _get_range, _set_cell, _set_range, _add_sheet, _get_formulas
    ├── presentation.ts         # presentation_list_slides, _get_slide, _add_slide, _update_slide, _get_notes
    └── converter.ts            # document_save, _export, _convert
bin/
└── libreoffice-mcp.ts          # CLI entry: arg parsing → createMcpServer() → stdio transport
```

## Coding Conventions

### Tool Definition Pattern (ALWAYS use this)
```typescript
// src/tools/example.ts
import { defineTool } from './ToolDefinition.js';
import { DocumentCategory } from './categories.js';
import { z } from 'zod';

export const myTool = defineTool({
  name: 'tool_name',           // snake_case
  description: 'What it does. Token-efficient: explain what subset of data it returns.',
  annotations: {
    category: DocumentCategory.READING,
    readOnlyHint: true,        // true if tool doesn't modify documents
  },
  schema: {
    docId: z.string().describe('Document handle returned by document_open'),
    limit: z.number().int().optional().default(50).describe('Max items to return'),
    offset: z.number().int().optional().default(0).describe('Pagination offset'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    // ... do work ...
    response.appendText('Result text');
    // OR:
    response.attachJson({ key: 'value' });
  },
});
```

### Error Handling
- Throw `Error` with descriptive messages; the framework catches and returns `isError: true`
- Check `context.getDocument(docId)` throws if docId not found
- LibreOffice subprocess failures should include the stderr output in the error message

### Response Conventions
- Use `response.appendText()` for human-readable summaries
- Use `response.attachJson()` for structured data (spreadsheet ranges, metadata)
- Use `response.attachMarkdown()` for document text content
- ALWAYS include a text summary even when attaching JSON (so LLM gets context)
- Keep text output ≤ 2000 chars by default; respect pagination

### LibreOffice Subprocess
- All LibreOffice calls go through `LibreOfficeAdapter` — NEVER spawn soffice directly in tools
- Use the `Mutex` before calling any LibreOffice subprocess (soffice cannot run concurrently)
- Set reasonable timeouts: 30s for conversions, 60s for complex operations
- Check for LibreOffice path at startup; warn if not found

### TypeScript
- All files use ESM (`import`/`export`), `.js` extensions in import paths (even for .ts source)
- Strict TypeScript: `strict: true` in tsconfig
- Never use `any` — use proper types or `unknown`
- Zod schemas for all tool inputs

## Tool Naming Conventions
- `document_*` — works on any document type
- `spreadsheet_*` — spreadsheet-specific (XLSX, ODS, CSV)
- `presentation_*` — presentation-specific (PPTX, ODP)
- All names are snake_case

## Complete Tool Catalog

### Document Management
| Tool | Category | R/W | Description |
|---|---|---|---|
| `document_open` | FILES | — | Register file → return docId. Auto-bridges DOC/XLS/PPT via LibreOffice |
| `document_close` | FILES | — | Release docId and temp files |
| `document_list` | FILES | Read | List open documents with format, path, size |
| `document_create` | FILES | Write | Create empty document (writer/calc/impress) |
| `document_save` | FILES | Write | Save to current or new path |
| `document_export` | FILES | Write | Export via LibreOffice (PDF/HTML/CSV/TXT) |
| `document_convert` | FILES | Write | Convert format (DOCX→ODT, etc.) |

### Reading
| Tool | Category | R/W | Description |
|---|---|---|---|
| `document_get_metadata` | READING | Read | Title, author, page/word count, format, dates |
| `document_get_outline` | READING | Read | Headings (Writer), sheet names (Calc), slide titles (Impress) |
| `document_read_text` | READING | Read | Full or range-based text as markdown |
| `document_read_range` | READING | Read | Specific paragraph range N–M or page N |
| `document_search` | READING | Read | Find text with surrounding context |

### Writing (Writer/Text)
| Tool | Category | R/W | Description |
|---|---|---|---|
| `document_insert_text` | WRITING | Write | Insert at position/heading/bookmark |
| `document_replace_text` | WRITING | Write | Find & replace (first or all occurrences) |
| `document_insert_paragraph` | WRITING | Write | Insert paragraph with optional style |
| `document_apply_style` | WRITING | Write | Apply heading/character style to range |

### Spreadsheet
| Tool | Category | R/W | Description |
|---|---|---|---|
| `spreadsheet_list_sheets` | SPREADSHEET | Read | Sheet names with row/col counts |
| `spreadsheet_get_range` | SPREADSHEET | Read | Cells as JSON {headers, rows} or markdown table |
| `spreadsheet_set_cell` | SPREADSHEET | Write | Set cell value/formula |
| `spreadsheet_set_range` | SPREADSHEET | Write | Set 2D range of values |
| `spreadsheet_add_sheet` | SPREADSHEET | Write | Add new sheet |
| `spreadsheet_get_formulas` | SPREADSHEET | Read | Formulas in range (not computed values) |

### Presentation
| Tool | Category | R/W | Description |
|---|---|---|---|
| `presentation_list_slides` | PRESENTATION | Read | Slide titles with index |
| `presentation_get_slide` | PRESENTATION | Read | Full slide: title, body, notes |
| `presentation_add_slide` | PRESENTATION | Write | Add slide with title/content |
| `presentation_update_slide` | PRESENTATION | Write | Update title or body of existing slide |
| `presentation_get_notes` | PRESENTATION | Read | Speaker notes (all or specific slide) |

## Environment
- Node.js ≥ 20 required
- LibreOffice must be installed for: legacy format support, PDF export, format conversion
- LibreOffice path auto-detected: Windows `C:\Program Files\LibreOffice\program\soffice.exe`, macOS `/Applications/LibreOffice.app/Contents/MacOS/soffice`, Linux `/usr/bin/soffice`
- Override with `--libreoffice-path` CLI arg or `SOFFICE_PATH` env var

## MCP Configuration (.mcp.json)
```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "node",
      "args": ["./build/bin/libreoffice-mcp.js"],
      "env": {}
    }
  }
}
```
