# LibreOffice MCP Tools

[![npm version](https://img.shields.io/npm/v/@passerbyflutter/libreoffice-mcp-tools.svg)](https://npmjs.org/package/@passerbyflutter/libreoffice-mcp-tools)

> [!WARNING]
> 本專案由 **GitHub Copilot** 自動生成，未經完整人工審閱。
> 程式碼可能包含錯誤、安全疑慮或非預期行為，**請自行評估使用風險**，勿直接用於生產環境。
> This project was **written by GitHub Copilot**. Use at your own risk.

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that gives AI agents (Claude, Copilot, Gemini, Cursor, etc.) the ability to **read, write, and edit Office documents** via LibreOffice — with a token-efficient design that minimizes LLM context usage.

Inspired by the architecture of [chrome-devtools-mcp](https://github.com/ChromeDevTools/chrome-devtools-mcp).

## ✨ Features

- **22 MCP tools** covering reading, writing, spreadsheets, and presentations
- **Token-efficient design**: outline-first navigation, range-based access, pagination
- **Broad format support**: DOCX, DOC, XLSX, XLS, PPTX, PPT, ODT, ODS, ODP, RTF, CSV, TXT, PDF
- **Legacy format bridge**: `.doc`, `.xls`, `.ppt` auto-converted via LibreOffice before parsing
- **No LibreOffice required for basic reads**: native parsers handle DOCX, XLSX, PPTX directly
- **LibreOffice required for**: legacy formats, PDF export, format conversion

## 📋 Supported Formats

| Format | Extensions | Read | Write | Method |
|---|---|---|---|---|
| Word 2007+ | `.docx`, `.dotx` | ✅ | ✅ | Native (mammoth read / JSZip OOXML write) |
| Word 97-2003 | `.doc`, `.dot` | ✅ | ✅ | LibreOffice bridge |
| Excel 2007+ | `.xlsx`, `.xlsm` | ✅ | ✅ | Native (ExcelJS) |
| Excel 97-2003 | `.xls` | ✅ | ✅ | LibreOffice bridge |
| PowerPoint 2007+ | `.pptx` | ✅ | ✅ | Native (JSZip OOXML) |
| PowerPoint 97-2003 | `.ppt` | ✅ | ✅ | LibreOffice bridge |
| OpenDocument Text | `.odt` | ✅ | ✅ | LibreOffice bridge |
| OpenDocument Spreadsheet | `.ods` | ✅ | ✅ | LibreOffice bridge |
| OpenDocument Presentation | `.odp` | ✅ | ✅ | LibreOffice bridge |
| Rich Text Format | `.rtf` | ✅ | ✅ | LibreOffice bridge |
| CSV | `.csv` | ✅ | ✅ | Native |
| PDF | `.pdf` | ✅ (text) | ❌ | LibreOffice CLI |
| Plain text | `.txt` | ✅ | ✅ | Native |

## 🚀 Quick Start

### Prerequisites

- **Node.js 20+**
- **LibreOffice** (optional for basic DOCX/XLSX/PPTX reads; required for .doc/.xls/.ppt and format conversion)
  - Windows: [Download LibreOffice](https://www.libreoffice.org/download/download/)
  - macOS: `brew install --cask libreoffice`
  - Linux: `sudo apt install libreoffice` or `sudo dnf install libreoffice`

### Installation

**Using npx (recommended — no install needed):**

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "npx",
      "args": ["-y", "@passerbyflutter/libreoffice-mcp-tools"]
    }
  }
}
```

**Global install:**

```bash
npm install -g @passerbyflutter/libreoffice-mcp-tools
```

**From source:**

```bash
git clone https://github.com/passerbyflutter/libreoffice-mcp-tools
cd libreoffice-mcp-tools
npm install
npm run build
```

### Configure your MCP client

Add to your MCP client configuration (e.g., Claude Desktop `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "npx",
      "args": ["-y", "@passerbyflutter/libreoffice-mcp-tools"],
      "env": {
        "SOFFICE_PATH": "/path/to/soffice"
      }
    }
  }
}
```

Or use `.mcp.json` at your project root:

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "npx",
      "args": ["-y", "@passerbyflutter/libreoffice-mcp-tools"]
    }
  }
}
```

### CLI Options

```
node build/bin/libreoffice-mcp.js [options]

  --libreoffice-path <path>   Path to soffice executable
                              (default: auto-detected or SOFFICE_PATH env)
```

## 🛠 Tool Reference

### Document Management

| Tool | Description |
|---|---|
| `document_open` | Open a file → returns `docId` handle. Auto-bridges legacy formats. |
| `document_close` | Release document handle and temp files |
| `document_list` | List all open documents |
| `document_create` | Create new empty document (writer/calc/impress) |
| `document_save` | Save to current or new path |
| `document_export` | Export via LibreOffice (PDF, HTML, CSV, etc.) |
| `document_convert` | Convert file format (DOC→DOCX, XLSX→CSV, etc.) |

### Reading (Token-Efficient)

| Tool | Description |
|---|---|
| `document_get_metadata` | Title, author, word/page count, dates |
| `document_get_outline` | Headings (Writer) / sheet names (Calc) / slide titles (Impress) |
| `document_read_text` | Paginated document text as Markdown |
| `document_read_range` | Specific paragraph or slide range |
| `document_search` | Find text with surrounding context |

### Writing (Writer)

| Tool | Description |
|---|---|
| `document_insert_text` | Insert at start/end/after heading |
| `document_replace_text` | Find & replace (first or all occurrences) |
| `document_insert_paragraph` | Insert paragraph at specific index |
| `document_apply_style` | Apply heading/paragraph style |

### Spreadsheet (Calc)

| Tool | Description |
|---|---|
| `spreadsheet_list_sheets` | Sheet names with row/col counts |
| `spreadsheet_get_range` | Cell range as JSON + markdown table |
| `spreadsheet_set_cell` | Set cell value or formula |
| `spreadsheet_set_range` | Set 2D range of values |
| `spreadsheet_add_sheet` | Add new sheet |
| `spreadsheet_get_formulas` | Get formula expressions in range |

### Presentation (Impress)

| Tool | Description |
|---|---|
| `presentation_list_slides` | Slide titles with index |
| `presentation_get_slide` | Full slide content (title, body, notes) |
| `presentation_get_notes` | Speaker notes |
| `presentation_add_slide` | Add new slide (requires LibreOffice) |
| `presentation_update_slide` | Update slide content |

## 💡 Token-Saving Workflow

For maximum token efficiency, follow this pattern:

```
1. document_open(filePath) → get docId
2. document_get_metadata(docId) → understand size/type
3. document_get_outline(docId) → see structure
4. document_read_range(docId, startIndex=N, endIndex=M) → read specific section
```

Instead of dumping the entire document, you navigate to exactly what you need.

**Spreadsheet workflow:**
```
1. document_open(path) → docId
2. spreadsheet_list_sheets(docId) → see all sheets
3. spreadsheet_get_range(docId, sheetName="Sales", range="A1:D20") → targeted data
```

## 🏗 Architecture

```
src/
├── index.ts                # createMcpServer() — MCP server factory
├── LibreOfficeAdapter.ts   # soffice subprocess manager
├── DocumentContext.ts      # Open document registry
├── DocumentSession.ts      # Per-document state + format bridge
├── McpResponse.ts          # Response builder (text/JSON/markdown)
├── Mutex.ts                # Serializes LibreOffice subprocess calls
├── parsers/
│   ├── DocxParser.ts           # DOCX read → {paragraphs, outline, metadata} (mammoth)
│   ├── DocxOoxmlEditor.ts      # DOCX write → direct JSZip OOXML manipulation (format-preserving)
│   ├── XlsxParser.ts           # XLSX read/write via ExcelJS
│   ├── PptxParser.ts           # PPTX read → {slides[]} (JSZip XML)
│   └── PptxOoxmlEditor.ts      # PPTX write → add/update slides, create PPTX (JSZip OOXML)
├── formatters/
│   ├── MarkdownFormatter.ts
│   ├── JsonFormatter.ts
│   └── TableFormatter.ts   # Spreadsheet → Markdown table
└── tools/
    ├── documents.ts         # open/close/list/create
    ├── reader.ts            # metadata/outline/read/search
    ├── writer.ts            # insert/replace/style
    ├── spreadsheet.ts       # get/set cells/ranges/sheets
    ├── presentation.ts      # slides/notes
    └── converter.ts         # save/export/convert
```

## 🧪 Testing

```bash
# Create sample fixtures
node tests/create-fixtures.mjs

# Run smoke tests
npm test
```

## 📝 Environment Variables

| Variable | Description |
|---|---|
| `SOFFICE_PATH` | Path to LibreOffice `soffice` executable |
| `DEBUG` | Set to `lo-mcp:*` for verbose logging |

## 📄 License

MIT
