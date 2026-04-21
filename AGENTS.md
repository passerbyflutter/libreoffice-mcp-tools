# AGENTS.md — LibreOffice MCP Tools

## What This Is
TypeScript MCP server for AI agents to read/write/edit Office documents via LibreOffice.
Mirrors chrome-devtools-mcp architecture. Token-efficient by design.

## Quick Context
- Entry: `src/index.ts` → `createMcpServer()`
- Tools defined in `src/tools/*.ts` using `defineTool()` from `src/tools/ToolDefinition.ts`
- All tools take `docId` (from `document_open`) + optional pagination params
- LibreOffice subprocess managed by `src/LibreOfficeAdapter.ts` behind `src/Mutex.ts`
- DOC/XLS/PPT legacy formats auto-converted on `document_open` via LibreOffice CLI

## Key Files to Read First
1. `src/tools/ToolDefinition.ts` — how to define a tool
2. `src/DocumentContext.ts` — session/document management
3. `src/tools/tools.ts` — all tool registration
4. `.github/copilot-instructions.md` — full context

## Coding Rules
- TypeScript ESM, `.js` imports even for .ts files
- Zod schemas for all tool params
- Never spawn `soffice` directly — use `LibreOfficeAdapter`
- Always acquire `Mutex` before LibreOffice calls
- Response: `appendText()` for summaries, `attachJson()` for data, `attachMarkdown()` for doc content

## Supported Formats
DOCX, DOC (via LO bridge), XLSX, XLS (via LO bridge), PPTX, PPT (via LO bridge),
ODT, ODS, ODP, RTF (all via LO bridge), CSV, TXT (native), PDF (read-only via LO)

## Build & Run
```bash
npm install
npm run build
node build/bin/libreoffice-mcp.js
```

## Tests
```bash
npm test
```
