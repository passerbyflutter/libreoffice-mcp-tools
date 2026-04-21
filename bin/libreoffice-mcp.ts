#!/usr/bin/env node
/**
 * CLI entry point for LibreOffice MCP server.
 * Usage: node build/bin/libreoffice-mcp.js [--libreoffice-path <path>]
 */

import { startServer } from '../src/index.js';

const args = process.argv.slice(2);
let sofficePath: string | undefined;

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--libreoffice-path' && args[i + 1]) {
    sofficePath = args[i + 1];
    i++;
  } else if (args[i]?.startsWith('--libreoffice-path=')) {
    sofficePath = args[i]!.split('=')[1];
  }
}

process.stderr.write(
  'LibreOffice MCP Tools — reading, writing, and editing Office documents via LibreOffice\n' +
  'Uses LibreOffice for: legacy format support (.doc, .xls, .ppt), PDF export, format conversion\n' +
  'Set SOFFICE_PATH env var or use --libreoffice-path to specify LibreOffice location\n',
);

startServer({ sofficePath }).catch(err => {
  process.stderr.write(`Fatal error: ${err.message}\n`);
  process.exit(1);
});
