import type { ToolDefinition } from './ToolDefinition.js';
import * as documentTools from './documents.js';
import * as readerTools from './reader.js';
import * as writerTools from './writer.js';
import * as spreadsheetTools from './spreadsheet.js';
import * as presentationTools from './presentation.js';
import * as converterTools from './converter.js';

export function createTools(): ToolDefinition[] {
  const allTools = [
    ...Object.values(documentTools),
    ...Object.values(readerTools),
    ...Object.values(writerTools),
    ...Object.values(spreadsheetTools),
    ...Object.values(presentationTools),
    ...Object.values(converterTools),
  ] as ToolDefinition[];

  return allTools.sort((a, b) => a.name.localeCompare(b.name));
}
