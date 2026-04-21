/**
 * Shared types for libreoffice-mcp-tools
 */

export type DocumentType = 'writer' | 'calc' | 'impress' | 'unknown';

export interface DocumentMetadata {
  title?: string;
  author?: string;
  subject?: string;
  description?: string;
  keywords?: string[];
  created?: string;
  modified?: string;
  pageCount?: number;
  wordCount?: number;
  characterCount?: number;
  format: string;
  filePath: string;
  fileSize?: number;
  documentType: DocumentType;
}

export interface OutlineItem {
  level: number;
  title: string;
  index: number;
}

export interface Paragraph {
  index: number;
  text: string;
  style?: string;
  level?: number;
}

export interface SearchResult {
  paragraphIndex: number;
  text: string;
  context: string;
  matchStart: number;
  matchEnd: number;
}

export interface SheetInfo {
  name: string;
  index: number;
  rowCount: number;
  colCount: number;
}

export interface CellRange {
  headers?: string[];
  rows: (string | number | boolean | null)[][];
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  sheetName: string;
}

export interface SlideContent {
  index: number;
  title?: string;
  body?: string;
  notes?: string;
  layoutName?: string;
}

export interface PaginationParams {
  limit?: number;
  offset?: number;
}

export const DEFAULT_LIMIT = 50;
export const DEFAULT_TIMEOUT_MS = 30_000;
export const SOFFICE_TIMEOUT_MS = 60_000;
