import type { DocumentMetadata, OutlineItem, SheetInfo, SlideContent } from '../types.js';

export function formatMetadata(metadata: DocumentMetadata): Record<string, unknown> {
  const result: Record<string, unknown> = {
    format: metadata.format,
    documentType: metadata.documentType,
    filePath: metadata.filePath,
  };
  if (metadata.title) result['title'] = metadata.title;
  if (metadata.author) result['author'] = metadata.author;
  if (metadata.subject) result['subject'] = metadata.subject;
  if (metadata.pageCount !== undefined) result['pageCount'] = metadata.pageCount;
  if (metadata.wordCount !== undefined) result['wordCount'] = metadata.wordCount;
  if (metadata.fileSize !== undefined) result['fileSize'] = metadata.fileSize;
  if (metadata.created) result['created'] = metadata.created;
  if (metadata.modified) result['modified'] = metadata.modified;
  return result;
}

export function formatOutline(items: OutlineItem[]): Array<{ level: number; title: string; index: number }> {
  return items.map(item => ({ level: item.level, title: item.title, index: item.index }));
}

export function formatSheetList(sheets: SheetInfo[]): Array<{ name: string; index: number; rows: number; cols: number }> {
  return sheets.map(s => ({ name: s.name, index: s.index, rows: s.rowCount, cols: s.colCount }));
}

export function formatSlideList(slides: SlideContent[]): Array<{ index: number; title: string | undefined }> {
  return slides.map(s => ({ index: s.index, title: s.title }));
}
