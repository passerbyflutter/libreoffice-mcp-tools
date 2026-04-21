/**
 * DOCX parser using mammoth for text extraction and native XML for structure.
 * Handles .docx, .dotx files. Legacy .doc files are bridged via LibreOffice first.
 */
import { readFile } from 'node:fs/promises';
import mammoth from 'mammoth';
import type {
  DocumentMetadata,
  OutlineItem,
  Paragraph,
  SearchResult,
} from '../types.js';

export interface DocxDocument {
  metadata: DocumentMetadata;
  paragraphs: Paragraph[];
  outline: OutlineItem[];
}

export async function parseDocx(filePath: string): Promise<DocxDocument> {
  const buffer = await readFile(filePath);

  // Extract structured content with mammoth
  const { value: html } = await mammoth.convertToHtml({ buffer });

  // Parse paragraphs from HTML output
  const paragraphs = extractParagraphsFromHtml(html);
  const outline = extractOutline(paragraphs);
  const metadata = await extractDocxMetadata(filePath, paragraphs);

  return { metadata, paragraphs, outline };
}

function stripTags(html: string): string {
  return html.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").trim();
}

function extractParagraphsFromHtml(html: string): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  let index = 0;

  // Match block-level elements: headings and paragraphs
  const blockPattern = /<(h[1-6]|p)(\s[^>]*)?>[\s\S]*?<\/\1>/gi;
  const matches = html.matchAll(blockPattern);

  for (const match of matches) {
    const tag = match[1]!.toLowerCase();
    const text = stripTags(match[0]);
    if (!text) continue;

    let level: number | undefined;
    let style: string | undefined;

    if (tag.startsWith('h')) {
      level = parseInt(tag[1]!, 10);
      style = `Heading ${level}`;
    }

    paragraphs.push({ index: index++, text, style, level });
  }

  return paragraphs;
}

function extractOutline(paragraphs: Paragraph[]): OutlineItem[] {
  return paragraphs
    .filter(p => p.level !== undefined)
    .map((p) => ({
      level: p.level!,
      title: p.text,
      index: p.index,
    }));
}

async function extractDocxMetadata(filePath: string, paragraphs: Paragraph[]): Promise<DocumentMetadata> {
  const wordCount = paragraphs.reduce((acc, p) => acc + p.text.split(/\s+/).filter(Boolean).length, 0);

  // Try to get file size
  let fileSize: number | undefined;
  try {
    const { statSync } = await import('node:fs');
    fileSize = statSync(filePath).size;
  } catch {}

  return {
    format: 'DOCX',
    filePath,
    documentType: 'writer',
    wordCount,
    fileSize,
  };
}

export function paragraphsToMarkdown(paragraphs: Paragraph[], startIndex = 0, endIndex?: number): string {
  const slice = paragraphs.slice(startIndex, endIndex);
  return slice.map(p => {
    if (p.level !== undefined) {
      return `${'#'.repeat(p.level)} ${p.text}`;
    }
    return p.text;
  }).join('\n\n');
}

export function searchDocx(paragraphs: Paragraph[], query: string, limit = 10): SearchResult[] {
  const results: SearchResult[] = [];
  const lowerQuery = query.toLowerCase();

  for (const p of paragraphs) {
    const lowerText = p.text.toLowerCase();
    let pos = 0;
    while ((pos = lowerText.indexOf(lowerQuery, pos)) !== -1 && results.length < limit) {
      const contextStart = Math.max(0, pos - 80);
      const contextEnd = Math.min(p.text.length, pos + query.length + 80);
      results.push({
        paragraphIndex: p.index,
        text: p.text,
        context: p.text.slice(contextStart, contextEnd),
        matchStart: pos,
        matchEnd: pos + query.length,
      });
      pos += query.length;
    }
    if (results.length >= limit) break;
  }

  return results;
}
