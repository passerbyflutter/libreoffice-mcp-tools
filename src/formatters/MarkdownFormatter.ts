import type { Paragraph } from '../types.js';

/**
 * Format document paragraphs as Markdown.
 */
export function paragraphsToMarkdown(paragraphs: Paragraph[]): string {
  return paragraphs.map(p => {
    if (p.level !== undefined && p.level > 0) {
      return `${'#'.repeat(p.level)} ${p.text}`;
    }
    return p.text;
  }).join('\n\n');
}

/**
 * Truncate markdown to a max character count, appending a note if truncated.
 */
export function truncateMarkdown(markdown: string, maxChars: number): string {
  if (markdown.length <= maxChars) return markdown;
  return markdown.slice(0, maxChars) + '\n\n... *(content truncated — use document_read_range for more)*';
}
