/**
 * OOXML-based DOCX editor using JSZip.
 *
 * Reads and writes word/document.xml directly, preserving all document formatting
 * (fonts, colors, tables, images, headers, footers, styles, hyperlinks, etc.)
 * that is not explicitly targeted by the edit operation.
 *
 * Trade-off vs. mammoth+docx rebuild: This approach is surgical — it only
 * modifies what it needs to, leaving everything else intact. The limitation
 * is that text search can miss matches split across multiple <w:t> runs within
 * a paragraph (e.g. due to spell-check mid-word splits).
 */

import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

// ─── Public API ────────────────────────────────────────────────────────────────

/** Open a DOCX for in-place XML editing. */
export async function openDocxForEdit(filePath: string): Promise<DocxOoxmlEditor> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  return new DocxOoxmlEditor(zip, filePath);
}

export class DocxOoxmlEditor {
  private readonly zip: JSZip;
  private readonly filePath: string;

  constructor(zip: JSZip, filePath: string) {
    this.zip = zip;
    this.filePath = filePath;
  }

  // ─── Core helpers ──────────────────────────────────────────────────────────

  private async readDocXml(): Promise<string> {
    const file = this.zip.file('word/document.xml');
    if (!file) throw new Error('Invalid DOCX: word/document.xml not found.');
    return file.async('string');
  }

  private writeDocXml(xml: string): void {
    this.zip.file('word/document.xml', xml);
  }

  async save(targetPath?: string): Promise<void> {
    const buffer = await this.zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
    await writeFile(targetPath ?? this.filePath, buffer);
  }

  // ─── Text replacement ───────────────────────────────────────────────────────

  /**
   * Find and replace text within <w:t> nodes.
   * Returns the number of substitutions made.
   *
   * Known limitation: text split across multiple <w:t> runs within a paragraph
   * (e.g., caused by spell-check boundaries) will not be matched as a single unit.
   */
  async replaceText(
    find: string,
    replace: string,
    opts: { replaceAll?: boolean; caseInsensitive?: boolean } = {},
  ): Promise<number> {
    const { replaceAll = true, caseInsensitive = false } = opts;
    let xml = await this.readDocXml();
    let count = 0;

    const escapedFind = escapeRegex(find);
    const flags = (replaceAll ? 'g' : '') + (caseInsensitive ? 'i' : '');
    const regex = new RegExp(escapedFind, flags || undefined);

    // Only target text content inside <w:t> elements
    xml = xml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (_, attrs: string, content: string) => {
      if (!replaceAll && count > 0) return `<w:t${attrs}>${content}</w:t>`;
      const newContent = content.replace(regex, (_m: string) => {
        count++;
        return escapeXml(replace);
      });
      return `<w:t${attrs}>${newContent}</w:t>`;
    });

    this.writeDocXml(xml);
    return count;
  }

  // ─── Paragraph insertion ────────────────────────────────────────────────────

  /**
   * Insert a new paragraph at the specified position.
   * All existing content and formatting is preserved.
   */
  async insertParagraph(text: string, opts: {
    position?: 'start' | 'end';
    afterIndex?: number;   // 0-based paragraph index to insert after
    style?: string;        // e.g. "Heading 1", "Normal"
  } = {}): Promise<void> {
    const { position = 'end', afterIndex, style } = opts;
    let xml = await this.readDocXml();

    const newPara = buildParagraphXml(text, style);

    if (afterIndex !== undefined) {
      xml = insertAfterNthParagraph(xml, newPara, afterIndex);
    } else if (position === 'start') {
      xml = insertAtBodyStart(xml, newPara);
    } else {
      xml = insertAtBodyEnd(xml, newPara);
    }

    this.writeDocXml(xml);
  }

  // ─── Style application ──────────────────────────────────────────────────────

  /**
   * Apply a paragraph style to the Nth paragraph (0-based index) in the document body.
   * Only the paragraph's <w:pPr><w:pStyle> is modified; all other content is preserved.
   */
  async applyParagraphStyle(paragraphIndex: number, style: string): Promise<void> {
    let xml = await this.readDocXml();
    const styleId = styleNameToId(style);
    const result = modifyParagraphStyle(xml, paragraphIndex, styleId);
    if (!result.found) {
      throw new Error(`Paragraph index ${paragraphIndex} is out of range.`);
    }
    xml = result.xml;
    this.writeDocXml(xml);
  }

  /**
   * Count top-level paragraphs in the document body.
   * Useful for validating paragraph indices before operations.
   */
  async getParagraphCount(): Promise<number> {
    const xml = await this.readDocXml();
    return countBodyParagraphs(xml);
  }
}

// ─── XML helpers ───────────────────────────────────────────────────────────────

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Map human-friendly style names to OOXML style IDs. */
function styleNameToId(name: string): string {
  const map: Record<string, string> = {
    'Heading 1': 'Heading1',
    'Heading 2': 'Heading2',
    'Heading 3': 'Heading3',
    'Heading 4': 'Heading4',
    'Heading 5': 'Heading5',
    'Heading 6': 'Heading6',
    'Normal': 'Normal',
    'Default Paragraph Style': 'Normal',
  };
  return map[name] ?? name.replace(/\s+/g, '');
}

/** Build a minimal <w:p> XML string with optional style. */
function buildParagraphXml(text: string, style?: string): string {
  const pPr = style
    ? `<w:pPr><w:pStyle w:val="${styleNameToId(style)}"/></w:pPr>`
    : '';
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

// ─── Paragraph location utilities ─────────────────────────────────────────────

/**
 * Find all <w:p> elements within the <w:body> that contain visible text, in document order.
 * Empty structural paragraphs (no <w:t> content) are excluded, matching mammoth's
 * paragraph numbering which also skips empty paragraphs.
 *
 * This allows paragraph indices from document_read_text (mammoth) to align with the
 * indices used for document_apply_style and document_insert_paragraph.
 *
 * Returns spans as absolute positions within the full `xml` string.
 */
function findBodyParagraphSpans(xml: string): Array<{ start: number; end: number }> {
  const bodyOpen = xml.indexOf('<w:body');
  const bodyClose = xml.lastIndexOf('</w:body>');
  if (bodyOpen === -1 || bodyClose === -1) return [];

  const spans: Array<{ start: number; end: number }> = [];
  let i = bodyOpen;

  while (i < bodyClose) {
    const pStart = xml.indexOf('<w:p', i);
    if (pStart === -1 || pStart >= bodyClose) break;

    // Validate it's a <w:p> element (not <w:pStyle, <w:pPr, etc.)
    const ch = xml[pStart + 4]; // char after '<w:p'
    if (ch !== '>' && ch !== ' ' && ch !== '/') {
      i = pStart + 4;
      continue;
    }

    // Check if it's self-closing (<w:p/>)
    const tagClose = xml.indexOf('>', pStart);
    if (tagClose === -1) break;
    if (xml[tagClose - 1] === '/') {
      // Self-closing: no text content possible, skip (no need to include in spans)
      i = tagClose + 1;
      continue;
    }

    const closePos = findMatchingClose(xml, pStart, 'w:p');
    if (closePos === -1) break;

    const endPos = closePos + '</w:p>'.length;
    const block = xml.slice(pStart, endPos);

    // Only include paragraphs that have visible text content (matching mammoth's filter)
    if (/<w:t[^>]*>[^<\s][^<]*<\/w:t>/.test(block) || /<w:t[^>]*>\s*\S[^<]*<\/w:t>/.test(block)) {
      spans.push({ start: pStart, end: endPos });
    }

    // Always skip past this paragraph (don't re-enter its children)
    i = endPos;
  }

  return spans;
}

/**
 * Count top-level paragraphs in the document body.
 */
function countBodyParagraphs(xml: string): number {
  return findBodyParagraphSpans(xml).length;
}

/**
 * Find the position of the closing </w:p> for the <w:p> that starts at `openPos`
 * within `xml`. Returns the index of the '<' in '</w:p>'.
 * Accounts for nested <w:p> elements.
 */
function findMatchingClose(xml: string, openPos: number, tagName: string): number {
  const openTag = `<${tagName}`;
  const closeTag = `</${tagName}>`;
  let depth = 1;
  let i = openPos + openTag.length;

  while (i < xml.length && depth > 0) {
    const nextOpen = xml.indexOf(openTag, i);
    const nextClose = xml.indexOf(closeTag, i);

    if (nextClose === -1) return -1;

    if (nextOpen !== -1 && nextOpen < nextClose) {
      // Verify it's actually <w:p or <w:p> (not <w:pStyle etc.)
      const ch = xml[nextOpen + openTag.length];
      if (ch === '>' || ch === ' ') {
        // And not self-closing
        const te = xml.indexOf('>', nextOpen);
        if (te !== -1 && xml[te - 1] !== '/') {
          depth++;
        }
      }
      i = nextOpen + openTag.length;
    } else {
      depth--;
      if (depth === 0) return nextClose;
      i = nextClose + closeTag.length;
    }
  }

  return -1;
}

// ─── Insertion helpers ─────────────────────────────────────────────────────────

function insertAtBodyStart(xml: string, newPara: string): string {
  // Insert after the <w:body> or <w:body ...> opening tag
  const bodyTagEnd = xml.indexOf('>', xml.indexOf('<w:body'));
  if (bodyTagEnd === -1) return xml;
  return xml.slice(0, bodyTagEnd + 1) + newPara + xml.slice(bodyTagEnd + 1);
}

function insertAtBodyEnd(xml: string, newPara: string): string {
  // Insert before <w:sectPr> if present (it must be last in w:body)
  // otherwise before </w:body>
  const sectPrIdx = xml.lastIndexOf('<w:sectPr');
  const bodyCloseIdx = xml.lastIndexOf('</w:body>');
  if (bodyCloseIdx === -1) return xml + newPara;

  const insertAt = (sectPrIdx !== -1 && sectPrIdx < bodyCloseIdx) ? sectPrIdx : bodyCloseIdx;
  return xml.slice(0, insertAt) + newPara + xml.slice(insertAt);
}

function insertAfterNthParagraph(xml: string, newPara: string, targetIndex: number): string {
  const spans = findBodyParagraphSpans(xml);
  const span = spans[targetIndex];
  if (!span) {
    // Index out of range — append at end
    return insertAtBodyEnd(xml, newPara);
  }
  return xml.slice(0, span.end) + newPara + xml.slice(span.end);
}

// ─── Style modification ────────────────────────────────────────────────────────

function modifyParagraphStyle(
  xml: string,
  targetIndex: number,
  styleId: string,
): { xml: string; found: boolean } {
  const spans = findBodyParagraphSpans(xml);
  const span = spans[targetIndex];
  if (!span) return { xml, found: false };

  const pXml = xml.slice(span.start, span.end);
  const modifiedP = setParagraphStyle(pXml, styleId);
  return {
    xml: xml.slice(0, span.start) + modifiedP + xml.slice(span.end),
    found: true,
  };
}

/**
 * Set or replace <w:pStyle> within a <w:p> XML snippet.
 * Adds <w:pPr> if not already present.
 */
function setParagraphStyle(paragraphXml: string, styleId: string): string {
  const styleTag = `<w:pStyle w:val="${styleId}"/>`;

  if (/<w:pPr(?:\s[^>]*)?>/.test(paragraphXml)) {
    if (paragraphXml.includes('<w:pStyle')) {
      // Replace existing pStyle
      return paragraphXml.replace(/<w:pStyle[^/]*\/>/, styleTag);
    } else {
      // Add pStyle inside existing pPr
      return paragraphXml.replace(/(<w:pPr(?:\s[^>]*)?>)/, `$1${styleTag}`);
    }
  } else {
    // Add <w:pPr> with pStyle after <w:p> opening tag
    const pTagClose = paragraphXml.indexOf('>');
    if (pTagClose === -1) return paragraphXml;
    return (
      paragraphXml.slice(0, pTagClose + 1) +
      `<w:pPr>${styleTag}</w:pPr>` +
      paragraphXml.slice(pTagClose + 1)
    );
  }
}
