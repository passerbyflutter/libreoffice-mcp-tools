import { z } from 'zod';
import { defineTool, docIdSchema } from './ToolDefinition.js';
import { DocumentCategory } from './categories.js';
import { openDocxForEdit } from '../parsers/DocxOoxmlEditor.js';

// Heading style map reused across write tools
const HEADING_STYLES: Record<string, string> = {
  'Heading 1': 'Heading 1',
  'Heading 2': 'Heading 2',
  'Heading 3': 'Heading 3',
  'Heading 4': 'Heading 4',
  'Heading 5': 'Heading 5',
  'Heading 6': 'Heading 6',
  'Normal': 'Normal',
};

export const documentInsertText = defineTool({
  name: 'document_insert_text',
  description: `Insert text into a Writer document at a specified position.
Position options: 'start' (beginning), 'end' (end of document), or after a specific heading text.`,
  annotations: {
    category: DocumentCategory.WRITING,
    readOnlyHint: false,
    title: 'Insert Text',
  },
  schema: {
    ...docIdSchema,
    text: z.string().describe('Text content to insert'),
    position: z.enum(['start', 'end', 'after_heading']).optional().default('end').describe('Where to insert: start, end, or after_heading'),
    headingText: z.string().optional().describe('Required when position is after_heading: the heading text to insert after'),
    style: z.string().optional().describe('Paragraph style (e.g., "Heading 1", "Normal"). Default: Normal'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'writer') {
      throw new Error('document_insert_text only works with Writer (text) documents.');
    }

    const { text, position = 'end', style, headingText } = request.params;
    const editor = await openDocxForEdit(session.parsedPath);

    if (position === 'after_heading' && headingText) {
      // Find the paragraph index of the heading
      const { parseDocx } = await import('../parsers/DocxParser.js');
      const docData = await parseDocx(session.parsedPath);
      const headingIdx = docData.paragraphs.findIndex(
        p => p.text.toLowerCase() === headingText.toLowerCase(),
      );
      if (headingIdx === -1) throw new Error(`Heading "${headingText}" not found.`);
      await editor.insertParagraph(text, { afterIndex: headingIdx, style });
    } else {
      await editor.insertParagraph(text, {
        position: position === 'start' ? 'start' : 'end',
        style,
      });
    }

    await editor.save();
    session.invalidateCache();
    response.appendText(`Text inserted at ${position}.`);
  },
});

export const documentReplaceText = defineTool({
  name: 'document_replace_text',
  description: 'Find and replace text in a Writer document. Optionally replace all occurrences.',
  annotations: {
    category: DocumentCategory.WRITING,
    readOnlyHint: false,
    title: 'Replace Text',
  },
  schema: {
    ...docIdSchema,
    find: z.string().describe('Text to find (case-sensitive by default)'),
    replace: z.string().describe('Replacement text'),
    replaceAll: z.boolean().optional().default(true).describe('Replace all occurrences (default: true)'),
    caseInsensitive: z.boolean().optional().default(false).describe('Case-insensitive matching'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'writer') {
      throw new Error('document_replace_text only works with Writer (text) documents.');
    }

    const { find, replace, replaceAll = true, caseInsensitive = false } = request.params;
    const editor = await openDocxForEdit(session.parsedPath);
    const count = await editor.replaceText(find, replace, { replaceAll, caseInsensitive });
    await editor.save();
    session.invalidateCache();

    response.appendText(`Replaced ${count} occurrence(s) of "${find}" with "${replace}".`);
  },
});

export const documentInsertParagraph = defineTool({
  name: 'document_insert_paragraph',
  description: 'Insert a new paragraph at a specific index in a Writer document.',
  annotations: {
    category: DocumentCategory.WRITING,
    readOnlyHint: false,
    title: 'Insert Paragraph',
  },
  schema: {
    ...docIdSchema,
    text: z.string().describe('Paragraph text'),
    index: z.number().int().describe('Index at which to insert the paragraph (0-based)'),
    style: z.string().optional().describe('Style name (e.g., "Heading 1", "Normal")'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'writer') {
      throw new Error('document_insert_paragraph only works with Writer documents.');
    }

    const { text, index, style } = request.params;
    const editor = await openDocxForEdit(session.parsedPath);
    const total = await editor.getParagraphCount();
    const clampedIndex = Math.min(Math.max(index - 1, -1), total - 1);

    if (clampedIndex < 0) {
      // Insert at the very beginning
      await editor.insertParagraph(text, { position: 'start', style });
    } else {
      await editor.insertParagraph(text, { afterIndex: clampedIndex, style });
    }

    await editor.save();
    session.invalidateCache();
    response.appendText(`Paragraph inserted at index ${index}.`);
  },
});

export const documentApplyStyle = defineTool({
  name: 'document_apply_style',
  description: 'Apply a heading or paragraph style to a specific paragraph by index.',
  annotations: {
    category: DocumentCategory.WRITING,
    readOnlyHint: false,
    title: 'Apply Style',
  },
  schema: {
    ...docIdSchema,
    paragraphIndex: z.number().int().describe('Index of the paragraph to style (0-based)'),
    style: z.string().describe(`Style to apply (e.g., ${Object.keys(HEADING_STYLES).join(', ')})`),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'writer') {
      throw new Error('document_apply_style only works with Writer documents.');
    }

    const { paragraphIndex, style } = request.params;
    const editor = await openDocxForEdit(session.parsedPath);
    await editor.applyParagraphStyle(paragraphIndex, style);
    await editor.save();
    session.invalidateCache();

    response.appendText(`Style "${style}" applied to paragraph ${paragraphIndex}.`);
  },
});

