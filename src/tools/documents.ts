import { z } from 'zod';
import { existsSync } from 'node:fs';
import { writeFile, mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';
import { defineTool, docIdSchema } from './ToolDefinition.js';
import { DocumentCategory } from './categories.js';

export const documentOpen = defineTool({
  name: 'document_open',
  description: `Open a document file and return a docId handle for use with other tools.
Supports: .docx, .doc, .xlsx, .xls, .pptx, .ppt, .odt, .ods, .odp, .rtf, .csv, .txt, .pdf
Legacy binary formats (.doc, .xls, .ppt, .odt, .ods, .odp, .rtf) are automatically converted via LibreOffice before parsing.
Returns docId to use with all other document tools.`,
  annotations: {
    category: DocumentCategory.FILES,
    readOnlyHint: false,
    title: 'Open Document',
  },
  schema: {
    filePath: z.string().describe('Absolute or relative path to the document file'),
  },
  handler: async (request, response, context) => {
    const { filePath } = request.params;
    if (!existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }
    const session = await context.openDocument(filePath);
    const needsBridge = session.originalPath !== session.parsedPath;
    response.appendText(`Document opened successfully.`);
    response.attachJson({
      docId: session.docId,
      filePath: session.originalPath,
      format: session.originalExt.replace('.', '').toUpperCase(),
      documentType: session.getDocumentType(),
      bridged: needsBridge,
      ...(needsBridge ? { parsedAs: session.parsedExt.replace('.', '').toUpperCase() } : {}),
    });
    response.appendText(`Use docId "${session.docId}" with other tools. Start with document_get_metadata or document_get_outline.`);
  },
});

export const documentClose = defineTool({
  name: 'document_close',
  description: 'Close an open document and release its resources (including any temp files from format bridging).',
  annotations: {
    category: DocumentCategory.FILES,
    readOnlyHint: false,
    title: 'Close Document',
  },
  schema: {
    ...docIdSchema,
  },
  handler: async (request, response, context) => {
    const { docId } = request.params;
    await context.closeDocument(docId);
    response.appendText(`Document ${docId} closed.`);
  },
});

export const documentList = defineTool({
  name: 'document_list',
  description: 'List all currently open documents with their docId, file path, format, and size.',
  annotations: {
    category: DocumentCategory.FILES,
    readOnlyHint: true,
    title: 'List Open Documents',
  },
  schema: {},
  handler: async (request, response, context) => {
    const docs = context.listDocuments();
    if (docs.length === 0) {
      response.appendText('No documents are currently open. Use document_open to open a file.');
      return;
    }
    response.appendText(`${docs.length} document(s) open:`);
    response.attachJson(docs);
  },
});

export const documentCreate = defineTool({
  name: 'document_create',
  description: `Create a new empty document and open it. Returns a docId.
Supports creating: writer (text document → .docx), calc (spreadsheet → .xlsx), impress (presentation → .pptx).`,
  annotations: {
    category: DocumentCategory.FILES,
    readOnlyHint: false,
    title: 'Create Document',
  },
  schema: {
    documentType: z.enum(['writer', 'calc', 'impress']).describe('Type of document to create'),
    filePath: z.string().describe('Path where the new document will be saved'),
  },
  handler: async (request, response, context) => {
    const { documentType, filePath } = request.params;

    const extMap: Record<string, string> = { writer: '.docx', calc: '.xlsx', impress: '.pptx' };
    const ext = extMap[documentType]!;
    const targetPath = filePath.endsWith(ext) ? filePath : `${filePath}${ext}`;

    await mkdir(dirname(targetPath), { recursive: true });

    if (documentType === 'calc') {
      const { createEmptyXlsx } = await import('../parsers/XlsxParser.js');
      await createEmptyXlsx(targetPath);
    } else if (documentType === 'writer') {
      const { Document, Packer } = await import('docx');
      const doc = new Document({ sections: [{ children: [] }] });
      const buffer = await Packer.toBuffer(doc);
      await writeFile(targetPath, buffer);
    } else {
      // impress — build a minimal valid PPTX using JSZip (no LibreOffice required)
      const { createEmptyPptx } = await import('../parsers/PptxOoxmlEditor.js');
      await createEmptyPptx(targetPath);
    }

    const session = await context.openDocument(targetPath);
    response.appendText(`New ${documentType} document created and opened.`);
    response.attachJson({
      docId: session.docId,
      filePath: targetPath,
      documentType,
    });
  },
});
