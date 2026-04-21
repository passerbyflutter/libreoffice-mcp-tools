import { z } from 'zod';
import { copyFile, rm } from 'node:fs/promises';
import { defineTool, docIdSchema } from './ToolDefinition.js';
import { DocumentCategory } from './categories.js';
import type { SupportedOutputFormat } from '../LibreOfficeAdapter.js';

const SUPPORTED_OUTPUT_FORMATS = ['pdf', 'docx', 'xlsx', 'pptx', 'html', 'txt', 'csv', 'odt', 'ods', 'odp'] as const;

export const documentSave = defineTool({
  name: 'document_save',
  description: 'Save the current state of a document to its file path (or a new path).',
  annotations: {
    category: DocumentCategory.FILES,
    readOnlyHint: false,
    title: 'Save Document',
  },
  schema: {
    ...docIdSchema,
    filePath: z.string().optional().describe('New file path to save to (optional — saves to original path if omitted)'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    const targetPath = request.params.filePath ?? session.originalPath;

    if (session.parsedPath !== session.originalPath && !request.params.filePath) {
      // Bridged file (e.g. ODT was opened as DOCX for editing).
      // Must reverse-convert back to original format — copying raw DOCX bytes would corrupt it.
      if (!context.adapter.isAvailable()) {
        throw new Error(
          `Saving to ${session.originalExt} requires LibreOffice (used for format bridging). ` +
          `Install LibreOffice, or provide an explicit filePath ending in .docx/.xlsx/.pptx to save as the working format.`,
        );
      }
      const originalFormat = session.originalExt.replace('.', '') as SupportedOutputFormat;
      const guard = await context.mutex.acquire();
      try {
        const { outputPath, tempDir } = await context.adapter.convertFile(
          session.parsedPath,
          originalFormat,
        );
        try {
          await copyFile(outputPath, targetPath);
        } finally {
          await rm(tempDir, { recursive: true, force: true }).catch(() => {});
        }
      } finally {
        guard.dispose();
      }
    } else if (request.params.filePath) {
      // Explicit target path — just copy the working file (parsedPath) there
      await copyFile(session.parsedPath, targetPath);
    }
    // If parsedPath === originalPath, write operations already saved in-place; nothing to do

    response.appendText(`Document saved to: ${targetPath}`);
  },
});

export const documentExport = defineTool({
  name: 'document_export',
  description: `Export a document to a different format using LibreOffice.
Supports exporting to: PDF, HTML, TXT, DOCX, XLSX, PPTX, ODT, ODS, ODP, CSV.
LibreOffice must be installed.`,
  annotations: {
    category: DocumentCategory.FILES,
    readOnlyHint: false,
    title: 'Export Document',
  },
  schema: {
    ...docIdSchema,
    format: z.enum(SUPPORTED_OUTPUT_FORMATS).describe('Target format'),
    outputPath: z.string().optional().describe('Output file path (default: same directory as source, new extension)'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (!context.adapter.isAvailable()) {
      throw new Error(`document_export requires LibreOffice. Install it or set SOFFICE_PATH.`);
    }

    const guard = await context.mutex.acquire();
    try {
      const { outputPath, tempDir } = await context.adapter.convertFile(
        session.parsedPath,
        request.params.format as SupportedOutputFormat,
      );

      const finalPath = request.params.outputPath ?? outputPath;
      try {
        if (finalPath !== outputPath) {
          await copyFile(outputPath, finalPath);
        }
        response.appendText(`Exported to ${request.params.format.toUpperCase()}: ${finalPath}`);
      } finally {
        await rm(tempDir, { recursive: true, force: true }).catch(() => {});
      }
    } finally {
      guard.dispose();
    }
  },
});

export const documentConvert = defineTool({
  name: 'document_convert',
  description: `Convert a document file to a different format using LibreOffice CLI.
Creates a new file; original is not modified. Useful for: DOC→DOCX, DOCX→PDF, XLSX→CSV, etc.
LibreOffice must be installed.`,
  annotations: {
    category: DocumentCategory.FILES,
    readOnlyHint: false,
    title: 'Convert Document',
  },
  schema: {
    filePath: z.string().describe('Source file path'),
    format: z.enum(SUPPORTED_OUTPUT_FORMATS).describe('Target format'),
    outputPath: z.string().optional().describe('Output file path (optional)'),
  },
  handler: async (request, response, context) => {
    if (!context.adapter.isAvailable()) {
      throw new Error(`document_convert requires LibreOffice. Install it or set SOFFICE_PATH.`);
    }
    const guard = await context.mutex.acquire();
    try {
      const { outputPath, tempDir } = await context.adapter.convertFile(
        request.params.filePath,
        request.params.format as SupportedOutputFormat,
      );
      const finalPath = request.params.outputPath ?? outputPath;
      try {
        if (finalPath !== outputPath) {
          await copyFile(outputPath, finalPath);
        }
        response.appendText(`Converted ${request.params.filePath} → ${request.params.format.toUpperCase()}: ${finalPath}`);
      } finally {
        await rm(tempDir, { recursive: true, force: true }).catch(() => {});
      }
    } finally {
      guard.dispose();
    }
  },
});
