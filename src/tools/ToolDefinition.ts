import { z } from 'zod';
import type { DocumentContext } from '../DocumentContext.js';
import type { McpResponse } from '../McpResponse.js';
import type { DocumentCategory } from './categories.js';

export interface ToolAnnotations {
  category: DocumentCategory;
  readOnlyHint: boolean;
  title?: string;
}

export interface ToolRequest<Schema extends z.ZodRawShape> {
  params: z.objectOutputType<Schema, z.ZodTypeAny>;
}

export interface ToolDefinition<Schema extends z.ZodRawShape = z.ZodRawShape> {
  name: string;
  description: string;
  annotations: ToolAnnotations;
  schema: Schema;
  handler: (
    request: ToolRequest<Schema>,
    response: McpResponse,
    context: DocumentContext,
  ) => Promise<void>;
}

export function defineTool<Schema extends z.ZodRawShape>(
  definition: ToolDefinition<Schema>,
): ToolDefinition<Schema> {
  return definition;
}

/** Common pagination schema fields to include in tools that return lists */
export const paginationSchema = {
  limit: z.number().int().optional().default(50).describe('Maximum number of items to return. Default: 50'),
  offset: z.number().int().optional().default(0).describe('Pagination offset. Default: 0'),
};

/** Common docId schema field */
export const docIdSchema = {
  docId: z.string().describe('Document handle returned by document_open'),
};

/** Common sheet name schema */
export const sheetNameSchema = {
  sheetName: z.string().optional().describe('Sheet name. Defaults to the first sheet if not specified'),
};
