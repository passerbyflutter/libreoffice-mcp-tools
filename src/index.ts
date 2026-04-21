import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { DocumentContext } from './DocumentContext.js';
import { McpResponse } from './McpResponse.js';
import { logger } from './logger.js';
import { createTools } from './tools/tools.js';
import type { ToolDefinition } from './tools/ToolDefinition.js';

export interface McpServerOptions {
  sofficePath?: string;
}

export async function createMcpServer(options: McpServerOptions = {}): Promise<McpServer> {
  const server = new McpServer({
    name: 'libreoffice-mcp',
    version: '0.1.0',
  });

  const context = new DocumentContext(options.sofficePath);

  const tools = createTools();

  for (const tool of tools) {
    registerTool(server, tool, context);
  }

  // Clean up on SIGINT/SIGTERM — 'exit' event is synchronous so async cleanup runs on signals instead
  process.on('SIGINT', async () => { await context.closeAll(); process.exit(0); });
  process.on('SIGTERM', async () => { await context.closeAll(); process.exit(0); });

  return server;
}

function registerTool(server: McpServer, tool: ToolDefinition, context: DocumentContext): void {
  type Params = z.objectOutputType<z.ZodRawShape, z.ZodTypeAny>;

  server.tool(
    tool.name,
    tool.description,
    tool.schema,
    async (params: Params) => {
      logger(`Tool call: ${tool.name}`, JSON.stringify(params));
      const response = new McpResponse();
      try {
        await tool.handler({ params: params as z.objectOutputType<z.ZodRawShape, z.ZodTypeAny> }, response, context);
        const text = response.build();
        return {
          content: [{ type: 'text' as const, text: text || '(no output)' }],
        };
      } catch (err) {
        const errorText = err instanceof Error ? err.message : String(err);
        logger(`Tool error: ${tool.name}`, errorText);
        return {
          content: [{ type: 'text' as const, text: `Error: ${errorText}` }],
          isError: true,
        };
      }
    },
  );
}

export async function startServer(options: McpServerOptions = {}): Promise<void> {
  const server = await createMcpServer(options);
  const transport = new StdioServerTransport();
  await server.connect(transport);
  logger('LibreOffice MCP server started');
}
