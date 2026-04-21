/**
 * Response builder for MCP tool handlers.
 * Provides typed methods for different content types.
 */
export class McpResponse {
  private textLines: string[] = [];
  private jsonAttachments: unknown[] = [];
  private markdownBlocks: string[] = [];

  appendText(line: string): void {
    this.textLines.push(line);
  }

  attachJson(data: unknown): void {
    this.jsonAttachments.push(data);
  }

  attachMarkdown(markdown: string): void {
    this.markdownBlocks.push(markdown);
  }

  build(): string {
    const parts: string[] = [];

    if (this.textLines.length > 0) {
      parts.push(this.textLines.join('\n'));
    }

    for (const json of this.jsonAttachments) {
      parts.push('```json\n' + JSON.stringify(json, null, 2) + '\n```');
    }

    for (const md of this.markdownBlocks) {
      parts.push(md);
    }

    return parts.join('\n\n');
  }

  isEmpty(): boolean {
    return this.textLines.length === 0 && this.jsonAttachments.length === 0 && this.markdownBlocks.length === 0;
  }
}
