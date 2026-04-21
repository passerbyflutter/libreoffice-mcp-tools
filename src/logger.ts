const isDebug = process.env.DEBUG?.includes('lo-mcp') || process.env.DEBUG === '*';

export const logger = (...args: unknown[]): void => {
  if (isDebug) {
    const timestamp = new Date().toISOString();
    process.stderr.write(`[lo-mcp ${timestamp}] ${args.map(String).join(' ')}\n`);
  }
};
