import { existsSync, statSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { DocumentSession } from './DocumentSession.js';
import { LibreOfficeAdapter } from './LibreOfficeAdapter.js';
import { Mutex } from './Mutex.js';
import { logger } from './logger.js';

export class DocumentContext {
  private readonly sessions = new Map<string, DocumentSession>();
  readonly adapter: LibreOfficeAdapter;
  readonly mutex: Mutex;

  constructor(sofficePath?: string) {
    this.adapter = new LibreOfficeAdapter(sofficePath);
    this.mutex = new Mutex();
  }

  async openDocument(filePath: string): Promise<DocumentSession> {
    if (!existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }

    // Check if already open
    for (const session of this.sessions.values()) {
      if (session.originalPath === filePath && !session.isClosed()) {
        logger(`Reusing existing session ${session.docId} for ${filePath}`);
        return session;
      }
    }

    const docId = randomUUID();
    const session = await DocumentSession.create(docId, filePath, this.adapter, this.mutex);
    this.sessions.set(docId, session);
    logger(`Opened document ${docId}: ${filePath}`);
    return session;
  }

  getDocument(docId: string): DocumentSession {
    const session = this.sessions.get(docId);
    if (!session || session.isClosed()) {
      throw new Error(`Document not found: ${docId}. Use document_open first.`);
    }
    return session;
  }

  async closeDocument(docId: string): Promise<void> {
    const session = this.sessions.get(docId);
    if (session) {
      await session.close();
      this.sessions.delete(docId);
      logger(`Closed document ${docId}`);
    }
  }

  listDocuments(): Array<{ docId: string; filePath: string; format: string; size: number }> {
    return Array.from(this.sessions.values())
      .filter(s => !s.isClosed())
      .map(s => ({
        docId: s.docId,
        filePath: s.originalPath,
        format: s.originalExt.replace('.', '').toUpperCase(),
        size: existsSync(s.originalPath) ? statSync(s.originalPath).size : 0,
      }));
  }

  async closeAll(): Promise<void> {
    for (const session of this.sessions.values()) {
      await session.close().catch(() => {});
    }
    this.sessions.clear();
  }
}
