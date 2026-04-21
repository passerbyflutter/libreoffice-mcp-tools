import { rm } from 'node:fs/promises';
import { extname } from 'node:path';
import type { DocumentMetadata, DocumentType } from './types.js';
import { LibreOfficeAdapter } from './LibreOfficeAdapter.js';
import type { Mutex } from './Mutex.js';
import { logger } from './logger.js';

// Forward-declare cache types to avoid circular imports at runtime
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyCache = any;

export class DocumentSession {
  readonly docId: string;
  readonly originalPath: string;
  readonly originalExt: string;
  
  /** Path used for parsing — may differ from originalPath if LibreOffice bridging was used */
  parsedPath: string;
  parsedExt: string;

  private tempDir: string | null = null;
  private metadataCache: DocumentMetadata | null = null;
  private closed = false;

  /** Parsed document caches — invalidated on any write operation */
  private _docxCache: AnyCache | null = null;
  private _xlsxCache: AnyCache | null = null;
  private _pptxCache: AnyCache | null = null;

  constructor(
    docId: string,
    originalPath: string,
    parsedPath: string,
    tempDir: string | null,
  ) {
    this.docId = docId;
    this.originalPath = originalPath;
    this.originalExt = extname(originalPath).toLowerCase();
    this.parsedPath = parsedPath;
    this.parsedExt = extname(parsedPath).toLowerCase();
    this.tempDir = tempDir;
  }

  static async create(
    docId: string,
    filePath: string,
    adapter: LibreOfficeAdapter,
    mutex: Mutex,
  ): Promise<DocumentSession> {
    if (LibreOfficeAdapter.needsBridge(filePath)) {
      const targetFormat = LibreOfficeAdapter.getBridgeFormat(filePath)!;
      logger(`Bridging ${filePath} → ${targetFormat} via LibreOffice`);
      const guard = await mutex.acquire();
      try {
        const { outputPath, tempDir } = await adapter.convertFile(filePath, targetFormat);
        return new DocumentSession(docId, filePath, outputPath, tempDir);
      } finally {
        guard.dispose();
      }
    }
    return new DocumentSession(docId, filePath, filePath, null);
  }

  getMetadataCache(): DocumentMetadata | null {
    return this.metadataCache;
  }

  setMetadataCache(metadata: DocumentMetadata): void {
    this.metadataCache = metadata;
  }

  /** Retrieve or populate the DOCX parsed document cache. */
  async getOrParseDocx<T>(parser: () => Promise<T>): Promise<T> {
    if (this._docxCache === null) {
      this._docxCache = await parser();
    }
    return this._docxCache as T;
  }

  /** Retrieve or populate the XLSX workbook cache. */
  async getOrParseXlsx<T>(parser: () => Promise<T>): Promise<T> {
    if (this._xlsxCache === null) {
      this._xlsxCache = await parser();
    }
    return this._xlsxCache as T;
  }

  /** Retrieve or populate the PPTX presentation cache. */
  async getOrParsePptx<T>(parser: () => Promise<T>): Promise<T> {
    if (this._pptxCache === null) {
      this._pptxCache = await parser();
    }
    return this._pptxCache as T;
  }

  /** Invalidate all parsed document caches. Call after any write operation. */
  invalidateCache(): void {
    this._docxCache = null;
    this._xlsxCache = null;
    this._pptxCache = null;
    this.metadataCache = null;
  }

  getDocumentType(): DocumentType {
    const ext = this.parsedExt;
    if (['.docx', '.dotx', '.doc', '.odt', '.rtf', '.txt'].includes(ext)) return 'writer';
    if (['.xlsx', '.xlsm', '.xls', '.ods', '.csv'].includes(ext)) return 'calc';
    if (['.pptx', '.ppt', '.odp'].includes(ext)) return 'impress';
    return 'unknown';
  }

  isClosed(): boolean {
    return this.closed;
  }

  async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;
    if (this.tempDir) {
      logger(`Cleaning up temp dir ${this.tempDir}`);
      await rm(this.tempDir, { recursive: true, force: true }).catch(() => {});
      this.tempDir = null;
    }
  }
}
