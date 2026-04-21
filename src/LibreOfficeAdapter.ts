import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { existsSync } from 'node:fs';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, extname, basename } from 'node:path';
import { logger } from './logger.js';
import { SOFFICE_TIMEOUT_MS } from './types.js';

const execFileAsync = promisify(execFile);

const KNOWN_SOFFICE_PATHS: Record<NodeJS.Platform, string[]> = {
  win32: [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  ],
  darwin: [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  ],
  linux: [
    '/usr/bin/soffice',
    '/usr/local/bin/soffice',
    '/snap/bin/libreoffice',
  ],
  // fallbacks for other platforms
  aix: [], freebsd: [], openbsd: [], sunos: [], android: [], cygwin: [], netbsd: [], haiku: [],
};

export type SupportedOutputFormat = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'html' | 'txt' | 'csv' | 'odt' | 'ods' | 'odp';

export interface ConvertResult {
  outputPath: string;
  tempDir: string;
}

export class LibreOfficeAdapter {
  private sofficePath: string;
  private available: boolean;

  constructor(sofficePath?: string) {
    this.sofficePath = sofficePath || process.env['SOFFICE_PATH'] || this.detectSofficePath();
    this.available = existsSync(this.sofficePath);
    if (!this.available) {
      logger(`LibreOffice not found at ${this.sofficePath}. Legacy format support and conversion will be unavailable.`);
    } else {
      logger(`LibreOffice found at ${this.sofficePath}`);
    }
  }

  isAvailable(): boolean {
    return this.available;
  }

  getSofficePath(): string {
    return this.sofficePath;
  }

  private detectSofficePath(): string {
    const platform = process.platform;
    const paths = KNOWN_SOFFICE_PATHS[platform] ?? [];
    for (const p of paths) {
      if (existsSync(p)) return p;
    }
    // Fallback: try PATH
    return 'soffice';
  }

  /**
   * Convert a document file to a different format using LibreOffice CLI.
   * Returns the path to the converted file and a temp directory (caller must clean up).
   */
  async convertFile(
    inputPath: string,
    outputFormat: SupportedOutputFormat,
    timeoutMs = SOFFICE_TIMEOUT_MS,
  ): Promise<ConvertResult> {
    this.requireAvailable();
    const tempDir = await mkdtemp(join(tmpdir(), 'lo-mcp-'));
    try {
      logger(`Converting ${inputPath} → ${outputFormat} in ${tempDir}`);
      await execFileAsync(
        this.sofficePath,
        ['--headless', '--convert-to', outputFormat, '--outdir', tempDir, inputPath],
        { timeout: timeoutMs },
      );
      const inputBase = basename(inputPath, extname(inputPath));
      const outputPath = join(tempDir, `${inputBase}.${outputFormat}`);
      if (!existsSync(outputPath)) {
        throw new Error(`LibreOffice conversion succeeded but output file not found: ${outputPath}`);
      }
      return { outputPath, tempDir };
    } catch (err) {
      // Clean up on error
      await rm(tempDir, { recursive: true, force: true }).catch(() => {});
      throw err;
    }
  }

  /**
   * Determine if a file needs LibreOffice bridge conversion before parsing.
   */
  static needsBridge(filePath: string): boolean {
    const ext = extname(filePath).toLowerCase();
    return BRIDGE_EXTENSIONS.has(ext);
  }

  /**
   * Get the target format for bridging a file.
   */
  static getBridgeFormat(filePath: string): SupportedOutputFormat | null {
    const ext = extname(filePath).toLowerCase();
    return BRIDGE_FORMAT_MAP[ext] ?? null;
  }

  private requireAvailable(): void {
    if (!this.available) {
      throw new Error(
        `LibreOffice is not available at ${this.sofficePath}. ` +
        `Install LibreOffice or set the SOFFICE_PATH environment variable.`,
      );
    }
  }
}

/** File extensions that need LibreOffice conversion before native parsing */
const BRIDGE_EXTENSIONS = new Set([
  '.doc', '.dot',       // Word 97-2003
  '.xls', '.xlt',       // Excel 97-2003
  '.ppt', '.pot',       // PowerPoint 97-2003
  '.odt', '.ott',       // OpenDocument Text
  '.ods', '.ots',       // OpenDocument Spreadsheet
  '.odp', '.otp',       // OpenDocument Presentation
  '.rtf',               // Rich Text Format
  '.wps',               // Works
  '.wpd',               // WordPerfect
]);

const BRIDGE_FORMAT_MAP: Record<string, SupportedOutputFormat> = {
  '.doc': 'docx',
  '.dot': 'docx',
  '.xls': 'xlsx',
  '.xlt': 'xlsx',
  '.ppt': 'pptx',
  '.pot': 'pptx',
  '.odt': 'docx',
  '.ott': 'docx',
  '.ods': 'xlsx',
  '.ots': 'xlsx',
  '.odp': 'pptx',
  '.otp': 'pptx',
  '.rtf': 'docx',
  '.wps': 'docx',
  '.wpd': 'docx',
};
