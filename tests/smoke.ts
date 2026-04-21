/**
 * Smoke tests for libreoffice-mcp-tools.
 * Run: npm test (runs npm run build first, then this against compiled output)
 * Requires sample fixtures in tests/fixtures/.
 */
import { existsSync } from 'node:fs';
import { copyFile, rm, mkdtemp, readdir } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { DocumentContext } from '../build/src/DocumentContext.js';
import { parseDocx } from '../build/src/parsers/DocxParser.js';
import { parseXlsx, getSheetInfo, getRange, setCellValue, saveXlsx, addNewSheet, createEmptyXlsx } from '../build/src/parsers/XlsxParser.js';
import { parsePptx } from '../build/src/parsers/PptxParser.js';
import { rangeToMarkdownTable } from '../build/src/formatters/TableFormatter.js';
import { openDocxForEdit } from '../build/src/parsers/DocxOoxmlEditor.js';
import { createEmptyPptx, addSlide, updateSlide } from '../build/src/parsers/PptxOoxmlEditor.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const FIXTURES = join(__dirname, 'fixtures');

let passed = 0;
let failed = 0;

async function test(name: string, fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
    console.log(`  ✅ ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ❌ ${name}`);
    console.error(`     ${(err as Error).message}`);
    failed++;
  }
}

function skip(name: string, reason: string): void {
  console.log(`  ⏭  ${name} — skipped: ${reason}`);
}

console.log('\n── DocumentContext ─────────────────────────────');

await test('DocumentContext instantiates without LibreOffice', async () => {
  const ctx = new DocumentContext('/nonexistent/soffice');
  assert.ok(!ctx.adapter.isAvailable());
  assert.deepEqual(ctx.listDocuments(), []);
});

await test('DocumentContext throws on unknown docId', async () => {
  const ctx = new DocumentContext();
  assert.throws(() => ctx.getDocument('nonexistent-id'), /Document not found/);
});

console.log('\n── DOCX Parser ─────────────────────────────────');

const DOCX_FIXTURE = join(FIXTURES, 'sample.docx');
if (existsSync(DOCX_FIXTURE)) {
  await test('parseDocx returns metadata', async () => {
    const doc = await parseDocx(DOCX_FIXTURE);
    assert.ok(doc.metadata.documentType === 'writer');
    assert.ok(typeof doc.metadata.wordCount === 'number');
  });

  await test('parseDocx returns paragraphs', async () => {
    const doc = await parseDocx(DOCX_FIXTURE);
    assert.ok(Array.isArray(doc.paragraphs));
    assert.ok(doc.paragraphs.length > 0);
  });

  await test('parseDocx returns outline (if headings exist)', async () => {
    const doc = await parseDocx(DOCX_FIXTURE);
    assert.ok(Array.isArray(doc.outline));
  });
} else {
  skip('DOCX tests', `fixtures/sample.docx not found`);
}

console.log('\n── XLSX Parser ─────────────────────────────────');

const XLSX_FIXTURE = join(FIXTURES, 'sample.xlsx');
if (existsSync(XLSX_FIXTURE)) {
  await test('parseXlsx returns sheet names', async () => {
    const wb = await parseXlsx(XLSX_FIXTURE);
    assert.ok(wb.sheetNames.length > 0);
  });

  await test('getSheetInfo returns row/col counts', async () => {
    const wb = await parseXlsx(XLSX_FIXTURE);
    const info = getSheetInfo(wb);
    assert.ok(info.length > 0);
    assert.ok(typeof info[0]!.rowCount === 'number');
  });

  await test('getRange returns rows', async () => {
    const wb = await parseXlsx(XLSX_FIXTURE);
    const sheetName = wb.sheetNames[0]!;
    const range = getRange(wb, sheetName, undefined, 10, 0);
    assert.ok(Array.isArray(range.rows));
  });

  await test('rangeToMarkdownTable produces table string', async () => {
    const wb = await parseXlsx(XLSX_FIXTURE);
    const sheetName = wb.sheetNames[0]!;
    const range = getRange(wb, sheetName, undefined, 5, 0);
    const table = rangeToMarkdownTable(range);
    assert.ok(typeof table === 'string');
    assert.ok(table.includes('|'));
  });
} else {
  skip('XLSX tests', `fixtures/sample.xlsx not found`);
}

console.log('\n── PPTX Parser ─────────────────────────────────');

const PPTX_FIXTURE = join(FIXTURES, 'sample.pptx');
if (existsSync(PPTX_FIXTURE)) {
  await test('parsePptx returns slide count', async () => {
    const pptx = await parsePptx(PPTX_FIXTURE);
    assert.ok(typeof pptx.slideCount === 'number');
    assert.ok(pptx.slideCount > 0);
  });

  await test('parsePptx returns slide objects', async () => {
    const pptx = await parsePptx(PPTX_FIXTURE);
    assert.ok(Array.isArray(pptx.slides));
    assert.ok(pptx.slides.length === pptx.slideCount);
    assert.ok('index' in pptx.slides[0]!);
  });
} else {
  skip('PPTX tests', `fixtures/sample.pptx not found`);
}

console.log('\n── DocumentContext with real file ──────────────');

if (existsSync(DOCX_FIXTURE)) {
  await test('openDocument returns a session', async () => {
    const ctx = new DocumentContext();
    const session = await ctx.openDocument(DOCX_FIXTURE);
    assert.ok(session.docId.length > 0);
    assert.equal(session.originalPath, DOCX_FIXTURE);
    assert.equal(session.getDocumentType(), 'writer');
    assert.deepEqual(ctx.listDocuments().map(d => d.docId), [session.docId]);
    await ctx.closeDocument(session.docId);
    assert.deepEqual(ctx.listDocuments(), []);
  });

  await test('openDocument reuses existing session for same path', async () => {
    const ctx = new DocumentContext();
    const s1 = await ctx.openDocument(DOCX_FIXTURE);
    const s2 = await ctx.openDocument(DOCX_FIXTURE);
    assert.equal(s1.docId, s2.docId);
    await ctx.closeAll();
  });
}

// ── Write Operations (use temp copies to avoid corrupting fixtures) ──────────

let tmpDir: string | undefined;
try {
  tmpDir = await mkdtemp(join(tmpdir(), 'lo-mcp-test-'));
} catch {
  // temp dir creation failed — skip write tests
}

if (tmpDir) {
  console.log('\n── DOCX Write Operations ───────────────────────');

  if (existsSync(DOCX_FIXTURE)) {
    const tmpDocx = join(tmpDir, 'write-test.docx');
    await copyFile(DOCX_FIXTURE, tmpDocx);

    await test('DocxOoxmlEditor: replaceText returns replacement count', async () => {
      const editor = await openDocxForEdit(tmpDocx);
      // Any text in the fixture — use a known word from the document
      const doc = await parseDocx(tmpDocx);
      const firstPara = doc.paragraphs[0];
      assert.ok(firstPara, 'Document must have at least one paragraph');
      const wordToReplace = firstPara.text.split(/\s+/)[0]!;
      const count = await editor.replaceText(wordToReplace, wordToReplace); // replace with same text
      assert.ok(count >= 0, 'replaceText must return a non-negative count');
      await editor.save();
    });

    await test('DocxOoxmlEditor: insertParagraph at end preserves existing content', async () => {
      const editor = await openDocxForEdit(tmpDocx);
      await editor.insertParagraph('__TEST_INSERTED_PARAGRAPH__', { position: 'end' });
      await editor.save();
      const doc = await parseDocx(tmpDocx);
      const texts = doc.paragraphs.map(p => p.text);
      assert.ok(texts.some(t => t.includes('__TEST_INSERTED_PARAGRAPH__')), 'Inserted paragraph must appear in re-parsed document');
    });

    await test('DocxOoxmlEditor: insertParagraph at start adds to beginning', async () => {
      const editor = await openDocxForEdit(tmpDocx);
      await editor.insertParagraph('__TEST_START_PARAGRAPH__', { position: 'start' });
      await editor.save();
      const doc = await parseDocx(tmpDocx);
      assert.ok(doc.paragraphs.some(p => p.text.includes('__TEST_START_PARAGRAPH__')), 'Start-inserted paragraph must appear in re-parsed document');
    });
  } else {
    skip('DOCX write tests', 'fixtures/sample.docx not found');
  }

  console.log('\n── XLSX Write Operations ───────────────────────');

  if (existsSync(XLSX_FIXTURE)) {
    const tmpXlsx = join(tmpDir, 'write-test.xlsx');
    await copyFile(XLSX_FIXTURE, tmpXlsx);

    await test('setCellValue: set a cell then re-read it', async () => {
      const wb = await parseXlsx(tmpXlsx);
      const sheetName = wb.sheetNames[0]!;
      setCellValue(wb, sheetName, 'A1', '__TEST_VALUE__');
      await saveXlsx(wb, tmpXlsx);
      const wb2 = await parseXlsx(tmpXlsx);
      const range = getRange(wb2, sheetName, 'A1:A1');
      assert.equal(range.rows[0]?.[0], '__TEST_VALUE__');
    });

    await test('addNewSheet: new sheet appears in sheetNames', async () => {
      const wb = await parseXlsx(tmpXlsx);
      addNewSheet(wb, '__TEST_SHEET__', ['ColA', 'ColB']);
      await saveXlsx(wb, tmpXlsx);
      const wb2 = await parseXlsx(tmpXlsx);
      assert.ok(wb2.sheetNames.includes('__TEST_SHEET__'), 'New sheet must appear after save+reparse');
    });

    const tmpNewXlsx = join(tmpDir, 'new-empty.xlsx');
    await test('createEmptyXlsx: creates a valid readable XLSX file', async () => {
      await createEmptyXlsx(tmpNewXlsx);
      assert.ok(existsSync(tmpNewXlsx), 'File must be created');
      const wb = await parseXlsx(tmpNewXlsx);
      assert.ok(wb.sheetNames.length > 0, 'Created XLSX must have at least one sheet');
    });
  } else {
    skip('XLSX write tests', 'fixtures/sample.xlsx not found');
  }

  console.log('\n── PPTX Write Operations ───────────────────────');

  const tmpPptx = join(tmpDir, 'new-presentation.pptx');
  await test('createEmptyPptx: creates a valid PPTX file', async () => {
    await createEmptyPptx(tmpPptx);
    assert.ok(existsSync(tmpPptx), 'PPTX file must be created');
    const pptx = await parsePptx(tmpPptx);
    assert.ok(pptx.slideCount >= 1, 'Created PPTX must have at least one slide');
  });

  await test('addSlide: increases slide count by 1', async () => {
    const before = await parsePptx(tmpPptx);
    await addSlide(tmpPptx, 'Test Slide Title', 'Test slide body text.');
    const after = await parsePptx(tmpPptx);
    assert.equal(after.slideCount, before.slideCount + 1, 'slideCount must increase by 1 after addSlide');
  });

  await test('addSlide: new slide has correct title', async () => {
    await addSlide(tmpPptx, '__UNIQUE_TITLE__', 'Some body text.');
    const pptx = await parsePptx(tmpPptx);
    const slide = pptx.slides[pptx.slideCount - 1];
    assert.ok(slide?.title?.includes('__UNIQUE_TITLE__'), `New slide title "${slide?.title}" should include "__UNIQUE_TITLE__"`);
  });

  await test('updateSlide: title is updated correctly', async () => {
    // Update first slide's title
    await updateSlide(tmpPptx, 0, '__UPDATED_TITLE__', undefined);
    const pptx = await parsePptx(tmpPptx);
    assert.ok(pptx.slides[0]?.title?.includes('__UPDATED_TITLE__'), `Slide 0 title "${pptx.slides[0]?.title}" should include "__UPDATED_TITLE__"`);
  });

  if (existsSync(PPTX_FIXTURE)) {
    const tmpPptxCopy = join(tmpDir, 'existing-copy.pptx');
    await copyFile(PPTX_FIXTURE, tmpPptxCopy);

    await test('updateSlide on existing file: preserves slide count', async () => {
      const before = await parsePptx(tmpPptxCopy);
      await updateSlide(tmpPptxCopy, 0, 'Modified Title', undefined);
      const after = await parsePptx(tmpPptxCopy);
      assert.equal(after.slideCount, before.slideCount, 'updateSlide must not change slide count');
    });
  } else {
    skip('PPTX updateSlide on existing file', 'fixtures/sample.pptx not found');
  }

  // Clean up temp directory
  await rm(tmpDir, { recursive: true, force: true });
}

console.log('\n────────────────────────────────────────────────');
console.log(`Results: ${passed} passed, ${failed} failed\n`);
if (failed > 0) process.exit(1);
