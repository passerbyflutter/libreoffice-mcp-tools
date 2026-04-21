/**
 * OOXML-based PPTX editor using JSZip.
 *
 * Reads and writes PPTX XML files directly, preserving all formatting,
 * animations, transitions, images, and embedded objects not explicitly targeted.
 *
 * Supports:
 *   - updateSlide: Replace title and/or body text in an existing slide
 *   - addSlide: Append a new slide with title and optional body/notes
 *   - createEmptyPptx: Build a minimal valid PPTX file from scratch
 */

import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

// ─── Public API ────────────────────────────────────────────────────────────────

/**
 * Update the title and/or body text of an existing slide (0-based index).
 * Changes are saved in-place to `filePath`.
 */
export async function updateSlide(
  filePath: string,
  slideIndex: number,
  title: string | undefined,
  body: string | undefined,
): Promise<void> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  const slideTargets = await resolveSlideTargets(zip);
  const slideTarget = slideTargets[slideIndex];
  if (!slideTarget) {
    throw new Error(`Slide index ${slideIndex} is out of range (${slideTargets.length} slides found).`);
  }

  const slideFilePath = `ppt/${slideTarget}`;
  const slideFile = zip.file(slideFilePath);
  if (!slideFile) throw new Error(`Slide file not found in PPTX: ${slideFilePath}`);

  let slideXml = await slideFile.async('string');

  if (title !== undefined) {
    slideXml = replacePlaceholderText(slideXml, ['title', 'ctrTitle'], null, title);
  }
  if (body !== undefined) {
    // Content placeholder uses idx="1" without type; some use type="body"
    let updated = replacePlaceholderText(slideXml, [], '1', body);
    if (updated === slideXml) {
      updated = replacePlaceholderText(slideXml, ['body'], null, body);
    }
    slideXml = updated;
  }

  zip.file(slideFilePath, slideXml);
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await writeFile(filePath, output);
}

/**
 * Add a new slide to an existing PPTX file.
 * The slide is appended at the end of the presentation.
 * Changes are saved in-place to `filePath`.
 */
export async function addSlide(
  filePath: string,
  title: string,
  body?: string,
  notes?: string,
): Promise<void> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  // Determine the next slide number
  const existingSlides = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  const existingNums = existingSlides.map(f => parseInt(f.match(/slide(\d+)\.xml$/)![1]!, 10));
  const nextNum = existingNums.length > 0 ? Math.max(...existingNums) + 1 : 1;

  // Get slideLayout rId from an existing slide's rels (or default to rId1)
  const layoutRelId = await getSlideLayoutRelId(zip, existingSlides[0]);

  // Build the new slide XML
  const slideXml = buildSlideXml(title, body);
  const notesXml = notes ? buildNotesXml(notes, nextNum) : null;

  // Add slide files to zip
  const slideZipPath = `ppt/slides/slide${nextNum}.xml`;
  const slideRelsZipPath = `ppt/slides/_rels/slide${nextNum}.xml.rels`;
  zip.file(slideZipPath, slideXml);
  zip.file(slideRelsZipPath, buildSlideRels(layoutRelId));
  if (notesXml) {
    zip.file(`ppt/notesSlides/notesSlide${nextNum}.xml`, notesXml);
    zip.file(`ppt/notesSlides/_rels/notesSlide${nextNum}.xml.rels`,
      buildNotesRels(nextNum));
  }

  // Update presentation.xml rels: add new rId → slides/slideN.xml
  const newSlideRId = await addSlideRelationship(zip, nextNum);

  // Update presentation.xml: add <p:sldId> to <p:sldIdLst>
  await addSlideIdToPresentation(zip, newSlideRId);

  // Update [Content_Types].xml: add Override for new slide
  await addContentTypeOverride(zip, nextNum, notesXml !== null);

  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await writeFile(filePath, output);
}

/**
 * Create a minimal valid PPTX file with one empty slide.
 * The file is written to `filePath`.
 */
export async function createEmptyPptx(filePath: string): Promise<void> {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', MINIMAL_CONTENT_TYPES);
  zip.file('_rels/.rels', MINIMAL_RELS);
  zip.file('ppt/presentation.xml', MINIMAL_PRESENTATION);
  zip.file('ppt/_rels/presentation.xml.rels', MINIMAL_PRESENTATION_RELS);
  zip.file('ppt/slides/slide1.xml', buildSlideXml('', undefined));
  zip.file('ppt/slides/_rels/slide1.xml.rels', buildSlideRels('rId1'));
  zip.file('ppt/slideLayouts/slideLayout1.xml', MINIMAL_SLIDE_LAYOUT);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', MINIMAL_SLIDE_LAYOUT_RELS);
  zip.file('ppt/slideMasters/slideMaster1.xml', MINIMAL_SLIDE_MASTER);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', MINIMAL_SLIDE_MASTER_RELS);
  zip.file('ppt/theme/theme1.xml', MINIMAL_THEME);

  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await writeFile(filePath, output);
}

// ─── Slide XML builders ─────────────────────────────────────────────────────

function buildSlideXml(title: string, body: string | undefined): string {
  const titleParagraphs = buildTxBodyParagraphs(title);
  const bodyParagraphs = body !== undefined
    ? buildTxBodyParagraphs(body)
    : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/>${titleParagraphs}</p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/>${bodyParagraphs}</p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClr/></p:clrMapOvr>
</p:sld>`;
}

function buildNotesXml(notes: string, _slideNum: number): string {
  const paragraphs = buildTxBodyParagraphs(notes);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClr/></p:clrMapOvr>
</p:notes>`;
}

function buildSlideRels(layoutRelId: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${layoutRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
}

function buildNotesRels(slideNum: number): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${slideNum}.xml"/>
</Relationships>`;
}

/** Build one or more <a:p> elements from text. Multi-line text splits into multiple paragraphs. */
function buildTxBodyParagraphs(text: string): string {
  if (!text) return '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>';
  return text.split('\n').map(line => {
    if (!line.trim()) return '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>';
    return `<a:p><a:r><a:t>${encodeXmlEntities(line)}</a:t></a:r></a:p>`;
  }).join('');
}

function encodeXmlEntities(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ─── Placeholder text replacement ──────────────────────────────────────────

/**
 * Find a <p:sp> containing a <p:ph> with matching type(s) or idx,
 * then replace its <p:txBody> content with `text`.
 *
 * @param phTypes - acceptable values for `type="..."` attribute (empty = don't filter by type)
 * @param idx - required `idx="..."` value, or null to not require
 */
function replacePlaceholderText(
  slideXml: string,
  phTypes: string[],
  idx: string | null,
  text: string,
): string {
  // Find all <p:sp>...</p:sp> blocks and locate the one matching our placeholder
  let searchPos = 0;
  while (true) {
    const spStart = slideXml.indexOf('<p:sp>', searchPos);
    if (spStart === -1) break;

    const spEnd = slideXml.indexOf('</p:sp>', spStart);
    if (spEnd === -1) break;

    const spBlock = slideXml.slice(spStart, spEnd + '</p:sp>'.length);

    if (matchesPlaceholder(spBlock, phTypes, idx)) {
      const updatedBlock = replaceTxBodyContent(spBlock, text);
      return slideXml.slice(0, spStart) + updatedBlock + slideXml.slice(spEnd + '</p:sp>'.length);
    }

    searchPos = spEnd + 1;
  }

  return slideXml; // no matching placeholder found, return unchanged
}

/** Check if an <p:sp> block contains a placeholder matching given type(s)/idx. */
function matchesPlaceholder(spBlock: string, phTypes: string[], idx: string | null): boolean {
  // Must have a <p:ph> element inside <p:nvPr>
  if (!spBlock.includes('<p:ph')) return false;

  // Extract the <p:ph .../> element
  const phMatch = spBlock.match(/<p:ph([^/]*)\/?>/);
  if (!phMatch) return false;
  const phAttrs = phMatch[1]!;

  // Check type requirement
  if (phTypes.length > 0) {
    const typeMatch = phAttrs.match(/type="([^"]+)"/);
    if (!typeMatch || !phTypes.includes(typeMatch[1]!)) return false;
  }

  // Check idx requirement
  if (idx !== null) {
    const idxMatch = phAttrs.match(/idx="([^"]+)"/);
    if (!idxMatch || idxMatch[1] !== idx) return false;
  }

  return true;
}

/** Replace the <p:txBody>...</p:txBody> content inside an <p:sp> block. */
function replaceTxBodyContent(spBlock: string, text: string): string {
  const txBodyStart = spBlock.indexOf('<p:txBody>');
  const txBodyEnd = spBlock.indexOf('</p:txBody>');
  if (txBodyStart === -1 || txBodyEnd === -1) return spBlock;

  const newTxBody = `<p:txBody><a:bodyPr/><a:lstStyle/>${buildTxBodyParagraphs(text)}</p:txBody>`;
  return spBlock.slice(0, txBodyStart) + newTxBody + spBlock.slice(txBodyEnd + '</p:txBody>'.length);
}

// ─── Presentation mutation helpers ─────────────────────────────────────────

/**
 * Resolve ordered slide file targets from ppt/presentation.xml + _rels.
 * Returns paths relative to ppt/ (e.g. "slides/slide1.xml").
 */
async function resolveSlideTargets(zip: JSZip): Promise<string[]> {
  const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
  const presentationFile = zip.file('ppt/presentation.xml');
  if (!relsFile || !presentationFile) return [];

  const relsXml = await relsFile.async('string');
  const presentationXml = await presentationFile.async('string');

  const rIdToTarget: Record<string, string> = {};
  for (const match of relsXml.matchAll(/<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"[^>]*\/?>/g)) {
    const target = match[2]!;
    if (target.includes('slides/slide')) rIdToTarget[match[1]!] = target;
  }

  const sldIdLstMatch = presentationXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
  if (!sldIdLstMatch) return [];

  const rIds: string[] = [];
  for (const match of sldIdLstMatch[1]!.matchAll(/r:id="([^"]+)"/g)) {
    rIds.push(match[1]!);
  }

  return rIds.map(rId => rIdToTarget[rId]).filter((t): t is string => t !== undefined);
}

/** Get the slideLayout relationship Id from a slide's _rels file. */
async function getSlideLayoutRelId(zip: JSZip, slideZipPath?: string): Promise<string> {
  if (slideZipPath) {
    const slideNum = slideZipPath.match(/slide(\d+)\.xml$/)?.[1];
    if (slideNum) {
      const relsFile = zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`);
      if (relsFile) {
        const relsXml = await relsFile.async('string');
        const match = relsXml.match(/<Relationship[^>]+Type="[^"]*slideLayout[^"]*"[^>]+Id="([^"]+)"/);
        if (match) return match[1]!;
      }
    }
  }
  return 'rId1'; // fallback
}

/** Add a relationship for a new slide to ppt/_rels/presentation.xml.rels. Returns the new rId. */
async function addSlideRelationship(zip: JSZip, slideNum: number): Promise<string> {
  const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
  if (!relsFile) throw new Error('Missing ppt/_rels/presentation.xml.rels');

  let relsXml = await relsFile.async('string');

  // Find max existing rId number
  const rIdNums = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1]!, 10));
  const nextRIdNum = rIdNums.length > 0 ? Math.max(...rIdNums) + 1 : 1;
  const newRId = `rId${nextRIdNum}`;

  const newRel = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`;
  relsXml = relsXml.replace('</Relationships>', `  ${newRel}\n</Relationships>`);
  zip.file('ppt/_rels/presentation.xml.rels', relsXml);

  return newRId;
}

/** Add <p:sldId> entry to <p:sldIdLst> in ppt/presentation.xml. */
async function addSlideIdToPresentation(zip: JSZip, newSlideRId: string): Promise<void> {
  const presentationFile = zip.file('ppt/presentation.xml');
  if (!presentationFile) throw new Error('Missing ppt/presentation.xml');

  let xml = await presentationFile.async('string');

  // Find max sldId value
  const sldIdNums = [...xml.matchAll(/p:sldId[^>]+id="(\d+)"/g)].map(m => parseInt(m[1]!, 10));
  const nextId = sldIdNums.length > 0 ? Math.max(...sldIdNums) + 1 : 256;

  const newEntry = `<p:sldId id="${nextId}" r:id="${newSlideRId}"/>`;

  if (xml.includes('</p:sldIdLst>')) {
    xml = xml.replace('</p:sldIdLst>', `  ${newEntry}\n  </p:sldIdLst>`);
  } else if (xml.includes('<p:sldIdLst/>')) {
    xml = xml.replace('<p:sldIdLst/>', `<p:sldIdLst>\n    ${newEntry}\n  </p:sldIdLst>`);
  } else {
    // No sldIdLst — insert before </p:presentation>
    xml = xml.replace('</p:presentation>', `<p:sldIdLst>\n    ${newEntry}\n  </p:sldIdLst>\n</p:presentation>`);
  }

  zip.file('ppt/presentation.xml', xml);
}

/** Add content type Override entries for the new slide (and optionally notes slide). */
async function addContentTypeOverride(zip: JSZip, slideNum: number, hasNotes: boolean): Promise<void> {
  const ctFile = zip.file('[Content_Types].xml');
  if (!ctFile) throw new Error('Missing [Content_Types].xml');

  let xml = await ctFile.async('string');
  const slideOverride = `<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  xml = xml.replace('</Types>', `  ${slideOverride}\n</Types>`);

  if (hasNotes) {
    const notesOverride = `<Override PartName="/ppt/notesSlides/notesSlide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
    xml = xml.replace('</Types>', `  ${notesOverride}\n</Types>`);
  }

  zip.file('[Content_Types].xml', xml);
}

// ─── Minimal PPTX templates ─────────────────────────────────────────────────

const MINIMAL_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;

const MINIMAL_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const MINIMAL_PRESENTATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;

const MINIMAL_PRESENTATION_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`;

const MINIMAL_SLIDE_LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">
  <p:cSld name="Title Slide"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClr/></p:clrMapOvr>
</p:sldLayout>`;

const MINIMAL_SLIDE_LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const MINIMAL_SLIDE_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lstStyle/></p:titleStyle>
    <p:bodyStyle><a:lstStyle/></p:bodyStyle>
    <p:otherStyle><a:lstStyle/></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`;

const MINIMAL_SLIDE_MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

const MINIMAL_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1>
      <a:lt1><a:sysClr lastClr="ffffff" val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="1F3864"/></a:dk2>
      <a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill><a:gsLst/><a:lin ang="5400000" scaled="0"/></a:gradFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
