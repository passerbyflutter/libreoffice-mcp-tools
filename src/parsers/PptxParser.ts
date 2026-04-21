/**
 * PPTX parser using JSZip + XML parsing.
 * Handles .pptx files. Legacy .ppt and .odp are bridged via LibreOffice.
 */
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { SlideContent } from '../types.js';

export interface PptxPresentation {
  slideCount: number;
  slides: SlideContent[];
}

export async function parsePptx(filePath: string): Promise<PptxPresentation> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  // Resolve slide order via relationship file (reliable, avoids heuristic regex)
  const slideTargets = await resolveSlideTargets(zip);

  const slides: SlideContent[] = [];

  for (let i = 0; i < slideTargets.length; i++) {
    const slideTarget = slideTargets[i]!;
    // Targets are relative to ppt/, e.g. "slides/slide1.xml"
    const slideFile = zip.file(`ppt/${slideTarget}`);
    if (!slideFile) continue;

    const slideXml = await slideFile.async('string');

    // Derive notes path from slide path: slides/slide1.xml → notesSlides/notesSlide1.xml
    const slideNum = slideTarget.match(/slide(\d+)\.xml$/)?.[1];
    const notesFile = slideNum ? zip.file(`ppt/notesSlides/notesSlide${slideNum}.xml`) : null;
    const notesXml = notesFile ? await notesFile.async('string') : undefined;

    const slide = parseSlideXml(slideXml, i, notesXml);
    slides.push(slide);
  }

  return { slideCount: slides.length, slides };
}

/**
 * Resolve the ordered list of slide file targets by reading:
 * 1. ppt/presentation.xml — for the ordered sldIdLst with r:id refs
 * 2. ppt/_rels/presentation.xml.rels — to map r:id → slide file Target
 */
async function resolveSlideTargets(zip: JSZip): Promise<string[]> {
  const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
  const presentationFile = zip.file('ppt/presentation.xml');

  if (!relsFile || !presentationFile) return [];

  const relsXml = await relsFile.async('string');
  const presentationXml = await presentationFile.async('string');

  // Build map: rId → Target (e.g. "rId2" → "slides/slide1.xml")
  const rIdToTarget = parseRelationships(relsXml);

  // Extract slide rIds in order from <p:sldIdLst>
  const orderedSlideRIds = extractOrderedSlideRIds(presentationXml);

  return orderedSlideRIds
    .map(rId => rIdToTarget[rId])
    .filter((t): t is string => t !== undefined);
}

/**
 * Parse a .rels XML file and return a map of Id → Target for slide relationships only.
 */
function parseRelationships(relsXml: string): Record<string, string> {
  const map: Record<string, string> = {};
  // Match <Relationship ... Id="rId1" ... Target="slides/slide1.xml" .../>
  const pattern = /<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"[^>]*\/?>/g;
  for (const match of relsXml.matchAll(pattern)) {
    const id = match[1]!;
    const target = match[2]!;
    // Only include actual slide references (Target contains "slides/slide")
    if (target.includes('slides/slide')) {
      map[id] = target;
    }
  }
  return map;
}

/**
 * Extract slide rIds in presentation order from <p:sldIdLst>.
 */
function extractOrderedSlideRIds(presentationXml: string): string[] {
  // Find the <p:sldIdLst> block
  const sldIdLstMatch = presentationXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
  if (!sldIdLstMatch) return [];

  const sldIdLst = sldIdLstMatch[1]!;
  const rIds: string[] = [];

  // Each slide: <p:sldId id="..." r:id="rId2"/>
  for (const match of sldIdLst.matchAll(/r:id="([^"]+)"/g)) {
    rIds.push(match[1]!);
  }

  return rIds;
}

function parseSlideXml(xml: string, index: number, notesXml?: string): SlideContent {
  const title = extractTitle(xml);
  const body = extractBody(xml, title);
  const notes = notesXml ? extractNotes(notesXml) : undefined;

  return { index, title, body, notes };
}

function extractTextRuns(xml: string): string[] {
  const texts: string[] = [];
  // Match <a:t>text</a:t> elements
  const matches = xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g);
  for (const match of matches) {
    const text = match[1]?.trim();
    if (text) texts.push(decodeXmlEntities(text));
  }
  return texts;
}

function extractTitle(xml: string): string | undefined {
  // Title is in <p:sp> with <p:ph type="title"> or <p:ph type="ctrTitle">
  const titleMatch = xml.match(/<p:sp>(?:(?!<p:sp>).)*?<p:ph[^>]*type="(?:title|ctrTitle)"[^>]*>(?:(?!<p:sp>).)*?<\/p:sp>/s);
  if (!titleMatch) return undefined;
  const texts = extractTextRuns(titleMatch[0]);
  return texts.join(' ').trim() || undefined;
}

function extractBody(xml: string, title?: string): string | undefined {
  // Extract paragraphs from body placeholder
  const bodyMatch = xml.match(/<p:sp>(?:(?!<p:sp>).)*?<p:ph[^>]*type="body"[^>]*>(?:(?!<p:sp>).)*?<\/p:sp>/s);
  if (bodyMatch) {
    const bodyTexts = extractTextRuns(bodyMatch[0]);
    return bodyTexts.join('\n').trim() || undefined;
  }

  // Fallback: all text that isn't the title
  const allTexts = extractTextRuns(xml);
  const titleText = title ?? '';
  const bodyTexts = allTexts.filter(t => t !== titleText && !titleText.includes(t));
  return bodyTexts.join('\n').trim() || undefined;
}

function extractNotes(notesXml: string): string | undefined {
  const texts = extractTextRuns(notesXml);
  return texts.join('\n').trim() || undefined;
}

function decodeXmlEntities(text: string): string {
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}
