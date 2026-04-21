import { z } from 'zod';
import { defineTool, docIdSchema } from './ToolDefinition.js';
import { DocumentCategory } from './categories.js';
import { parsePptx } from '../parsers/PptxParser.js';
import { updateSlide, addSlide } from '../parsers/PptxOoxmlEditor.js';

export const presentationListSlides = defineTool({
  name: 'presentation_list_slides',
  description: 'List all slides in a presentation with their index and title. Token-efficient overview before reading content.',
  annotations: {
    category: DocumentCategory.PRESENTATION,
    readOnlyHint: true,
    title: 'List Slides',
  },
  schema: {
    ...docIdSchema,
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'impress') {
      throw new Error('presentation_list_slides only works with presentation documents.');
    }
    const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
    response.appendText(`${pptx.slideCount} slide(s):`);
    response.attachJson(pptx.slides.map(s => ({ index: s.index, title: s.title ?? '(no title)' })));
  },
});

export const presentationGetSlide = defineTool({
  name: 'presentation_get_slide',
  description: 'Get the full content of a specific slide: title, body text, and speaker notes.',
  annotations: {
    category: DocumentCategory.PRESENTATION,
    readOnlyHint: true,
    title: 'Get Slide Content',
  },
  schema: {
    ...docIdSchema,
    slideIndex: z.number().int().describe('Slide index (0-based)'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'impress') {
      throw new Error('presentation_get_slide only works with presentation documents.');
    }
    const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
    const slide = pptx.slides[request.params.slideIndex];
    if (!slide) {
      throw new Error(`Slide ${request.params.slideIndex} not found. Total slides: ${pptx.slideCount}`);
    }
    response.appendText(`Slide ${slide.index + 1}:`);
    response.attachJson(slide);
    const md = [
      `# ${slide.title ?? '(no title)'}`,
      slide.body ?? '*(no content)*',
      slide.notes ? `\n---\n**Notes:** ${slide.notes}` : '',
    ].join('\n\n');
    response.attachMarkdown(md);
  },
});

export const presentationGetNotes = defineTool({
  name: 'presentation_get_notes',
  description: 'Get speaker notes for all slides or a specific slide.',
  annotations: {
    category: DocumentCategory.PRESENTATION,
    readOnlyHint: true,
    title: 'Get Speaker Notes',
  },
  schema: {
    ...docIdSchema,
    slideIndex: z.number().int().optional().describe('Specific slide index (0-based). Omit to get all notes.'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'impress') {
      throw new Error('presentation_get_notes only works with presentation documents.');
    }
    const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
    const slides = request.params.slideIndex !== undefined
      ? [pptx.slides[request.params.slideIndex]].filter((s): s is typeof pptx.slides[number] => s !== undefined)
      : pptx.slides;

    const notes = slides.filter(s => s.notes).map(s => ({
      slideIndex: s.index,
      title: s.title,
      notes: s.notes,
    }));

    if (notes.length === 0) {
      response.appendText('No speaker notes found.');
    } else {
      response.appendText(`Notes for ${notes.length} slide(s):`);
      response.attachJson(notes);
    }
  },
});

export const presentationAddSlide = defineTool({
  name: 'presentation_add_slide',
  description: 'Add a new slide to a presentation.',
  annotations: {
    category: DocumentCategory.PRESENTATION,
    readOnlyHint: false,
    title: 'Add Slide',
  },
  schema: {
    ...docIdSchema,
    title: z.string().describe('Slide title'),
    body: z.string().optional().describe('Slide body text'),
    notes: z.string().optional().describe('Speaker notes'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'impress') {
      throw new Error('presentation_add_slide only works with presentation documents.');
    }

    const { title, body, notes } = request.params;
    await addSlide(session.parsedPath, title, body, notes);
    session.invalidateCache();

    const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
    response.appendText(`Slide added. Presentation now has ${pptx.slideCount} slide(s).`);
    response.attachJson({ newSlideIndex: pptx.slideCount - 1, title });
  },
});

export const presentationUpdateSlide = defineTool({
  name: 'presentation_update_slide',
  description: 'Update the title or body text of an existing slide.',
  annotations: {
    category: DocumentCategory.PRESENTATION,
    readOnlyHint: false,
    title: 'Update Slide',
  },
  schema: {
    ...docIdSchema,
    slideIndex: z.number().int().describe('Slide index (0-based)'),
    title: z.string().optional().describe('New title (omit to keep existing)'),
    body: z.string().optional().describe('New body text (omit to keep existing)'),
  },
  handler: async (request, response, context) => {
    const session = context.getDocument(request.params.docId);
    if (session.getDocumentType() !== 'impress') {
      throw new Error('presentation_update_slide only works with presentation documents.');
    }

    const { slideIndex, title, body } = request.params;
    if (title === undefined && body === undefined) {
      throw new Error('Provide at least one of title or body to update.');
    }

    await updateSlide(session.parsedPath, slideIndex, title, body);
    session.invalidateCache();

    const pptx = await session.getOrParsePptx(() => parsePptx(session.parsedPath));
    const slide = pptx.slides[slideIndex];
    response.appendText(`Slide ${slideIndex + 1} updated.`);
    response.attachJson({ slideIndex, title: slide?.title, body: slide?.body });
  },
});
