// tests/create-fixtures.mjs
// Creates minimal but valid fixture files for smoke tests
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = fileURLToPath(new URL('.', import.meta.url));
const FIXTURES = join(__dirname, 'fixtures');

// Create sample.xlsx using xlsx package
import xlsx from 'xlsx';
const wb = xlsx.utils.book_new();
const wsData = [
  ['Name', 'Department', 'Salary', 'Years'],
  ['Alice', 'Engineering', 95000, 5],
  ['Bob', 'Marketing', 72000, 3],
  ['Carol', 'Engineering', 88000, 7],
  ['David', 'HR', 65000, 2],
];
const ws = xlsx.utils.aoa_to_sheet(wsData);
xlsx.utils.book_append_sheet(wb, ws, 'Employees');
const ws2 = xlsx.utils.aoa_to_sheet([
  ['Quarter', 'Revenue', 'Expenses'],
  ['Q1', 150000, 90000],
  ['Q2', 175000, 95000],
  ['Q3', 160000, 88000],
  ['Q4', 200000, 110000],
]);
xlsx.utils.book_append_sheet(wb, ws2, 'Financials');
xlsx.writeFile(wb, join(FIXTURES, 'sample.xlsx'));
console.log('Created sample.xlsx');

// Create sample.docx using docx package
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
const doc = new Document({
  sections: [{
    children: [
      new Paragraph({ text: 'Sample Document', heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ children: [new TextRun('This is an introduction paragraph with some sample text to test the MCP tools.')] }),
      new Paragraph({ text: 'Section 1: Overview', heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ children: [new TextRun('The LibreOffice MCP Tools project provides AI agents with token-efficient access to Office documents.')] }),
      new Paragraph({ text: 'Section 2: Features', heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ children: [new TextRun('Key features include: outline-first navigation, range-based reading, and pagination support.')] }),
      new Paragraph({ text: 'Subsection 2.1: Reading', heading: HeadingLevel.HEADING_3 }),
      new Paragraph({ children: [new TextRun('Use document_get_outline to see structure before reading content.')] }),
      new Paragraph({ text: 'Conclusion', heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ children: [new TextRun('This sample document is used for testing the LibreOffice MCP tools smoke tests.')] }),
    ],
  }],
});
const docBuffer = await Packer.toBuffer(doc);
await writeFile(join(FIXTURES, 'sample.docx'), docBuffer);
console.log('Created sample.docx');

// Create a minimal valid PPTX (ZIP with required XML files)
import JSZip from 'jszip';
const pptx = new JSZip();
pptx.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`);
pptx.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);
pptx.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>`);
pptx.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>`);
pptx.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Introduction to LibreOffice MCP</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>This tool enables AI agents to work with Office documents efficiently</a:t></a:r></a:p>
          <a:p><a:r><a:t>Token-efficient design for LLM usage</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`);
pptx.file('ppt/slides/slide2.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Key Features</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Outline-first document navigation</a:t></a:r></a:p>
          <a:p><a:r><a:t>Range-based cell and paragraph access</a:t></a:r></a:p>
          <a:p><a:r><a:t>22 specialized MCP tools</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`);
pptx.file('ppt/notesSlides/notesSlide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:txBody><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>Speaker notes for slide 1: Welcome the audience and introduce the MCP concept.</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>`);
const pptxBuffer = await pptx.generateAsync({ type: 'nodebuffer' });
await writeFile(join(FIXTURES, 'sample.pptx'), pptxBuffer);
console.log('Created sample.pptx');

console.log('\nAll fixtures created successfully in tests/fixtures/');
