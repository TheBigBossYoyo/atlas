import { promises as fs } from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'

import { Document, Packer, Paragraph, TextRun, PageBreak } from 'docx'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()
const outputDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

/**
 * Hand-rolls a minimal, valid single-section PDF (no external dependency
 * pulls in a real PDF writer just for fixtures) with one page per entry in
 * `pages`, each with its own MediaBox size and a Tj-drawn text string.
 * `withOutline` additionally emits a flat bookmark per page — the P12/PDF-16
 * e2e coverage wants a document that exercises pdf.js's real outline path
 * (PdfViewer's `buildOutlineNavItems`), not only the no-outline fallback the
 * single-page `sample.pdf` fixture below exercises.
 */
function buildPdfDocument(pages, { withOutline = false } = {}) {
  const objects = []
  const catalogNum = 1
  const pagesNum = 2
  const fontNum = 3
  const outlinesNum = 4
  const firstContentObjNum = 5
  const pageObjNum = (index) => firstContentObjNum + index * 2
  const contentObjNum = (index) => firstContentObjNum + index * 2 + 1
  const outlineItemObjNum = (index) => firstContentObjNum + pages.length * 2 + index

  objects[catalogNum - 1] = withOutline
    ? `<< /Type /Catalog /Pages ${pagesNum} 0 R /Outlines ${outlinesNum} 0 R >>`
    : `<< /Type /Catalog /Pages ${pagesNum} 0 R >>`
  objects[pagesNum - 1] =
    `<< /Type /Pages /Kids [${pages.map((_, i) => `${pageObjNum(i)} 0 R`).join(' ')}] /Count ${pages.length} >>`
  objects[fontNum - 1] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
  objects[outlinesNum - 1] = withOutline
    ? `<< /Type /Outlines /First ${outlineItemObjNum(0)} 0 R /Last ${outlineItemObjNum(pages.length - 1)} 0 R /Count ${pages.length} >>`
    : '<< /Type /Outlines /Count 0 >>'

  pages.forEach((page, index) => {
    const contentStream = `BT\n/F1 18 Tf\n36 ${Math.max(36, page.height - 60)} Td\n(${page.text}) Tj\nET`
    objects[pageObjNum(index) - 1] =
      `<< /Type /Page /Parent ${pagesNum} 0 R /MediaBox [0 0 ${page.width} ${page.height}] ` +
      `/Contents ${contentObjNum(index)} 0 R /Resources << /Font << /F1 ${fontNum} 0 R >> >> >>`
    objects[contentObjNum(index) - 1] =
      `<< /Length ${contentStream.length} >>\nstream\n${contentStream}\nendstream`
  })

  if (withOutline) {
    pages.forEach((page, index) => {
      const prev = index > 0 ? `/Prev ${outlineItemObjNum(index - 1)} 0 R ` : ''
      const next = index < pages.length - 1 ? `/Next ${outlineItemObjNum(index + 1)} 0 R ` : ''
      objects[outlineItemObjNum(index) - 1] =
        `<< /Title (Bookmark ${index + 1}) /Parent ${outlinesNum} 0 R ` +
        `/Dest [${pageObjNum(index)} 0 R /Fit] ${prev}${next}>>`
    })
  }

  let pdf = '%PDF-1.4\n'
  const offsets = [0]

  for (let index = 0; index < objects.length; index += 1) {
    offsets.push(Buffer.byteLength(pdf, 'utf8'))
    pdf += `${index + 1} 0 obj\n${objects[index]}\nendobj\n`
  }

  const xrefOffset = Buffer.byteLength(pdf, 'utf8')
  pdf += `xref\n0 ${objects.length + 1}\n`
  pdf += '0000000000 65535 f \n'

  for (const offset of offsets.slice(1)) {
    pdf += `${String(offset).padStart(10, '0')} 00000 n \n`
  }

  pdf += `trailer\n<< /Root 1 0 R /Size ${objects.length + 1} >>\nstartxref\n${xrefOffset}\n%%EOF\n`
  return Buffer.from(pdf, 'utf8')
}

function makePdfBuffer() {
  return buildPdfDocument([{ width: 300, height: 144, text: 'Atlas PDF fixture' }])
}

/** A 5-page document mixing portrait and landscape page sizes (PDF-09/P8),
 * each page's own findable text (PDF-06/P6), and a real outline so pdf.js's
 * outline-based nav path (not just the flat fallback) gets exercised. */
function makeMultiPagePdfBuffer() {
  const pages = [
    { width: 300, height: 500, text: 'Atlas multipage fixture page one' },
    { width: 500, height: 300, text: 'Atlas multipage fixture page two landscape' },
    { width: 300, height: 500, text: 'Findable needle on page three' },
    { width: 500, height: 300, text: 'Atlas multipage fixture page four landscape' },
    { width: 300, height: 500, text: 'Atlas multipage fixture page five' },
  ]
  return buildPdfDocument(pages, { withOutline: true })
}

function makeWorkbookBuffer(bookType) {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet([
    ['Name', 'Value'],
    ['Atlas', 2],
    ['Phase', 1],
  ])
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')
  return XLSX.write(workbook, { type: 'buffer', bookType })
}

async function makeDocxBuffer() {
  const document = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            children: [new TextRun({ text: 'Atlas DOCX fixture', bold: true })],
          }),
          new Paragraph('Smoke test document.'),
        ],
      },
    ],
  })

  return Packer.toBuffer(document)
}

/** A real 2-page DOCX (an explicit `PageBreak` run, not just enough text to
 * overflow one page) — X1/export.spec.ts exports this to PDF and asserts
 * the result has exactly 2 pages, proving `exportDocxPdf` captures every
 * page rather than whatever the live DOM happened to have on screen. */
async function makeMultiPageDocxBuffer() {
  const document = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: 'Atlas multipage DOCX fixture — page one', bold: true }),
              new PageBreak(),
            ],
          }),
          new Paragraph({ children: [new TextRun({ text: 'Page two content' })] }),
        ],
      },
    ],
  })

  return Packer.toBuffer(document)
}

/** A real 2-sheet workbook (both visible) — X1/export.spec.ts exports this
 * to PDF and asserts one page-group per sheet. */
function makeMultiSheetWorkbookBuffer(bookType) {
  const workbook = XLSX.utils.book_new()
  const sheet1 = XLSX.utils.aoa_to_sheet([
    ['Name', 'Value'],
    ['Atlas', 2],
  ])
  const sheet2 = XLSX.utils.aoa_to_sheet([
    ['City', 'Country'],
    ['Berlin', 'Germany'],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet1, 'First')
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Second')
  return XLSX.write(workbook, { type: 'buffer', bookType })
}

async function makePptxBuffer() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Atlas PPTX fixture</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`)
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** A real 2-slide PPTX — X1/export.spec.ts exports this to PDF and asserts
 * the result has exactly 2 pages (SLD-01/UX-02's actual regression: the old
 * exporter rasterized whichever single slide the virtualized deck viewport
 * happened to have mounted). */
async function makeMultiSlidePptxBuffer() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Atlas multislide fixture — slide one</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`)
  zip.file('ppt/slides/slide2.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Slide two content</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`)
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function makeOdpBuffer() {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`)
  zip.file('content.xml', `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.2">
  <office:automatic-styles>
    <style:page-layout style:name="pm1">
      <style:page-layout-properties fo:page-width="33.867cm" fo:page-height="19.05cm"/>
    </style:page-layout>
  </office:automatic-styles>
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" draw:style-name="dp1" draw:master-page-name="Default">
        <draw:frame svg:x="1cm" svg:y="1cm" svg:width="20cm" svg:height="3cm">
          <draw:text-box>
            <text:p>Atlas ODP fixture</text:p>
          </draw:text-box>
        </draw:frame>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>`)
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function makeOdtBuffer() {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`)
  zip.file('styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">
  <office:styles />
</office:document-styles>`)
  zip.file('content.xml', `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.2">
  <office:body>
    <office:text>
      <text:p>Atlas ODT fixture</text:p>
      <text:p>Smoke test document.</text:p>
    </office:text>
  </office:body>
</office:document-content>`)
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function writeFixture(name, content) {
  const targetPath = path.join(outputDir, name)
  await fs.writeFile(targetPath, content)
  return targetPath
}

export async function generateFixtures() {
  await fs.mkdir(outputDir, { recursive: true })

  const docxBuffer = await makeDocxBuffer()
  const pptxBuffer = await makePptxBuffer()
  const odpBuffer = await makeOdpBuffer()
  const odtBuffer = await makeOdtBuffer()
  const pdfBuffer = makePdfBuffer()
  const multiPagePdfBuffer = makeMultiPagePdfBuffer()
  const multiPageDocxBuffer = await makeMultiPageDocxBuffer()
  const multiSlidePptxBuffer = await makeMultiSlidePptxBuffer()
  const multiSheetXlsxBuffer = makeMultiSheetWorkbookBuffer('xlsx')

  const writes = [
    writeFixture(
      'sample.md',
      '# Atlas Markdown fixture\n\nSmoke test paragraph.\n\n' +
        '```mermaid\nflowchart TD\n  A[Start] --> B[Atlas]\n```\n',
    ),
    writeFixture('sample.txt', 'Atlas text fixture\nSecond line\n'),
    writeFixture('sample.ts', 'export function atlasFixture(): string {\n  return \'ok\'\n}\n'),
    writeFixture('sample.csv', 'name,value\nAtlas,2\nPhase,1\n'),
    writeFixture('sample.tsv', 'name\tvalue\nAtlas\t2\nPhase\t1\n'),
    writeFixture('sample.xlsx', makeWorkbookBuffer('xlsx')),
    writeFixture('sample.ods', makeWorkbookBuffer('ods')),
    writeFixture('sample.docx', docxBuffer),
    writeFixture('sample.rtf', '{\\rtf1\\ansi Atlas RTF fixture\\par Smoke test document.\\par}'),
    writeFixture('sample.odt', odtBuffer),
    writeFixture('sample.pdf', pdfBuffer),
    writeFixture('sample-multipage.pdf', multiPagePdfBuffer),
    writeFixture('sample.pptx', pptxBuffer),
    writeFixture('sample.odp', odpBuffer),
    writeFixture('sample-multipage.docx', multiPageDocxBuffer),
    writeFixture('sample-multislide.pptx', multiSlidePptxBuffer),
    writeFixture('sample-multisheet.xlsx', multiSheetXlsxBuffer),
  ]

  return Promise.all(writes)
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const files = await generateFixtures()
  for (const file of files) {
    console.log(path.relative(projectRoot, file).replace(/\\/g, '/'))
  }
}
