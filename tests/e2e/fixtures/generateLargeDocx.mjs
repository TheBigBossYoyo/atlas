// Builds a real, large (many-page, with images) .docx for exercising DocxViewer's
// page virtualization (D23-PERF-2) end-to-end: `Document.load()` and every
// existing docx e2e fixture stay small on purpose, so this is a standalone
// generator invoked on demand rather than added to `generate.mjs`'s normal set.
import { promises as fs } from 'node:fs'
import { pathToFileURL } from 'node:url'
import { Document, Packer, Paragraph, TextRun, PageBreak, ImageRun } from 'docx'

// A tiny (1x1) red PNG, valid image bytes — enough to exercise the real
// drawing/media-relationship path without shipping a real binary asset.
const RED_PIXEL_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  'base64',
)

export async function buildLargeDocx({ pageCount = 36 } = {}) {
  const children = []
  for (let i = 1; i <= pageCount; i++) {
    children.push(
      new Paragraph({
        children: [new TextRun({ text: `Page ${i} heading`, bold: true, size: 32 })],
      }),
    )
    children.push(
      new Paragraph({
        children: [
          new ImageRun({
            data: RED_PIXEL_PNG,
            transformation: { width: 120, height: 120 },
            type: 'png',
          }),
        ],
      }),
    )
    for (let line = 1; line <= 8; line++) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: `Page ${i}, paragraph line ${line} — filler body text for layout.` })],
        }),
      )
    }
    if (i < pageCount) {
      children.push(new Paragraph({ children: [new PageBreak()] }))
    }
  }
  // A final, uniquely-findable marker paragraph at the very end of the document.
  children.push(new Paragraph({ children: [new TextRun({ text: 'END-OF-DOCUMENT-MARKER-XYZZY' })] }))

  const doc = new Document({ sections: [{ children }] })
  return Packer.toBuffer(doc)
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const outPath = process.argv[2] || 'large.docx'
  const pageCount = Number(process.argv[3]) || 36
  const buffer = await buildLargeDocx({ pageCount })
  await fs.writeFile(outPath, buffer)
  console.log('wrote', outPath, buffer.length, 'bytes,', pageCount, 'pages')
}
