/**
 * D29 follow-up 2 — proves the mixed-paragraph header/footer text edit (see
 * `src/docx/editor/headerFooter.ts`) round-trips through Atlas's REAL
 * `loadDocx`/`saveDocx` save path into a package that still passes
 * `scripts/validate-office-file.mjs`'s spec-level OPC/OOXML checks, not just
 * that the in-memory model/XML-string checks in `headerFooter.test.ts` and
 * `headerWriter.test.ts` look right in isolation.
 *
 * The fixture is a real `.docx` built with the `docx` npm package (so its
 * content-types/relationships/media scaffolding for an embedded image is
 * genuinely valid — the same approach `src/docx/__tests__/index.test.ts`
 * uses), with its header part's OWN paragraph then hand-replaced by a
 * paragraph mixing a logo drawing, plain text, a tab, and a genuine complex
 * (`w:fldChar`/`w:instrText` SIBLING runs, not `w:fldSimple`) PAGE field —
 * the shape a real Word document produces and the one this feature exists
 * for. The image's own relationship id is read back out of the
 * `docx`-generated package rather than hardcoded, so the swapped-in header
 * still references a real, resolvable image relationship.
 */
import { Document as DocxJsDocument, Header as DocxJsHeader, ImageRun, Packer, Paragraph as DocxJsParagraph, TextRun } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { loadDocx, saveDocx } from '../index'
import { applyCommand } from '../editor/commands'
import { buildHeaderFooterSegmentEdit, listHeaderFooterParts } from '../editor/headerFooter'
import { validateOfficeFile } from '../../../scripts/lib/officeValidator.mjs'

// See `src/docx/__tests__/index.test.ts`'s own `toArrayBuffer` doc comment —
// copying through the *local* `Uint8Array` constructor guarantees a
// same-realm `ArrayBuffer`, which JSZip's `instanceof ArrayBuffer` check
// (used by both JSZip itself and Atlas's own `loadDocx`) otherwise silently
// refuses to load under Vitest's jsdom environment.
function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return new Uint8Array(bytes).buffer
}

function errorsOnly(issues: ReadonlyArray<{ readonly severity: string }>) {
  return issues.filter((issue) => issue.severity === 'error')
}

const TINY_PNG = Buffer.from(
  '89504e470d0a1a0a0000000d49484452000000010000000108020000009077'
    + '53de0000000c4944415478da6360000002000155a2645b0000000049454e44ae426082',
  'hex',
)

/**
 * Builds a real `.docx` (valid content-types/media/relationships courtesy of
 * the `docx` package), then swaps its header's paragraph for one this test
 * controls precisely: `[logo drawing][[[Chapter Title]][tab]][complex PAGE
 * field]`, reusing the SAME image relationship id `docx` already wired up.
 */
async function buildFixtureBuffer(): Promise<ArrayBuffer> {
  const doc = new DocxJsDocument({
    sections: [
      {
        headers: {
          default: new DocxJsHeader({
            children: [
              new DocxJsParagraph({
                children: [new ImageRun({ data: TINY_PNG, type: 'png', transformation: { width: 20, height: 20 } }), new TextRun('placeholder')],
              }),
            ],
          }),
        },
        children: [new DocxJsParagraph({ children: [new TextRun('Body text')] })],
      },
    ],
  })

  const packed = await Packer.toBuffer(doc)
  const zip = await JSZip.loadAsync(packed)

  const headerPath = Object.keys(zip.files).find((path) => /word\/header\d*\.xml$/.test(path))
  if (headerPath === undefined) throw new Error('fixture build: no header part in the docx package')
  const originalHeaderXml = await zip.file(headerPath)!.async('string')

  // Reuse the exact drawing markup (and so the exact, already-valid image
  // relationship id) `docx` produced — this test only needs to control the
  // paragraph's TEXT/tab/field content around it.
  const drawingMatch = /<w:drawing>[\s\S]*?<\/w:drawing>/.exec(originalHeaderXml)
  const rootOpenMatch = /^<\?xml[^>]*\?><w:hdr[^>]*>/.exec(originalHeaderXml)
  if (drawingMatch === null || rootOpenMatch === null) {
    throw new Error('fixture build: could not locate the generated drawing or w:hdr root element')
  }

  const mixedParagraphXml =
    '<w:p>'
    + `<w:r>${drawingMatch[0]}</w:r>`
    + '<w:r><w:t xml:space="preserve">Chapter Title</w:t></w:r>'
    + '<w:r><w:tab/></w:r>'
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t>1</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    + '</w:p>'

  zip.file(headerPath, `${rootOpenMatch[0]}${mixedParagraphXml}</w:hdr>`)

  const rebuilt: Uint8Array = await zip.generateAsync({ type: 'uint8array' })
  return toArrayBuffer(rebuilt)
}

describe('a mixed header paragraph (logo + text + tab + complex PAGE field) — full package save', () => {
  it('edits the text, saves through the real loadDocx/saveDocx path, and still passes spec-level OPC/OOXML validation', async () => {
    const fixtureBuffer = await buildFixtureBuffer()
    const bundle = await loadDocx(fixtureBuffer)

    expect(bundle.document.headers.size).toBe(1)
    const [headerId] = [...bundle.document.headers.keys()]

    // Confirm this parsed as a genuine 'mixed' row before editing it.
    const part = listHeaderFooterParts(bundle.document).find((entry) => entry.kind === 'header')!
    expect(part.rows).toEqual([
      {
        kind: 'mixed',
        blockIndex: 0,
        segments: [
          { kind: 'atom', label: '[Image]' },
          { kind: 'text', segmentIndex: 0, text: 'Chapter Title\t' },
          { kind: 'atom', label: '[Page number]' },
        ],
      },
    ])

    const command = buildHeaderFooterSegmentEdit(bundle.document, 'header', headerId, 0, 0, 'Executive Summary\t')!
    const edited = applyCommand(bundle.document, command)

    const saved = await saveDocx({ ...bundle, document: edited.document })

    const result = validateOfficeFile(Buffer.from(saved))
    expect(result.format).toEqual({ family: 'opc', kind: 'docx' })
    expect(errorsOnly(result.issues)).toEqual([])

    // The saved PACKAGE itself (not just the in-memory model) has the new
    // text, and the drawing/field/tab intact.
    const savedZip = await JSZip.loadAsync(toArrayBuffer(saved))
    const savedHeaderPath = Object.keys(savedZip.files).find((path) => /word\/header\d*\.xml$/.test(path))!
    const savedHeaderXml = await savedZip.file(savedHeaderPath)!.async('string')

    expect(savedHeaderXml).toContain('Executive Summary')
    expect(savedHeaderXml).not.toContain('Chapter Title')
    expect(savedHeaderXml).toContain('<w:drawing>')
    expect(savedHeaderXml).toContain('<w:tab/>')
    expect(savedHeaderXml).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
    expect(savedHeaderXml).toContain('w:fldCharType="begin"')
    expect(savedHeaderXml).toContain('w:fldCharType="separate"')
    expect(savedHeaderXml).toContain('w:fldCharType="end"')

    // And round-tripping the saved package back through loadDocx reads the
    // same edited text back (a real save/reopen, not just a serializer
    // string check).
    const reloaded = await loadDocx(toArrayBuffer(saved))
    const reloadedRows = listHeaderFooterParts(reloaded.document).find((entry) => entry.kind === 'header')!.rows
    expect(reloadedRows).toEqual([
      {
        kind: 'mixed',
        blockIndex: 0,
        segments: [
          { kind: 'atom', label: '[Image]' },
          { kind: 'text', segmentIndex: 0, text: 'Executive Summary\t' },
          { kind: 'atom', label: '[Page number]' },
        ],
      },
    ])
  })
})
