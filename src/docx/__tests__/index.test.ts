/**
 * Integration tests for `loadDocx`/`saveDocx` (src/docx/index.ts).
 *
 * Builds a real, realistic `.docx` fixture in-test with the `docx` npm
 * package (header, footer, footnote, endnote, and a fixed-width table) and
 * round-trips it through Atlas's own parser/serializer pipeline, covering:
 *   - P1.3 / DXP-03 / DXP-04 / DXS-01 — header/footer/footnote/endnote
 *     parts parse into real Paragraph[] blocks and no longer crash the save.
 *   - P1.5 / DXP-19 / DXS-02 — w:tblGrid survives parse → layout → save.
 *   - P1.7 / DXS-06 — a first comment added to a comment-less document
 *     gets its relationship + content-type registered.
 *   - D20 / DXS-19 — a corrupted in-memory document fails the save with a
 *     clear error instead of producing a broken package.
 */
import {
  Document as DocxJsDocument,
  EndnoteReferenceRun,
  Footer as DocxJsFooter,
  FootnoteReferenceRun,
  Header as DocxJsHeader,
  Packer,
  Paragraph as DocxJsParagraph,
  Table as DocxJsTable,
  TableCell as DocxJsTableCell,
  TableRow as DocxJsTableRow,
  TextRun,
  WidthType,
} from 'docx'
import JSZip from 'jszip'
import { beforeAll, describe, expect, it } from 'vitest'

import { loadDocx, saveDocx, type DocxBundle } from '..'
import { addCommentToDocument } from '../editor/commentMutations'

// `Packer.toBuffer` returns a Node `Buffer`, and `saveDocx` returns whatever
// realm's `Uint8Array` JSZip constructed it with — under Vitest's jsdom
// environment either can carry an `ArrayBuffer` from a different realm than
// this test file's globals, which JSZip's own `instanceof ArrayBuffer` type
// check then silently refuses to load. Copying through the *local*
// `Uint8Array` constructor guarantees a same-realm `ArrayBuffer`.
function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return new Uint8Array(bytes).buffer
}

async function buildFixtureBuffer(): Promise<ArrayBuffer> {
  const doc = new DocxJsDocument({
    sections: [
      {
        headers: {
          default: new DocxJsHeader({
            children: [new DocxJsParagraph({ children: [new TextRun('Header text')] })],
          }),
        },
        footers: {
          default: new DocxJsFooter({
            children: [new DocxJsParagraph({ children: [new TextRun('Footer text')] })],
          }),
        },
        children: [
          new DocxJsParagraph({
            children: [
              new TextRun('Body text with a footnote'),
              new FootnoteReferenceRun(1),
              new EndnoteReferenceRun(1),
            ],
          }),
          new DocxJsTable({
            width: { size: 5000, type: WidthType.DXA },
            columnWidths: [2400, 2600],
            rows: [
              new DocxJsTableRow({
                children: [
                  new DocxJsTableCell({ children: [new DocxJsParagraph('A1')] }),
                  new DocxJsTableCell({ children: [new DocxJsParagraph('A2')] }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
    footnotes: {
      1: { children: [new DocxJsParagraph({ children: [new TextRun('Footnote body text.')] })] },
    },
    endnotes: {
      1: { children: [new DocxJsParagraph({ children: [new TextRun('Endnote body text.')] })] },
    },
  })

  const buffer = await Packer.toBuffer(doc)
  return toArrayBuffer(buffer)
}

function findTable(document: DocxBundle['document']) {
  return document.sections.flatMap((section) => section.blocks).find((block) => block.kind === 'table')
}

describe('loadDocx / saveDocx integration', () => {
  let fixtureBuffer: ArrayBuffer

  beforeAll(async () => {
    fixtureBuffer = await buildFixtureBuffer()
  })

  it('parses header/footer/footnote/endnote bodies into real paragraph blocks and the table grid', async () => {
    const bundle = await loadDocx(fixtureBuffer)

    expect(bundle.document.headers.size).toBe(1)
    const header = [...bundle.document.headers.values()][0]
    expect(header?.blocks.some((block) => block.kind === 'paragraph')).toBe(true)

    expect(bundle.document.footers.size).toBe(1)
    const footer = [...bundle.document.footers.values()][0]
    expect(footer?.blocks.some((block) => block.kind === 'paragraph')).toBe(true)

    const footnote = [...bundle.document.footnotes.values()].find((note) => note.noteType === undefined)
    expect(footnote?.blocks[0]?.kind).toBe('paragraph')

    const endnote = [...bundle.document.endnotes.values()].find((note) => note.noteType === undefined)
    expect(endnote?.blocks[0]?.kind).toBe('paragraph')

    const table = findTable(bundle.document)
    expect(table?.kind).toBe('table')
    if (table?.kind === 'table') {
      expect(table.tblGrid).toEqual([2400, 2600])
    }
  })

  it('round-trips the full document without throwing (DXS-01 / DXS-02)', async () => {
    const bundle = await loadDocx(fixtureBuffer)

    const saved = await saveDocx(bundle)
    expect(saved.byteLength).toBeGreaterThan(0)

    const reloaded = await loadDocx(toArrayBuffer(saved))
    expect(reloaded.document.headers.size).toBe(1)
    expect(reloaded.document.footers.size).toBe(1)
    expect(reloaded.document.footnotes.size).toBe(bundle.document.footnotes.size)
    expect(reloaded.document.endnotes.size).toBe(bundle.document.endnotes.size)

    const table = findTable(reloaded.document)
    if (table?.kind === 'table') {
      expect(table.tblGrid).toEqual([2400, 2600])
    }
  })

  it('registers a relationship + content-type for the first comment added to a comment-less document (DXS-06)', async () => {
    const bundle = await loadDocx(fixtureBuffer)
    expect(bundle.document.comments.size).toBe(0)

    // The `docx` package always writes an (empty) comments part. Strip it so
    // this genuinely exercises "a document that never had one" — the exact
    // scenario DXS-06 was about.
    const rawArchive = new Map(bundle.rawArchive)
    rawArchive.delete('word/comments.xml')
    const relationships = (bundle.relationships ?? []).filter((rel) => !rel.type.endsWith('/comments'))
    const contentTypes = bundle.contentTypes
      ? {
          ...bundle.contentTypes,
          overrides: bundle.contentTypes.overrides.filter(
            (override) => override.partName !== '/word/comments.xml',
          ),
        }
      : undefined
    const strippedBundle: DocxBundle = { ...bundle, rawArchive, relationships, contentTypes }

    const { document: withComment } = addCommentToDocument(
      strippedBundle.document,
      {
        anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 },
        focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 4 },
      },
      'A new remark',
      'Test Author',
    )
    const mutatedBundle: DocxBundle = { ...strippedBundle, document: withComment }

    const saved = await saveDocx(mutatedBundle)
    const zip = await JSZip.loadAsync(saved)

    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string')
    expect(relsXml).toContain('relationships/comments')
    expect(relsXml).toContain('comments.xml')

    const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string')
    expect(contentTypesXml).toContain('/word/comments.xml')
    expect(contentTypesXml).toContain('wordprocessingml.comments+xml')

    const reloaded = await loadDocx(toArrayBuffer(saved))
    expect(reloaded.document.comments.size).toBe(1)
  })

  it('persists a resolved comment across save/reload via commentsExtended.xml (D16 / DXS-11)', async () => {
    const bundle = await loadDocx(fixtureBuffer)
    const { document: withComment, commentId } = addCommentToDocument(
      bundle.document,
      {
        anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 },
        focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 4 },
      },
      'Please fix this',
      'Reviewer',
    )
    const resolvedComments = new Map(withComment.comments)
    const comment = resolvedComments.get(commentId)
    if (!comment) {
      throw new Error('expected the newly added comment to exist')
    }
    resolvedComments.set(commentId, { ...comment, resolved: true })
    const resolvedBundle: DocxBundle = {
      ...bundle,
      document: { ...withComment, comments: resolvedComments },
    }

    const saved = await saveDocx(resolvedBundle)
    const zip = await JSZip.loadAsync(saved)

    const commentsExtendedXml = await zip.file('word/commentsExtended.xml')?.async('string')
    expect(commentsExtendedXml).toContain('w15:done="1"')

    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string')
    expect(relsXml).toContain('commentsExtended')

    const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string')
    expect(contentTypesXml).toContain('commentsExtended')

    const reloaded = await loadDocx(toArrayBuffer(saved))
    const reloadedComment = [...reloaded.document.comments.values()][0]
    expect(reloadedComment?.resolved).toBe(true)
  })

  it('removes stale comments.xml/commentsExtended.xml once every comment is deleted (wave 1 follow-up)', async () => {
    const bundle = await loadDocx(fixtureBuffer)
    const { document: withComment, commentId } = addCommentToDocument(
      bundle.document,
      {
        anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 },
        focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 4 },
      },
      'Temporary remark',
      'Reviewer',
    )
    const resolvedComments = new Map(withComment.comments)
    const comment = resolvedComments.get(commentId)
    if (!comment) {
      throw new Error('expected the newly added comment to exist')
    }
    resolvedComments.set(commentId, { ...comment, resolved: true })
    const withResolvedComment: DocxBundle = {
      ...bundle,
      document: { ...withComment, comments: resolvedComments },
    }

    // First save: comments.xml + commentsExtended.xml both get created and
    // registered — this reproduces the "stale archive copy" scenario by
    // giving the *next* save a rawArchive that already contains both parts.
    const firstSave = await saveDocx(withResolvedComment)
    const afterFirstSave = await loadDocx(toArrayBuffer(firstSave))
    expect(afterFirstSave.document.comments.size).toBe(1)

    // Second save: the user deleted the only comment. Nothing in the model
    // references it any more.
    const emptiedBundle: DocxBundle = {
      ...afterFirstSave,
      document: { ...afterFirstSave.document, comments: new Map() },
    }
    const secondSave = await saveDocx(emptiedBundle)
    const zip = await JSZip.loadAsync(secondSave)

    expect(zip.file('word/comments.xml')).toBeNull()
    expect(zip.file('word/commentsExtended.xml')).toBeNull()

    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string')
    expect(relsXml ?? '').not.toContain('/comments')
    expect(relsXml ?? '').not.toContain('commentsExtended')

    const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string')
    expect(contentTypesXml ?? '').not.toContain('comments.xml')
    expect(contentTypesXml ?? '').not.toContain('commentsExtended.xml')

    const reloaded = await loadDocx(toArrayBuffer(secondSave))
    expect(reloaded.document.comments.size).toBe(0)
  })

  it('fails the save with a clear error instead of writing a broken package (D20)', async () => {
    const bundle = await loadDocx(fixtureBuffer)
    const table = findTable(bundle.document)
    if (table?.kind !== 'table') {
      throw new Error('expected the fixture to contain a table')
    }

    // Simulate a future regression that drops tblGrid again (DXS-02).
    const corruptedTable = { ...table, tblGrid: undefined }
    const corruptedSections = bundle.document.sections.map((section) => ({
      ...section,
      blocks: section.blocks.map((block) => (block === table ? corruptedTable : block)),
    }))
    const corruptedBundle: DocxBundle = {
      ...bundle,
      document: { ...bundle.document, sections: corruptedSections },
    }

    await expect(saveDocx(corruptedBundle)).rejects.toThrow(/tblGrid/)
  })
})
