// @vitest-environment node
/**
 * Real text editing, end to end on files from real producers: find a
 * paragraph, change it, write the file, then read the result back with pdf.js
 * (what Atlas and most viewers render with) and with this module's own
 * interpreter, and check what is — and is no longer — in the file.
 */
import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import * as pdfLib from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'
import { getDocument, GlobalWorkerOptions, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'

import { createInitialEditState } from '../../editState'
import { applyEditsToPdf } from '../../writePdf'
import { UnsupportedCharacterError } from '../../errors'
import type { PdfEditState, TextEdit } from '../../types'
import { formsSharedAcrossPages, interpretPage } from '../interpreter'
import { buildParagraphs, type StyledSpan, type TextParagraph } from '../paragraphs'
import { editParagraph, newTextEdit, spansText } from '../textEdits'
import { fallbackForTests } from './fallbackForTests'

const ROOT = path.resolve(__dirname, '..', '..', '..', '..', '..', '..')
GlobalWorkerOptions.workerSrc = pathToFileURL(path.join(ROOT, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs')).href
const STANDARD_FONTS = pathToFileURL(path.join(ROOT, 'node_modules', 'pdfjs-dist', 'standard_fonts')).href + '/'

const LIBREOFFICE = new Uint8Array(fs.readFileSync(path.join(__dirname, '..', '..', '__fixtures__', 'libreoffice-paragraphs.pdf')))

const tasks: Array<ReturnType<typeof getDocument>> = []
afterEach(async () => {
  await Promise.all(tasks.splice(0).map((task) => task.destroy()))
})

type Item = { str: string; x: number; y: number; fontName: string }

async function readBack(bytes: Uint8Array): Promise<Item[]> {
  const task = getDocument({ data: bytes.slice(0), standardFontDataUrl: STANDARD_FONTS })
  tasks.push(task)
  const pdf = await task.promise
  const content = await (await pdf.getPage(1)).getTextContent()
  return content.items
    .filter((item): item is typeof item & { str: string; transform: number[]; fontName: string } => 'str' in item)
    .map((item) => ({ str: item.str, x: item.transform[4], y: item.transform[5], fontName: item.fontName }))
}

async function paragraphsOf(bytes: Uint8Array): Promise<TextParagraph[]> {
  const doc = await pdfLib.PDFDocument.load(bytes)
  const page = doc.getPage(0)
  const text = interpretPage(pdfLib, doc.context, page.node, page.node.Resources())
  if (!text) throw new Error('page content could not be read')
  return buildParagraphs(text)
}

async function piecesOf(bytes: Uint8Array) {
  const doc = await pdfLib.PDFDocument.load(bytes)
  const page = doc.getPage(0)
  return interpretPage(pdfLib, doc.context, page.node, page.node.Resources())?.pieces ?? []
}

function stateWith(edits: TextEdit[]): PdfEditState {
  return { ...createInitialEditState(1), textEdits: edits }
}

function find(paragraphs: TextParagraph[], needle: string): TextParagraph {
  const found = paragraphs.find((paragraph) => spansText(paragraph.spans).includes(needle))
  if (!found) throw new Error(`no paragraph contains "${needle}"`)
  return found
}

function replaceIn(spans: ReadonlyArray<StyledSpan>, from: string, to: string): StyledSpan[] {
  return spans.map((span) => ({ ...span, text: span.text.replace(from, to) }))
}

const pageText = (items: Item[]) => items.map((item) => item.str).join('')

describe('reading paragraphs from a LibreOffice PDF', () => {
  it('finds the heading, the justified paragraph with its bold words, and the list items', async () => {
    const paragraphs = await paragraphsOf(LIBREOFFICE)
    const heading = find(paragraphs, 'Quarterly Report')
    expect(heading.lineCount).toBe(1)
    expect(heading.styles[0].font.bold).toBe(true)

    const body = find(paragraphs, 'quick brown fox')
    expect(body.lineCount).toBe(3)
    expect(body.align).toBe('justify')
    expect(body.blocked).toBeNull()
    const bold = body.spans.find((span) => span.text === 'annual budget')
    expect(bold && body.styles[bold.style].font.bold).toBe(true)

    const bullet = find(paragraphs, 'First bullet item')
    expect(spansText(bullet.spans)).toBe('•\tFirst bullet item')
    expect(bullet.tabStops).toHaveLength(1)
  })
})

describe('editing text in a LibreOffice PDF', () => {
  it('removes the old words from the file and writes the new ones, keeping the bold span and the other text in place', async () => {
    const before = await readBack(LIBREOFFICE)
    const paragraphs = await paragraphsOf(LIBREOFFICE)
    const body = find(paragraphs, 'quick brown fox')
    const edit = editParagraph('e1', 1, body, replaceIn(body.spans, 'quick brown fox', 'slow red hen'))

    const result = await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([edit]), 1, fallbackForTests)
    const after = await readBack(result)

    expect(pageText(after)).toContain('slow red hen')
    expect(pageText(after)).not.toContain('quick brown')
    // Not just hidden: no operator in the file shows those glyphs any more.
    expect((await piecesOf(result)).map((piece) => piece.text).join('')).not.toContain('quick')

    // Untouched text is exactly where it was.
    for (const needle of ['Quarterly', 'A short second paragraph.', 'Second bullet item']) {
      const was = before.find((item) => item.str.includes(needle.slice(0, 6)))
      const now = after.find((item) => item.str.includes(needle.slice(0, 6)))
      expect(now, needle).toBeDefined()
      expect(now?.x).toBeCloseTo(was?.x ?? NaN, 2)
      expect(now?.y).toBeCloseTo(was?.y ?? NaN, 2)
    }

    // The bold words are still drawn in the document's own bold font.
    const pieces = await piecesOf(result)
    const budget = pieces.find((piece) => piece.text.includes('budget'))
    expect(budget?.state.font?.baseFont).toMatch(/LiberationSerif-Bold$/)
    // And the unchanged words kept the document's own regular font too.
    const hen = pieces.find((piece) => piece.text.includes('hen'))
    expect(hen?.state.font?.baseFont).toMatch(/LiberationSerif$/)
  })

  it('re-wraps a longer paragraph to its own width and adds lines below', async () => {
    const paragraphs = await paragraphsOf(LIBREOFFICE)
    const body = find(paragraphs, 'quick brown fox')
    const longer = replaceIn(body.spans, 'The quick brown fox', 'The quick brown fox, which had been waiting patiently all morning,')
    const result = await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([editParagraph('e', 1, body, longer)]), 1, fallbackForTests)

    const edited = find(await paragraphsOf(result), 'waiting patiently')
    expect(edited.lineCount).toBe(4)
    // Wrapped within the original width (a hair of tolerance for rounding).
    expect(edited.width).toBeLessThanOrEqual(body.width + 1)
    expect(spansText(edited.spans)).toContain('without any major disruption to supply.')
  })

  it('falls back to a bundled font for characters the subset font lacks', async () => {
    const paragraphs = await paragraphsOf(LIBREOFFICE)
    const second = find(paragraphs, 'A short second paragraph.')
    // "Ж" and "ő" are certainly not in this document's subset of Liberation Serif.
    const edit = editParagraph('e', 1, second, [{ text: 'Жарко, ő!', style: 0 }])
    const result = await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([edit]), 1, fallbackForTests)

    expect(pageText(await readBack(result))).toContain('Жарко, ő!')
    const piece = (await piecesOf(result)).find((candidate) => candidate.text.includes('Жарко'))
    expect(piece?.state.font?.baseFont).toMatch(/Tinos/)
  })

  it('refuses characters no available font can draw', async () => {
    const paragraphs = await paragraphsOf(LIBREOFFICE)
    const second = find(paragraphs, 'A short second paragraph.')
    const edit = editParagraph('e', 1, second, [{ text: '漢字', style: 0 }])
    await expect(applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([edit]), 1, fallbackForTests)).rejects.toBeInstanceOf(UnsupportedCharacterError)
  })

  it('keeps a bullet item\'s hanging indent when its text changes', async () => {
    const before = await readBack(LIBREOFFICE)
    const bullet = find(await paragraphsOf(LIBREOFFICE), 'First bullet item')
    const edit = editParagraph('e', 1, bullet, replaceIn(bullet.spans, 'First', 'Opening'))
    const after = await readBack(await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([edit]), 1, fallbackForTests))

    const wasStart = before.find((item) => item.str.startsWith('F') && item.y > 660 && item.y < 680)
    const nowStart = after.find((item) => item.str.startsWith('Open'))
    expect(nowStart?.x).toBeCloseTo(wasStart?.x ?? NaN, 1)
  })

  it('deleting all of a paragraph\'s text removes it and leaves everything else alone', async () => {
    const before = await readBack(LIBREOFFICE)
    const heading = find(await paragraphsOf(LIBREOFFICE), 'Quarterly Report')
    const result = await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([editParagraph('e', 1, heading, [])]), 1, fallbackForTests)
    const after = await readBack(result)
    expect(pageText(after)).not.toContain('Quarterly')
    // Everything below the heading is still there, at the same coordinates.
    const below = (items: Item[]) =>
      items
        .filter((item) => item.str.trim() && item.y < 770)
        .map((item) => `${item.str}@${item.x.toFixed(2)},${item.y.toFixed(2)}`)
        .sort()
    expect(below(after)).toEqual(below(before))
  })

  it('adds new text where there was none', async () => {
    const edit = newTextEdit('n', 1, { x: 72, y: 300 }, { x: 1, y: 0 }, 'Added by Atlas', { family: 'helvetica', bold: false, italic: false }, 14, { r: 0.8, g: 0, b: 0 })
    const after = await readBack(await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([edit]), 1, fallbackForTests))
    const added = after.find((item) => item.str.includes('Added by Atlas'))
    expect(added?.x).toBeCloseTo(72, 1)
    expect(added?.y).toBeCloseTo(300, 1)
  })

  it('applies edits to two paragraphs in the same stream at once', async () => {
    const paragraphs = await paragraphsOf(LIBREOFFICE)
    const heading = find(paragraphs, 'Quarterly Report')
    const second = find(paragraphs, 'A short second paragraph.')
    const edits = [
      editParagraph('a', 1, heading, replaceIn(heading.spans, 'Quarterly', 'Annual')),
      editParagraph('b', 1, second, replaceIn(second.spans, 'short', 'brief')),
    ]
    const text = pageText(await readBack(await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith(edits), 1, fallbackForTests)))
    expect(text).toContain('Annual Report')
    expect(text).toContain('A brief second paragraph.')
    expect(text).not.toContain('Quarterly')
  })
})

describe('editing text in a Chromium (Skia) PDF', () => {
  const CHROMIUM = new Uint8Array(fs.readFileSync(path.join(__dirname, '..', '..', '__fixtures__', 'chromium-print.pdf')))

  it('reads its paragraphs: a serif heading, a justified paragraph with an italic span, a centred line', async () => {
    const paragraphs = await paragraphsOf(CHROMIUM)
    const heading = find(paragraphs, 'Invoice Summary')
    expect(heading.blocked).toBeNull()
    expect(heading.styles[0].font.family).toBe('times')

    const body = find(paragraphs, 'Payment is due')
    expect(body.lineCount).toBeGreaterThan(1)
    expect(body.align).toBe('justify')
    const italic = body.spans.find((span) => span.text.includes('two percent'))
    expect(italic && body.styles[italic.style].font.italic).toBe(true)
  })

  it('rewrites a paragraph in its composite fonts and keeps the rest of the page in place', async () => {
    const before = await readBack(CHROMIUM)
    const body = find(await paragraphsOf(CHROMIUM), 'Payment is due')
    const edit = editParagraph('e', 1, body, replaceIn(body.spans, 'thirty days', 'sixty days'))
    const result = await applyEditsToPdf(pdfLib, CHROMIUM, stateWith([edit]), 1, fallbackForTests)
    const after = await readBack(result)

    expect(pageText(after)).toContain('sixty days')
    expect(pageText(after)).not.toContain('thirty')
    const was = before.find((item) => item.str.includes('Thank you'))
    const now = after.find((item) => item.str.includes('Thank you'))
    expect(now?.x).toBeCloseTo(was?.x ?? NaN, 2)
    expect(now?.y).toBeCloseTo(was?.y ?? NaN, 2)
    // "sixty" only needs letters the original text used, so no fallback font is involved.
    const pieces = await piecesOf(result)
    expect(pieces.some((piece) => /^[A-Z]{6}\+Atlas|Arimo|Tinos/.test(piece.state.font?.baseFont ?? '') && piece.text.includes('sixty'))).toBe(false)
  })
})

/** Every glyph pdf.js will draw on page 1, with `!` after any its font doesn't contain. */
async function drawnGlyphs(bytes: Uint8Array): Promise<string> {
  const task = getDocument({ data: bytes.slice(0), standardFontDataUrl: STANDARD_FONTS })
  tasks.push(task)
  const page = await (await task.promise).getPage(1)
  const ops = await page.getOperatorList()
  let out = ''
  ops.fnArray.forEach((fn, index) => {
    if (fn !== OPS.showText) return
    for (const glyph of ops.argsArray[index][0] as Array<{ unicode?: string; isInFont?: boolean } | number>) {
      if (typeof glyph === 'number' || !glyph.unicode?.trim()) continue
      out += `${glyph.unicode}${glyph.isInFont ? '' : '!'}`
    }
  })
  return out
}

describe('what a reader actually draws', () => {
  it('reuses the subset font only for letters it has, and falls back for the rest — no missing glyphs either way', async () => {
    const paragraphs = await paragraphsOf(LIBREOFFICE)
    const heading = find(paragraphs, 'Quarterly Report')
    const second = find(paragraphs, 'A short second paragraph.')
    const edits = [
      // Letters the bold sans subset has, and "Zebra" (Z and b it doesn't).
      editParagraph('a', 1, heading, [{ text: 'Quarterly Zebra', style: 0 }]),
      editParagraph('b', 1, second, [{ text: 'A brief second paragraph.', style: 0 }]),
    ]
    const result = await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith(edits), 1, fallbackForTests)
    const drawn = await drawnGlyphs(result)
    expect(drawn).toContain('QuarterlyZebra')
    expect(drawn).toContain('Abriefsecondparagraph.')
    expect(drawn).not.toContain('!')
  })

  it('keeps a bullet item in the document\'s own font: its tab gap is positioning, not a character to encode', async () => {
    const bullet = find(await paragraphsOf(LIBREOFFICE), 'First bullet item')
    const edit = editParagraph('e', 1, bullet, replaceIn(bullet.spans, 'First', 'Fist'))
    const pieces = await piecesOf(await applyEditsToPdf(pdfLib, LIBREOFFICE, stateWith([edit]), 1, fallbackForTests))
    const item = pieces.find((piece) => piece.text.includes('Fist'))
    expect(item?.state.font?.baseFont).toMatch(/LiberationSerif$/)
  })
})

describe('form XObjects', () => {
  /**
   * Pages drawing one form XObject that has its OWN resources (its own font
   * dictionary), built by hand so nothing shares a resources dictionary by
   * accident. Page text and form text both use Helvetica.
   */
  async function formPdf(options: { readonly pages: number; readonly drawsPerPage: number }): Promise<Uint8Array> {
    const doc = await pdfLib.PDFDocument.create()
    const { context } = doc
    const helvetica = () => context.obj({ Type: 'Font', Subtype: 'Type1', BaseFont: 'Helvetica', Encoding: 'WinAnsiEncoding' })
    const form = context.register(
      context.flateStream('BT /FH 12 Tf 10 50 Td (Header text) Tj ET', {
        Type: 'XObject',
        Subtype: 'Form',
        BBox: [0, 0, 300, 100],
        Resources: { Font: { FH: context.register(helvetica()) } },
      }),
    )
    for (let page = 0; page < options.pages; page += 1) {
      const target = doc.addPage([612, 792])
      const draws = Array.from({ length: options.drawsPerPage }, (_, draw) => `q 1 0 0 1 50 ${650 - draw * 120} cm /X1 Do Q`)
      target.node.set(
        pdfLib.PDFName.of('Resources'),
        context.obj({ Font: { FP: context.register(helvetica()) }, XObject: { X1: form } }),
      )
      target.node.set(
        pdfLib.PDFName.of('Contents'),
        context.register(context.flateStream(`${draws.join(' ')} BT /FP 12 Tf 72 300 Td (Own page text) Tj ET`)),
      )
    }
    return doc.save()
  }

  it('blocks text in a form drawn on more than one page, or twice on one page — editing it would change every copy', async () => {
    for (const layout of [{ pages: 2, drawsPerPage: 1 }, { pages: 1, drawsPerPage: 2 }]) {
      const bytes = await formPdf(layout)
      const doc = await pdfLib.PDFDocument.load(bytes)
      const shared = formsSharedAcrossPages(pdfLib, doc.context, doc.getPages().map((page) => page.node.Resources()))
      const page = doc.getPage(0)
      const text = interpretPage(pdfLib, doc.context, page.node, page.node.Resources())
      const paragraphs = buildParagraphs(text!, shared)
      expect(find(paragraphs, 'Header text').blocked, JSON.stringify(layout)).toBe('shared')
      expect(find(paragraphs, 'Own page text').blocked).toBeNull()
    }
  })

  it('edits text in a form drawn once, and text on the page, both in fallback fonts, without mixing up their resources', async () => {
    const bytes = await formPdf({ pages: 1, drawsPerPage: 1 })
    const paragraphs = await paragraphsOf(bytes)
    const header = find(paragraphs, 'Header text')
    const own = find(paragraphs, 'Own page text')
    expect(header.streamKey).not.toBe('page')
    const edits = [
      editParagraph('a', 1, header, [{ text: 'Заголовок', style: 0 }]),
      editParagraph('b', 1, own, [{ text: 'Свой текст', style: 0 }]),
    ]
    const result = await applyEditsToPdf(pdfLib, bytes, stateWith(edits), 1, fallbackForTests)
    const drawn = await drawnGlyphs(result)
    expect(drawn).toContain('Заголовок')
    expect(drawn).toContain('Свойтекст')
    expect(drawn).not.toContain('!')
    // Each stream names a font its OWN resources define. (A reader given a
    // name its resources lack substitutes a font silently, so only this shows it.)
    const pieces = await piecesOf(result)
    for (const word of ['Заголовок', 'Свой']) {
      const piece = pieces.find((candidate) => candidate.text.includes(word))
      expect(piece?.state.font?.baseFont, word).toMatch(/Arimo/)
    }
  })
})

describe('pages that leave graphics states open', () => {
  it('places new text by the default coordinate system even after an unmatched q and cm', async () => {
    const doc = await pdfLib.PDFDocument.create()
    const page = doc.addPage([612, 792])
    // A scale, then a `q` never closed: one `Q` after the page restores the
    // scaled state, so anything appended after just one is drawn at double size.
    page.node.set(pdfLib.PDFName.of('Contents'), doc.context.register(doc.context.flateStream('2 0 0 2 0 0 cm q 0 0 10 10 re f')))
    const bytes = await doc.save()
    const edit = newTextEdit('n', 1, { x: 100, y: 200 }, { x: 1, y: 0 }, 'Placed', { family: 'helvetica', bold: false, italic: false }, 12, { r: 0, g: 0, b: 0 })
    const items = await readBack(await applyEditsToPdf(pdfLib, bytes, stateWith([edit]), 1, fallbackForTests))
    const placed = items.find((item) => item.str.includes('Placed'))
    expect(placed?.x).toBeCloseTo(100, 1)
    expect(placed?.y).toBeCloseTo(200, 1)
  })
})

describe('editing text in PDFs pdf-lib writes', () => {
  async function standardFontPdf(): Promise<Uint8Array> {
    const doc = await pdfLib.PDFDocument.create()
    const font = await doc.embedFont(pdfLib.StandardFonts.TimesRoman)
    const page = doc.addPage([612, 792])
    page.drawText('Simple font, WinAnsi, no ToUnicode map.', { x: 72, y: 700, size: 12, font })
    return doc.save()
  }

  async function embeddedFontPdf(): Promise<Uint8Array> {
    const doc = await pdfLib.PDFDocument.create()
    doc.registerFontkit(fontkit)
    const font = await doc.embedFont(await fallbackForTests.load({ family: 'times', bold: false, italic: false }), { subset: true })
    const page = doc.addPage([612, 792])
    page.drawText('Composite subset font with a ToUnicode map.', { x: 72, y: 700, size: 12, font })
    return doc.save()
  }

  it('edits a standard-font (simple, unembedded) run in its own font', async () => {
    const bytes = await standardFontPdf()
    const paragraph = find(await paragraphsOf(bytes), 'Simple font')
    const edit = editParagraph('e', 1, paragraph, replaceIn(paragraph.spans, 'Simple', 'Plain'))
    const result = await applyEditsToPdf(pdfLib, bytes, stateWith([edit]), 1, fallbackForTests)
    expect(pageText(await readBack(result))).toContain('Plain font, WinAnsi')
    const piece = (await piecesOf(result)).find((candidate) => candidate.text.includes('Plain'))
    expect(piece?.state.font?.baseFont).toBe('Times-Roman')
  })

  it('edits a composite (Identity-H) subset run, reusing its glyphs where it has them', async () => {
    const bytes = await embeddedFontPdf()
    const paragraph = find(await paragraphsOf(bytes), 'Composite subset')
    // Every letter of "comet" appears in the original text, so the subset can draw it.
    const edit = editParagraph('e', 1, paragraph, replaceIn(paragraph.spans, 'Composite', 'comet'))
    const result = await applyEditsToPdf(pdfLib, bytes, stateWith([edit]), 1, fallbackForTests)
    expect(pageText(await readBack(result))).toContain('comet subset font')
    const piece = (await piecesOf(result)).find((candidate) => candidate.text.includes('comet'))
    expect(piece?.state.font?.composite).toBe(true)
    expect(piece?.state.font?.baseFont).not.toMatch(/^AtlasF/)
  })
})
