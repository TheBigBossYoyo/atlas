/**
 * SHEETFMT-2 — writing the formatting the user applied into the original package.
 *
 * The bar is Excel's own: a format applied in Atlas has to come back as real
 * `<font>`/`<fill>`/`<xf>` entries in `xl/styles.xml` with the cell's `s` index
 * pointing at them, so Excel, LibreOffice and Atlas's own reader all agree. A
 * test that only checked Atlas could read its own output back would pass on a
 * file no other program understands.
 *
 * So these assert on the SAVED XML, and the round-trip test goes back through
 * the reader to check the two halves agree.
 */
import JSZip from 'jszip'
import { beforeEach, describe, expect, it } from 'vitest'

import { attachCellStyles, attachSheetSources, attachTables, cellStyleRequests, parseWorkbookBuffer } from '../../shared/spreadsheetGrid'
import {
  createDocument,
  formatForCell,
  insertRowAt,
  setCellValue,
  setRangeFormat,
  type SpreadsheetDocument,
} from '../spreadsheetDocument'
import { readSheetPartPaths, readSheetTables } from '../spreadsheetTables'
import { readWorkbookCellStyles } from '../xlsxCellStyles'
import { writeWorkbookThroughOriginal } from '../xlsxPassthrough'
import { buildStyledWorkbook } from './styledWorkbook'

async function load(buffer: ArrayBuffer): Promise<SpreadsheetDocument> {
  const [tables, partPaths] = await Promise.all([readSheetTables(buffer), readSheetPartPaths(buffer)])
  const sheets = attachSheetSources(attachTables(parseWorkbookBuffer(buffer), tables), partPaths)
  const styles = await readWorkbookCellStyles(buffer, cellStyleRequests(sheets))
  return createDocument(attachCellStyles(sheets, styles?.styles ?? null))
}

async function saved(original: ArrayBuffer, doc: SpreadsheetDocument) {
  const bytes = await writeWorkbookThroughOriginal(original, doc)
  expect(bytes, 'the passthrough writer refused the document').not.toBeNull()
  const zip = await JSZip.loadAsync(bytes!)
  return {
    bytes: bytes!,
    sheet: await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
    styles: await zip.file('xl/styles.xml')!.async('string'),
  }
}

/**
 * The `<xf>` entries of `<cellXfs>` ONLY, in order, so index n is what a cell's
 * `s="n"` means.
 *
 * `styles.xml` holds two `<xf>` blocks — `<cellStyleXfs>` (named-style
 * definitions) comes first, then `<cellXfs>`. Matching `<xf` across the whole
 * part returns the cellStyleXfs entries first and shifts every index, which
 * reads the wrong style while still looking plausible.
 */
function cellXfs(stylesXml: string): string[] {
  const block = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)
  if (!block) return []
  return block[1].match(/<xf\b[\s\S]*?(?:\/>|<\/xf>)/g) ?? []
}

/** The `<font>` entries of `<fonts>`, in order, so index n is what an xf's `fontId="n"` means. */
function fonts(stylesXml: string): string[] {
  const block = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(stylesXml)
  if (!block) return []
  return block[1].match(/<font>[\s\S]*?<\/font>|<font\s*\/>/g) ?? []
}

/** JSZip hands back a Uint8Array; the loader takes the buffer. */
function asArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

/** The `s` index of one cell in the saved worksheet, or `null` when it carries none. */
function styleOf(sheetXml: string, address: string): string | null {
  const cell = new RegExp(`<c[^>]*\\br="${address}"[^>]*>`).exec(sheetXml)
  if (!cell) return null
  const style = /\bs="(\d+)"/.exec(cell[0])
  return style ? style[1] : null
}

const only = (range: { row: number; col: number }) => ({
  row0: range.row,
  col0: range.col,
  row1: range.row,
  col1: range.col,
})

let original: ArrayBuffer
beforeEach(async () => {
  original = await buildStyledWorkbook()
})

describe('writing applied cell formatting (SHEETFMT-2)', () => {
  it('writes bold as a real font entry and points the cell at it', async () => {
    const doc = await load(original)
    // A2 ("Paper") is plain in the fixture.
    const formatted = setRangeFormat(doc, 0, only({ row: 1, col: 0 }), { bold: true })
    const out = await saved(original, formatted)

    const index = styleOf(out.sheet, 'A2')
    expect(index, 'A2 should now carry a style index').not.toBeNull()

    // That xf must reference a font that is actually bold.
    const fontId = Number(/\bfontId="(\d+)"/.exec(cellXfs(out.styles)[Number(index)])![1])
    expect(fonts(out.styles)[fontId]).toContain('<b')
  })

  it('writes a fill as a solid patternFill with the colour in fgColor', async () => {
    // fgColor, not bgColor: that is where Excel puts the visible colour of a
    // solid fill, and the reader half of this feature depends on it.
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 1, col: 0 }), { fill: '#FF00FF' })
    const out = await saved(original, formatted)

    expect(out.styles).toContain('patternType="solid"')
    expect(out.styles).toMatch(/<fgColor rgb="FFFF00FF"\s*\/>/)
  })

  it('writes an alignment and a number format onto the cell xf', async () => {
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 1, col: 2 }), {
      align: 'right',
      numberFormat: '0.000',
    })
    const out = await saved(original, formatted)

    const index = Number(styleOf(out.sheet, 'C2'))
    expect(cellXfs(out.styles)[index]).toContain('horizontal="right"')
    expect(out.styles).toContain('formatCode="0.000"')
  })

  it('keeps everything about a cell it only re-formatted', async () => {
    // The property that makes this safe to use on a real workbook: formatting a
    // cell the user never TYPED in must not rebuild it from display text. B2 is
    // a date (serial 45000 with a dd/mm/yyyy format) — rebuilt from its
    // rendered text it would come back as the string "15/03/2023".
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 1, col: 1 }), { bold: true })
    const out = await saved(original, formatted)

    const cell = /<c[^>]*\br="B2"[^>]*>[\s\S]*?<\/c>/.exec(out.sheet)![0]
    expect(cell, 'the date should still be a number, not its rendered text').toContain('<v>45000</v>')
    expect(cell).not.toContain('15/03/2023')
  })

  it('keeps a formula cell a formula when only its format changes', async () => {
    const doc = await load(original)
    // C3 is `=C2*2` with a cached 25.
    const formatted = setRangeFormat(doc, 0, only({ row: 2, col: 2 }), { color: '#FF0000' })
    const out = await saved(original, formatted)

    const cell = /<c[^>]*\br="C3"[^>]*>[\s\S]*?<\/c>/.exec(out.sheet)![0]
    expect(cell).toContain('<f>C2*2</f>')
  })

  it('composes with an edit to the same cell', async () => {
    const doc = await load(original)
    const edited = setCellValue(doc, 0, 1, 0, 'Card')
    const formatted = setRangeFormat(edited, 0, only({ row: 1, col: 0 }), { bold: true })
    const out = await saved(original, formatted)

    const cell = /<c[^>]*\br="A2"[^>]*>[\s\S]*?<\/c>/.exec(out.sheet)![0]
    expect(cell).toContain('Card')
    expect(styleOf(out.sheet, 'A2')).not.toBeNull()
  })

  it('reuses one style entry for a whole formatted range instead of one per cell', async () => {
    // Excel caps `cellXfs` at 64k, and a naive implementation appends an entry
    // per cell per save — so formatting a 3-cell row three times would add nine
    // entries, and a real column would exhaust the table.
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, { row0: 1, col0: 0, row1: 2, col1: 0 }, { bold: true })
    const out = await saved(original, formatted)

    const before = (/<cellXfs count="(\d+)"/.exec(
      await (await JSZip.loadAsync(original)).file('xl/styles.xml')!.async('string'),
    ))![1]
    const after = /<cellXfs count="(\d+)"/.exec(out.styles)![1]

    // Both cells share the SAME new entry, so exactly one was added.
    expect(Number(after)).toBe(Number(before) + 1)
    expect(styleOf(out.sheet, 'A2')).toBe(styleOf(out.sheet, 'A3'))
  })

  it('adds nothing at all when the same format is applied twice', async () => {
    const doc = await load(original)
    const once = setRangeFormat(doc, 0, only({ row: 1, col: 0 }), { bold: true })
    const first = await saved(original, once);
    // Saving the identical document again must be idempotent — otherwise every
    // Ctrl+S grows the file.
    const second = await saved(original, once)
    expect(/<cellXfs count="(\d+)"/.exec(second.styles)![1]).toBe(
      /<cellXfs count="(\d+)"/.exec(first.styles)![1],
    )
  })

  it('reuses an existing style when the format already exists in the file', async () => {
    // The fixture's header row is already bold via xf 2. Bolding a plain cell
    // should land on that same entry rather than minting a duplicate.
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 1, col: 0 }), { bold: true })
    const out = await saved(original, formatted)
    expect(Number(/<cellXfs count="(\d+)"/.exec(out.styles)![1])).toBeLessThanOrEqual(6)
  })

  it('can clear a format the file applied', async () => {
    // Un-bolding the fixture's own bold header: `false` has to be written as a
    // font WITHOUT `<b/>`, not merely left alone.
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 0, col: 0 }), { bold: false })
    const out = await saved(original, formatted)

    const index = Number(styleOf(out.sheet, 'A1'))
    const fontId = Number(/\bfontId="(\d+)"/.exec(cellXfs(out.styles)[index])![1])
    expect(fonts(out.styles)[fontId]).not.toContain('<b')
  })

  it("keeps styles.xml's fixed element order, including when it has to create the tables", async () => {
    // CT_Stylesheet fixes the sequence: numFmts, fonts, fills, borders,
    // cellStyleXfs, cellXfs. Excel refuses to open a file that breaks it, while
    // every value-level assertion in this file would still pass — so the order
    // is asserted directly rather than trusted to the structural validator,
    // which may not model this particular sequence.
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 1, col: 0 }), {
      bold: true,
      fill: '#00B0F0',
      numberFormat: '0.00',
    })
    const out = await saved(original, formatted)

    const ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs']
    const seen = (out.styles.match(/<(numFmts|fonts|fills|borders|cellStyleXfs|cellXfs)\b/g) ?? []).map((m) =>
      m.slice(1),
    )
    const expectedOrder = ORDER.filter((name) => seen.includes(name))
    expect(seen, `styles.xml element order: ${seen.join(', ')}`).toEqual(expectedOrder)
  })

  it('round-trips through the reader: what was written is what comes back', async () => {
    // The two halves of this feature are separate code; this is the only test
    // that holds them to each other.
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 1, col: 0 }), {
      bold: true,
      italic: true,
      fill: '#00B0F0',
      color: '#FF0000',
      align: 'center',
    })
    const out = await saved(original, formatted)

    const reloaded = await load(asArrayBuffer(out.bytes))
    const back = formatForCell(reloaded.sheets[0], 1, 0)
    expect(back.bold).toBe(true)
    expect(back.italic).toBe(true)
    expect(back.fill).toBe('#00B0F0')
    expect(back.color).toBe('#FF0000')
    expect(back.align).toBe('center')

    // And a neighbour it never touched is still unformatted.
    expect(formatForCell(reloaded.sheets[0], 1, 2).bold).toBe(false)
  })

  it('follows its cell through a row insert, in the saved file', async () => {
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 2, col: 0 }), { fill: '#FF00FF' })
    const shifted = insertRowAt(formatted, 0, 1)
    const out = await saved(original, shifted)

    // "Ink" moved from row 3 to row 4, and its fill came with it.
    const reloaded = await load(asArrayBuffer(out.bytes))
    expect(reloaded.sheets[0].rows[3][0]).toBe('Ink')
    expect(formatForCell(reloaded.sheets[0], 3, 0).fill).toBe('#FF00FF')
    expect(formatForCell(reloaded.sheets[0], 1, 0).fill).toBeUndefined()
  })

  it('writes an empty cell that exists only to carry a fill', async () => {
    // Filling a blank cell is a normal thing to do (a spacer column, a legend
    // box). The cell has no text, so it only survives if the writer emits a
    // styled-but-empty `<c>`.
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, only({ row: 2, col: 1 }), { fill: '#FFFF00' })
    const out = await saved(original, formatted)

    expect(styleOf(out.sheet, 'B3'), 'B3 should be written with a style').not.toBeNull()
    const reloaded = await load(asArrayBuffer(out.bytes))
    expect(formatForCell(reloaded.sheets[0], 2, 1).fill).toBe('#FFFF00')
  })
})
