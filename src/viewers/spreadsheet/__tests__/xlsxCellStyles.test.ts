/**
 * SHEETFMT-1 — reading the cell formatting a workbook carries.
 *
 * Atlas rendered every workbook as unstyled text before this: a bold
 * white-on-blue header with a currency column looked exactly like a CSV. These
 * tests pin the reading of each piece of that, and — just as importantly — the
 * cases where guessing would be worse than inheriting (a theme colour, a
 * gradient fill, the "automatic" palette slots).
 *
 * The fixtures are real OOXML packages built here, not hand-written expected
 * values: the whole point is to be right about what Excel's own output means.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  DEFAULT_CELL_FORMAT,
  formatAt,
  parseStyleTable,
  readSheetStyleIds,
  readWorkbookCellStyles,
} from '../xlsxCellStyles'
import { buildStyledWorkbook, STYLES_XML } from './styledWorkbook'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/**
 * A styles part exercising every piece this module reads, in the shapes Excel
 * itself writes them:
 *   - font 1: bold + italic + underline + strike, 14pt Arial, explicit red
 *   - font 2: a THEME colour, which must stay unresolved
 *   - fill 2: a solid fill whose visible colour is `fgColor` (the easy mistake)
 *   - fill 3: a gradient, which must not be approximated
 *   - border 1: a real bottom edge; border 2: every edge `none`
 *   - xf 5: centred + wrapped, middle-aligned
 */
function richStyles(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<styleSheet xmlns="${NS}">` +
    `<numFmts count="1"><numFmt numFmtId="166" formatCode="0.0&quot;%&quot;"/></numFmts>` +
    `<fonts count="4">` +
    `<font><sz val="11"/><name val="Calibri"/></font>` +
    `<font><b/><i/><u/><strike/><sz val="14"/><name val="Arial"/><color rgb="FFFF0000"/></font>` +
    `<font><sz val="11"/><color theme="4"/><name val="Calibri"/></font>` +
    `<font><u val="none"/><sz val="11"/><color indexed="10"/><name val="Calibri"/></font>` +
    `</fonts>` +
    `<fills count="4">` +
    `<fill><patternFill patternType="none"/></fill>` +
    `<fill><patternFill patternType="gray125"/></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF00B0F0"/><bgColor indexed="64"/></patternFill></fill>` +
    `<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"/></stop></gradientFill></fill>` +
    `</fills>` +
    `<borders count="3">` +
    `<border><left/><right/><top/><bottom/><diagonal/></border>` +
    `<border><left/><right/><top/><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>` +
    `<border><left style="none"/><right style="none"/><top style="none"/><bottom style="none"/><diagonal/></border>` +
    `</borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="7">` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
    `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
    `<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>` +
    `<xf numFmtId="0" fontId="0" fillId="3" borderId="2" xfId="0" applyFill="1"/>` +
    `<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1">` +
    `<alignment horizontal="center" vertical="center" wrapText="1"/></xf>` +
    `<xf numFmtId="44" fontId="3" fillId="0" borderId="0" xfId="0"/>` +
    `</cellXfs></styleSheet>`
  )
}

describe('parseStyleTable (SHEETFMT-1)', () => {
  it('reads every run property off a font', () => {
    const xfs = parseStyleTable(richStyles())!.xfs
    expect(xfs[1]).toMatchObject({
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      fontSize: 14,
      fontName: 'Arial',
      color: '#FF0000',
    })
  })

  it('treats the workbook default font as a baseline, not as a per-cell format', () => {
    // fonts[0] is the document default and virtually every <xf> references it,
    // stating a size ("11") and a name ("Calibri"). Reporting those as chosen
    // formatting makes Atlas override its own theme font for EVERY workbook,
    // including completely unstyled ones — which is what this caught when the
    // grid started pinning every cell to 16px.
    const xfs = parseStyleTable(richStyles())!.xfs
    expect(xfs[0].fontSize).toBeUndefined()
    expect(xfs[0].fontName).toBeUndefined()
    // A font that genuinely differs from the default still reports its size.
    expect(xfs[1].fontSize).toBe(14)
    expect(xfs[1].fontName).toBe('Arial')
  })

  it('leaves a theme colour unresolved rather than guessing one', () => {
    // `theme="4"` indexes the theme part's colour scheme. Picking an entry here
    // without reading that part produces a confidently wrong colour, which is
    // worse for the reader than inheriting the grid's own text colour.
    expect(parseStyleTable(richStyles())!.xfs[2].color).toBeUndefined()
  })

  it('treats <u val="none"/> as not underlined', () => {
    // The trap: a generic "any non-zero val is on" reader reports this as
    // underlined, because the val is the string "none".
    expect(parseStyleTable(richStyles())!.xfs[6].underline).toBe(false)
  })

  it('resolves an indexed palette colour', () => {
    expect(parseStyleTable(richStyles())!.xfs[6].color).toBe('#FF0000')
  })

  it('reads a solid fill from fgColor, not bgColor', () => {
    // In a solid patternFill the visible colour is fgColor; bgColor is almost
    // always `indexed="64"` (automatic). Reading bgColor gives the wrong answer
    // on the single most common case there is — a plain highlighted cell.
    expect(parseStyleTable(richStyles())!.xfs[3].fill).toBe('#00B0F0')
  })

  it('ignores a pattern or gradient fill rather than approximating it', () => {
    const xfs = parseStyleTable(richStyles())!.xfs
    expect(xfs[4].fill).toBeUndefined() // gradient
    expect(xfs[0].fill).toBeUndefined() // patternType="none"
  })

  it('reports a border only when an edge has a real style', () => {
    const xfs = parseStyleTable(richStyles())!.xfs
    expect(xfs[3].bordered).toBe(true) // bottom="thin"
    expect(xfs[4].bordered).toBe(false) // every edge style="none"
    expect(xfs[0].bordered).toBe(false) // no styles at all
  })

  it('reads alignment and wrapping', () => {
    expect(parseStyleTable(richStyles())!.xfs[5]).toMatchObject({
      align: 'center',
      valign: 'middle',
      wrap: true,
    })
  })

  it('resolves a custom number format by id', () => {
    expect(parseStyleTable(richStyles())!.xfs[5].numberFormat).toBe('0.0"%"')
  })

  it('resolves a builtin number format that the file never spells out', () => {
    // numFmtId 44 (accounting) appears in no `<numFmts>` anywhere — Excel's
    // currency buttons just reference it by id. Without the builtin table a
    // whole currency column reports as General.
    expect(parseStyleTable(richStyles())!.xfs[6].numberFormat).toContain('#,##0.00')
  })

  it('reports General as undefined, not as the string "General"', () => {
    expect(parseStyleTable(richStyles())!.xfs[0].numberFormat).toBeUndefined()
  })

  it('returns null for a styles part with no cellXfs, and for junk', () => {
    expect(parseStyleTable(`<styleSheet xmlns="${NS}"><fonts count="0"/></styleSheet>`)).toBeNull()
    expect(parseStyleTable('not xml at all <<<')).toBeNull()
  })

  it('reads the save-path fixture the same way the save path does', () => {
    // Shared with the passthrough tests, so a change to one notices the other.
    const xfs = parseStyleTable(STYLES_XML)!.xfs
    expect(xfs[2].bold).toBe(true)
    expect(xfs[3].numberFormat).toBe('dd/mm/yyyy')
    expect(xfs[1].bold).toBe(false)
  })
})

describe('readSheetStyleIds (SHEETFMT-1)', () => {
  const sheet =
    `<worksheet xmlns="${NS}"><sheetData>` +
    `<row r="1"><c r="A1" s="2" t="s"><v>0</v></c><c r="B1"><v>1</v></c><c r="C1" s="7"/></row>` +
    `<row r="3"><c r="A3" s="4"><v>5</v></c></row>` +
    `</sheetData></worksheet>`

  it('maps each cell to its style index, defaulting to 0', () => {
    const ids = readSheetStyleIds(sheet, 3, 3, 0, 0)
    expect(ids[0]).toEqual([2, 0, 7])
    expect(ids[1]).toEqual([0, 0, 0]) // row 2 absent from the file
    expect(ids[2]).toEqual([4, 0, 0])
  })

  it('shifts into the grid coordinates when the used range does not start at A1', () => {
    // `sheetToGrid` builds its rows from the sheet's `!ref`, so a sheet whose
    // data starts at B3 has grid (0,0) === B3. A style read against absolute
    // addresses would land every format one row/column off.
    const offset = `<worksheet xmlns="${NS}"><sheetData><row r="3"><c r="B3" s="5"/></row></sheetData></worksheet>`
    expect(readSheetStyleIds(offset, 1, 1, 2, 1)).toEqual([[5]])
  })

  it('ignores a cell outside the range it was asked for', () => {
    const ids = readSheetStyleIds(sheet, 1, 1, 0, 0)
    expect(ids).toEqual([[2]])
  })

  it('reads a self-closing cell', () => {
    // `<c r="B1" s="3"/>` — a styled but EMPTY cell, which is exactly how a
    // blank-but-formatted cell (a coloured spacer column, a bordered box) is
    // written, and the case a scan for `<c …></c>` would miss entirely.
    const selfClosing = `<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" s="3"/></row></sheetData></worksheet>`
    expect(readSheetStyleIds(selfClosing, 1, 1, 0, 0)).toEqual([[3]])
  })

  it('does not mistake a <col> or <conditionalFormatting> element for a cell', () => {
    // The real trap for a tag-name scan: a worksheet's `<cols>` block carries
    // `style="…"` on `<col>`, and `<c` is a prefix of `<col`. Matching it would
    // assign a column's own style to whatever cell address came along with it —
    // or, worse, to cell (0,0) by accident.
    const withCols =
      `<worksheet xmlns="${NS}">` +
      `<cols><col min="1" max="1" width="24" style="6" customWidth="1"/></cols>` +
      `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
      `<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="0" priority="1"/></conditionalFormatting>` +
      `</worksheet>`
    expect(readSheetStyleIds(withCols, 1, 1, 0, 0)).toEqual([[0]])
  })

  it('is not confused by an s= inside cell text', () => {
    // The scan reads opening-tag attributes only, so a cell whose TEXT happens
    // to contain `s="9"` is still unstyled.
    const tricky =
      `<worksheet xmlns="${NS}"><sheetData><row r="1">` +
      `<c r="A1" t="inlineStr"><is><t>s="9" is not an attribute here</t></is></c>` +
      `</row></sheetData></worksheet>`
    expect(readSheetStyleIds(tricky, 1, 1, 0, 0)).toEqual([[0]])
  })
})

describe('readWorkbookCellStyles (SHEETFMT-1)', () => {
  it('reads a real package end to end, cell by cell', async () => {
    const buffer = await buildStyledWorkbook()
    const result = await readWorkbookCellStyles(buffer, [
      { sourcePath: 'xl/worksheets/sheet1.xml', rowCount: 3, colCount: 3, offsetRow: 0, offsetCol: 0 },
    ])
    expect(result).not.toBeNull()

    const styles = result!.styles.get('xl/worksheets/sheet1.xml')
    // The header row is bold (s="2" -> fontId 1), the body is not.
    expect(formatAt(styles, 0, 0).bold).toBe(true)
    expect(formatAt(styles, 0, 2).bold).toBe(true)
    expect(formatAt(styles, 1, 0).bold).toBe(false)
    // And the date column carries its format.
    expect(formatAt(styles, 1, 1).numberFormat).toBe('dd/mm/yyyy')
  })

  it('returns null for a package with no styles.xml, and for a non-zip buffer', async () => {
    const bare = new JSZip()
    bare.file('hello.txt', 'not a workbook')
    const noStyles = await bare.generateAsync({ type: 'arraybuffer' })
    expect(await readWorkbookCellStyles(noStyles, [])).toBeNull()

    // A `.csv` or `.xls` reaching this by accident must not throw — formatting
    // is a visual extra and may never be the reason a file won't open.
    expect(await readWorkbookCellStyles(new TextEncoder().encode('a,b,c').buffer as ArrayBuffer, [])).toBeNull()
  })

  it('keeps the rest of the workbook when one worksheet part is missing', async () => {
    const buffer = await buildStyledWorkbook()
    const result = await readWorkbookCellStyles(buffer, [
      { sourcePath: 'xl/worksheets/sheet1.xml', rowCount: 3, colCount: 3, offsetRow: 0, offsetCol: 0 },
      { sourcePath: 'xl/worksheets/nope.xml', rowCount: 1, colCount: 1, offsetRow: 0, offsetCol: 0 },
    ])
    expect(result!.styles.has('xl/worksheets/sheet1.xml')).toBe(true)
    expect(result!.styles.has('xl/worksheets/nope.xml')).toBe(false)
  })

  it('shares one format table across cells instead of storing a format per cell', async () => {
    // The shape this module exists to keep: a real workbook has a few dozen
    // formats across a million cells, so the per-cell cost must stay one small
    // integer. A regression to per-cell objects would not fail any assertion
    // above, and would quietly multiply memory on a 100k-row sheet.
    const buffer = await buildStyledWorkbook()
    const result = await readWorkbookCellStyles(buffer, [
      { sourcePath: 'xl/worksheets/sheet1.xml', rowCount: 3, colCount: 3, offsetRow: 0, offsetCol: 0 },
    ])
    const styles = result!.styles.get('xl/worksheets/sheet1.xml')!
    expect(styles.formats.length).toBe(5) // the fixture's five cellXfs entries
    for (const row of styles.styleIds) {
      for (const id of row) expect(typeof id).toBe('number')
    }
  })
})

describe('formatAt (SHEETFMT-1)', () => {
  it('falls back to the default format for an unstyled sheet or an out-of-range cell', () => {
    expect(formatAt(undefined, 0, 0)).toBe(DEFAULT_CELL_FORMAT)
    const styles = { formats: [DEFAULT_CELL_FORMAT], styleIds: [[0]] }
    expect(formatAt(styles, 99, 99)).toBe(DEFAULT_CELL_FORMAT)
  })

  it('falls back when a cell points at a style index the file never defined', () => {
    // A corrupt or truncated styles part must render unstyled, not crash the
    // grid mid-scroll.
    const styles = { formats: [DEFAULT_CELL_FORMAT], styleIds: [[7]] }
    expect(formatAt(styles, 0, 0)).toBe(DEFAULT_CELL_FORMAT)
  })
})
