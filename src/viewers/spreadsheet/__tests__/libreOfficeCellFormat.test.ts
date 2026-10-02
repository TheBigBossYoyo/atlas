/**
 * SHEETFMT-2, external verification — does another program see the formatting
 * Atlas applied?
 *
 * Every other test in this feature is Atlas checking its own work: it reads its
 * own `styles.xml` back, or asserts on XML shapes this codebase itself decided
 * on. That can all be self-consistently wrong. The claim being made is
 * interoperability — "bold a cell in Atlas and Excel shows it bold" — and only
 * a different implementation can test it.
 *
 * LibreOffice is that implementation. Its Calc HTML export carries real CSS, so
 * `font-weight: bold`, `font-style: italic`, a background colour and an
 * alignment all show up in the output as properties LibreOffice itself computed
 * from the package Atlas wrote.
 *
 * Skips when LibreOffice is not installed. A green suite therefore does NOT
 * prove this ran — the same caveat `libreOfficeRoundTrip.test.ts` carries.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { describe, expect, it } from 'vitest'

import { attachSheetSources, attachTables, parseWorkbookBuffer } from '../../shared/spreadsheetGrid'
import { createDocument, setRangeFormat, type SpreadsheetDocument } from '../spreadsheetDocument'
import { readSheetPartPaths, readSheetTables } from '../spreadsheetTables'
import { writeWorkbookThroughOriginal } from '../xlsxPassthrough'
import { CONVERT, convert, findSoffice } from '../../../__tests__/helpers/libreOffice'
import { buildStyledWorkbook } from './styledWorkbook'

const soffice = findSoffice()
const describeMaybe = soffice ? describe : describe.skip

async function load(buffer: ArrayBuffer): Promise<SpreadsheetDocument> {
  const [tables, partPaths] = await Promise.all([readSheetTables(buffer), readSheetPartPaths(buffer)])
  return createDocument(attachSheetSources(attachTables(parseWorkbookBuffer(buffer), tables), partPaths))
}

/** Normalises the CSS LibreOffice emits so a declaration can be matched regardless of spacing. */
function css(html: string): string {
  return html.replace(/\s+/g, ' ').toLowerCase()
}

describeMaybe('LibreOffice reads the cell formatting Atlas applied (SHEETFMT-2)', () => {
  it('sees bold, italic, a fill and an alignment that Atlas wrote', async () => {
    const original = await buildStyledWorkbook()
    const doc = await load(original)
    // A2 is "Paper", plain in the fixture. Format it every way at once so one
    // conversion covers all four properties.
    const formatted = setRangeFormat(doc, 0, { row0: 1, col0: 0, row1: 1, col1: 0 }, {
      bold: true,
      italic: true,
      fill: '#FF0000',
      align: 'center',
    })
    const bytes = await writeWorkbookThroughOriginal(original, formatted)
    expect(bytes).not.toBeNull()

    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-fmt-'))
    const file = path.join(dir, 'formatted.xlsx')
    fs.writeFileSync(file, Buffer.from(bytes!))

    // `convert` returns the output file's PATH, not its contents.
    const out = convert(soffice!, file, CONVERT.calcHtml, dir)
    expect(out, 'LibreOffice could not read the formatted workbook at all').not.toBeNull()
    const styles = css(fs.readFileSync(out!, 'utf8'))

    // LibreOffice computed each of these from Atlas's package, not from
    // anything this repo told it.
    expect(styles, 'LibreOffice should render the cell bold').toMatch(/font-weight:\s*bold/)
    expect(styles, 'LibreOffice should render the cell italic').toMatch(/font-style:\s*italic/)
    expect(styles, 'LibreOffice should render the fill').toMatch(/background:\s*#ff0000|background-color:\s*#ff0000/)
    expect(styles, 'LibreOffice should render the alignment').toMatch(/text-align:\s*center/)

    // The self-check this whole harness needs: the comparison has to be able to
    // FAIL. The same workbook without the formatting must not produce those
    // declarations, or the assertions above would pass on anything.
    const plainBytes = await writeWorkbookThroughOriginal(original, doc)
    const plainFile = path.join(dir, 'plain.xlsx')
    fs.writeFileSync(plainFile, Buffer.from(plainBytes!))
    const plainOut = convert(soffice!, plainFile, CONVERT.calcHtml, dir)
    expect(plainOut, 'LibreOffice could not read the unformatted control').not.toBeNull()
    const plainStyles = css(fs.readFileSync(plainOut!, 'utf8'))
    // The control must be a real extraction, not an empty string that would
    // satisfy every `not.toMatch` below for the wrong reason.
    expect(plainStyles.length, 'the control extraction was empty').toBeGreaterThan(200)
    expect(plainStyles, 'the unformatted control should not be centred').not.toMatch(/text-align:\s*center/)
    expect(plainStyles, 'the unformatted control should have no red fill').not.toMatch(/#ff0000/)
  }, 120_000)

  it('sees a cell border Atlas applied (SHEETFMT-3)', async () => {
    const original = await buildStyledWorkbook()
    const doc = await load(original)
    const formatted = setRangeFormat(doc, 0, { row0: 1, col0: 0, row1: 1, col1: 0 }, {
      border: {
        top: undefined,
        right: undefined,
        bottom: { weight: 'medium', color: '#FF0000' },
        left: undefined,
      },
    })
    const bytes = await writeWorkbookThroughOriginal(original, formatted)
    expect(bytes).not.toBeNull()

    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-border-'))
    const file = path.join(dir, 'bordered.xlsx')
    fs.writeFileSync(file, Buffer.from(bytes!))

    const out = convert(soffice!, file, CONVERT.calcHtml, dir)
    expect(out, 'LibreOffice could not read the bordered workbook').not.toBeNull()
    const styles = css(fs.readFileSync(out!, 'utf8'))

    // LibreOffice emits per-edge LONGHAND properties, not the `border-bottom:`
    // shorthand — checked against its real output rather than assumed. The
    // colour is what proves it is reading Atlas's own border rather than
    // drawing a default grid line, and the three `none` edges prove it read
    // which edges Atlas asked for.
    expect(styles, `no bottom border colour in LibreOffice's output`).toMatch(
      /border-bottom-color:\s*#ff0000/,
    )
    expect(styles).toMatch(/border-bottom-style:\s*solid/)
    expect(styles).toMatch(/border-top-style:\s*none/)
    expect(styles).toMatch(/border-left-style:\s*none/)

    // Control: the unformatted workbook has no such edge.
    const plainBytes = await writeWorkbookThroughOriginal(original, doc)
    const plainFile = path.join(dir, 'plain.xlsx')
    fs.writeFileSync(plainFile, Buffer.from(plainBytes!))
    const plainOut = convert(soffice!, plainFile, CONVERT.calcHtml, dir)
    expect(plainOut).not.toBeNull()
    expect(css(fs.readFileSync(plainOut!, 'utf8'))).not.toMatch(/border-bottom-color:\s*#ff0000/)
  }, 120_000)

  it('sees a number format Atlas applied', async () => {
    const original = await buildStyledWorkbook()
    const doc = await load(original)
    // C2 holds 12.5; as a percentage LibreOffice must render 1250%, which is a
    // value no other reading of the cell produces.
    const formatted = setRangeFormat(doc, 0, { row0: 1, col0: 2, row1: 1, col1: 2 }, { numberFormat: '0%' })
    const bytes = await writeWorkbookThroughOriginal(original, formatted)

    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-numfmt-'))
    const file = path.join(dir, 'numfmt.xlsx')
    fs.writeFileSync(file, Buffer.from(bytes!))

    const out = convert(soffice!, file, CONVERT.calcHtml, dir)
    expect(out, 'LibreOffice could not read the workbook').not.toBeNull()
    expect(
      fs.readFileSync(out!, 'utf8'),
      'LibreOffice should apply the percent format Atlas wrote',
    ).toContain('1250%')
  }, 120_000)
})
