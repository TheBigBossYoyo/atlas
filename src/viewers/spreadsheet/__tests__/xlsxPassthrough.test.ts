/**
 * USR-17 — saving an .xlsx keeps everything Atlas does not model: cell
 * styles, number formats, conditional formatting, column widths, and the
 * exact contents of every cell the user did not touch.
 */
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { beforeEach, describe, expect, it } from 'vitest'

import { attachSheetSources, attachTables, parseWorkbookBuffer } from '../../shared/spreadsheetGrid'
import {
  createDocument,
  deleteRowAt,
  insertRowAt,
  renameSheet,
  setCellValue,
  addSheet,
  type SpreadsheetDocument,
} from '../spreadsheetDocument'
import { readSheetPartPaths, readSheetTables } from '../spreadsheetTables'
import { writeWorkbookThroughOriginal } from '../xlsxPassthrough'
import { buildStyledWorkbook, STYLES_XML } from './styledWorkbook'

async function load(buffer: ArrayBuffer): Promise<SpreadsheetDocument> {
  const [tables, partPaths] = await Promise.all([readSheetTables(buffer), readSheetPartPaths(buffer)])
  return createDocument(attachSheetSources(attachTables(parseWorkbookBuffer(buffer), tables), partPaths))
}

async function saved(original: ArrayBuffer, doc: SpreadsheetDocument) {
  const bytes = await writeWorkbookThroughOriginal(original, doc)
  expect(bytes).not.toBeNull()
  const zip = await JSZip.loadAsync(bytes!)
  return {
    zip,
    sheet: await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
    workbook: await zip.file('xl/workbook.xml')!.async('string'),
  }
}

let original: ArrayBuffer
beforeEach(async () => {
  original = await buildStyledWorkbook()
})

describe('writeWorkbookThroughOriginal', () => {
  it('keeps untouched cells, styles and everything outside the model exactly as they were', async () => {
    const doc = setCellValue(await load(original), 0, 1, 0, 'Recycled paper')
    const { zip, sheet } = await saved(original, doc)

    // The edited cell carries its new value AND the style it had.
    expect(sheet).toContain('<t xml:space="preserve">Recycled paper</t>')
    // Untouched cells keep their type, style and cached value byte-for-byte.
    expect(sheet).toContain('<c r="B2" s="3"><v>45000</v></c>')
    expect(sheet).toContain('<c r="C3" s="4"><f>C2*2</f><v>25</v></c>')
    expect(sheet).toContain('<row r="1" ht="30" customHeight="1" s="2" customFormat="1">')
    // Parts Atlas does not model are passed through untouched.
    expect(await zip.file('xl/styles.xml')!.async('string')).toBe(STYLES_XML)
    expect(sheet).toContain('<conditionalFormatting sqref="C2:C3">')
    expect(sheet).toContain('<col min="1" max="1" width="24" customWidth="1"/>')
    expect(await zip.file('xl/media/image1.png')!.async('uint8array')).toEqual(new Uint8Array([137, 80, 78, 71]))
  })

  it('drops the stale calculation chain and asks Excel to recalculate on load', async () => {
    const doc = setCellValue(await load(original), 0, 1, 2, '99')
    const { zip, sheet, workbook } = await saved(original, doc)

    expect(zip.file('xl/calcChain.xml')).toBeNull()
    expect(await zip.file('[Content_Types].xml')!.async('string')).not.toContain('calcChain')
    expect(workbook).toContain('fullCalcOnLoad="1"')
    expect(workbook).toContain('<definedName name="Costs">')
    // A number typed by the user is written as a number, keeping the currency style.
    expect(sheet).toContain('<c r="C2" s="4"><v>99</v></c>')
  })

  it('moves row styling with the row when rows are inserted or deleted', async () => {
    let doc = await load(original)
    doc = insertRowAt(doc, 0, 1)
    const afterInsert = await saved(original, doc)
    // The header keeps its style on row 1; the old row 2 (with its date style) is now row 3.
    expect(afterInsert.sheet).toContain('<row r="1" ht="30" customHeight="1" s="2" customFormat="1">')
    expect(afterInsert.sheet).toContain('<c r="B3" s="3"><v>45000</v></c>')

    doc = deleteRowAt(doc, 0, 1)
    doc = deleteRowAt(doc, 0, 1)
    const afterDelete = await saved(original, doc)
    expect(afterDelete.sheet).toContain('<c r="B2" s="3"><v>45010</v></c>')
  })

  it('renames a sheet in the workbook part', async () => {
    const doc = renameSheet(await load(original), 0, 'Plan')
    const { workbook } = await saved(original, doc)
    expect(workbook).toContain('name="Plan"')
  })

  it('declines (so the caller falls back) for a non-OOXML buffer or a structural sheet change', async () => {
    const doc = await load(original)
    expect(await writeWorkbookThroughOriginal(new TextEncoder().encode('a,b').buffer as ArrayBuffer, doc)).toBeNull()
    expect(await writeWorkbookThroughOriginal(original, addSheet(doc))).toBeNull()

    const plain = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(plain, XLSX.utils.aoa_to_sheet([['a']]), 'S')
    const freshBytes = XLSX.write(plain, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
    // A document built without source paths (e.g. from CSV) cannot be patched.
    const fromRows = createDocument(parseWorkbookBuffer(freshBytes))
    expect(await writeWorkbookThroughOriginal(freshBytes, fromRows)).toBeNull()
  })

  it('round-trips through the parser with the edit applied', async () => {
    const doc = setCellValue(await load(original), 0, 2, 0, 'Blue ink')
    const bytes = (await writeWorkbookThroughOriginal(original, doc))!
    const reopened = await load(bytes.buffer as ArrayBuffer)
    expect(reopened.sheets[0].rows[2][0]).toBe('Blue ink')
    expect(reopened.sheets[0].rows[1][0]).toBe('Paper')
    expect(reopened.sheets[0].formulas[2][2]).toBe('C2*2')
  })
})
