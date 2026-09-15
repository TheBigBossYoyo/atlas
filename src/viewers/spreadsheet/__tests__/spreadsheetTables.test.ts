/**
 * USR-17 — Excel tables survive load -> edit -> save.
 */
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'

import { createDocument, deleteColumnAt, deleteRowAt, insertColumnAt, insertRowAt, setCellValue } from '../spreadsheetDocument'
import { writeWorkbookBytesWithTables } from '../spreadsheetWrite'
import {
  decodeRange,
  parseTableXml,
  readSheetTables,
  tableHeaderNames,
  tablesAfterColumnDelete,
  tablesAfterColumnInsert,
  tablesAfterRowDelete,
  tablesAfterRowInsert,
  type SheetTable,
} from '../spreadsheetTables'
import { attachTables, parseWorkbookBuffer } from '../../shared/spreadsheetGrid'

const TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Fruit" displayName="Fruit" ref="A1:C3" totalsRowShown="0">
  <autoFilter ref="A1:C3"><filterColumn colId="1"><filters><filter val="3"/></filters></filterColumn></autoFilter>
  <tableColumns count="3">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Qty"/>
    <tableColumn id="3" name="Price"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`

/** A SheetJS workbook with an Excel table added the way Excel stores one. */
async function workbookWithTable(): Promise<ArrayBuffer> {
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['Name', 'Qty', 'Price'],
      ['Apple', 3, 1.5],
      ['Pear', 5, 2],
    ]),
    'Data',
  )
  const zip = await JSZip.loadAsync(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer)
  zip.file('xl/tables/table1.xml', TABLE_XML)
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`,
  )
  return zip.generateAsync({ type: 'arraybuffer' })
}

function table(overrides: Partial<SheetTable> = {}): SheetTable {
  return { ...parseTableXml(TABLE_XML)!, ...overrides }
}

async function loadDocument(buffer: ArrayBuffer) {
  const tables = await readSheetTables(buffer)
  return createDocument(attachTables(parseWorkbookBuffer(buffer), tables))
}

describe('readSheetTables', () => {
  it('reads a table definition from the sheet relationships', async () => {
    const tables = await readSheetTables(await workbookWithTable())
    expect(tables.Data).toHaveLength(1)
    expect(tables.Data[0]).toMatchObject({ name: 'Fruit', r0: 0, c0: 0, r1: 2, c1: 2, headerRow: true, totalsRow: false })
  })

  it('resolves {} for a workbook without tables and for non-zip input', async () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'S')
    expect(await readSheetTables(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer)).toEqual({})
    expect(await readSheetTables(new TextEncoder().encode('a,b').buffer as ArrayBuffer)).toEqual({})
  })
})

describe('table range bookkeeping', () => {
  it('decodes absolute and single-cell refs', () => {
    expect(decodeRange('$B$2:$D$10')).toEqual({ r0: 1, c0: 1, r1: 9, c1: 3 })
    expect(decodeRange('AA5')).toEqual({ r0: 4, c0: 26, r1: 4, c1: 26 })
    expect(decodeRange('nonsense')).toBeNull()
  })

  it('grows for a row inserted inside, shifts for one above, ignores one below', () => {
    expect(tablesAfterRowInsert([table()], 2)[0]).toMatchObject({ r0: 0, r1: 3 })
    expect(tablesAfterRowInsert([table()], 0)[0]).toMatchObject({ r0: 1, r1: 3 })
    expect(tablesAfterRowInsert([table()], 3)[0]).toMatchObject({ r0: 0, r1: 2 })
  })

  it('drops the table when its header row or last data row is deleted', () => {
    expect(tablesAfterRowDelete([table()], 0)).toEqual([])
    const shrunk = tablesAfterRowDelete([table()], 1)
    expect(shrunk[0]).toMatchObject({ r1: 1 })
    expect(tablesAfterRowDelete(shrunk, 1)).toEqual([])
  })

  it('tracks where each column came from across column inserts and deletes', () => {
    const inserted = tablesAfterColumnInsert([table()], 1)[0]
    expect(inserted).toMatchObject({ c0: 0, c1: 3, columnSources: [0, null, 1, 2] })
    const deleted = tablesAfterColumnDelete([inserted], 2)[0]
    expect(deleted).toMatchObject({ c1: 2, columnSources: [0, null, 2] })
    expect(tablesAfterColumnDelete([table({ c1: 0, columnSources: [0] })], 0)).toEqual([])
  })

  it('names columns from header cells, filling blanks and de-duplicating', () => {
    const rows = [['Name', '', 'name']]
    expect(tableHeaderNames(table(), rows)).toEqual(['Name', 'Column2', 'name2'])
  })
})

describe('save keeps Excel tables', () => {
  it('round-trips a table through edits, row/column inserts and save', async () => {
    let doc = await loadDocument(await workbookWithTable())
    doc = setCellValue(doc, 0, 1, 0, 'Plum')
    doc = setCellValue(doc, 0, 0, 1, 'Count')
    doc = insertRowAt(doc, 0, 2)
    doc = insertColumnAt(doc, 0, 3)
    doc = deleteColumnAt(doc, 0, 3)

    const bytes = await writeWorkbookBytesWithTables(doc, 'xlsx')
    const zip = await JSZip.loadAsync(bytes)
    const tableXml = await zip.file('xl/tables/table1.xml')!.async('string')
    expect(tableXml).toContain('ref="A1:C4"')
    expect(tableXml).toContain('<autoFilter ref="A1:C4"')
    expect(tableXml).not.toContain('filterColumn')
    expect(tableXml).toMatch(/name="Count"/)
    expect(tableXml).toContain('TableStyleMedium2')

    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheetXml).toMatch(/<tableParts count="1"><tablePart r:id="rIdAtlasTable1"\/><\/tableParts><\/worksheet>$/)
    expect(await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string')).toContain('../tables/table1.xml')
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/xl/tables/table1.xml')

    // Reopening sees the same table and the edited values.
    const reopened = await loadDocument(bytes.buffer as ArrayBuffer)
    expect(reopened.sheets[0].tables?.[0]).toMatchObject({ name: 'Fruit', r1: 3, c1: 2 })
    expect(reopened.sheets[0].rows[1][0]).toBe('Plum')
  })

  it('writes a numeric or blank header cell as the column name text', async () => {
    let doc = await loadDocument(await workbookWithTable())
    doc = setCellValue(doc, 0, 0, 1, '2024')
    doc = setCellValue(doc, 0, 0, 2, '')
    const wb = XLSX.read(await writeWorkbookBytesWithTables(doc, 'xlsx'))
    const sheet = wb.Sheets.Data
    expect(sheet.B1).toMatchObject({ t: 's', v: '2024' })
    expect(sheet.C1).toMatchObject({ t: 's', v: 'Column3' })
  })

  it('leaves non-OOXML output and table-free workbooks untouched', async () => {
    const doc = await loadDocument(await workbookWithTable())
    const ods = await writeWorkbookBytesWithTables(doc, 'ods')
    expect((await JSZip.loadAsync(ods)).file('xl/tables/table1.xml')).toBeNull()

    const plain = deleteRowAt(doc, 0, 0) // deleting the header row un-tables the range
    expect(plain.sheets[0].tables).toEqual([])
    const zip = await JSZip.loadAsync(await writeWorkbookBytesWithTables(plain, 'xlsx'))
    expect(zip.file('xl/tables/table1.xml')).toBeNull()
  })
})
