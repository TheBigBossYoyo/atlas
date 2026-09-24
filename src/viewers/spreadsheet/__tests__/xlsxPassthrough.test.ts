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
  addSheet,
  createDocument,
  deleteColumnAt,
  deleteRowAt,
  deleteSheet,
  insertColumnAt,
  insertRowAt,
  renameSheet,
  setCellValue,
  type SpreadsheetDocument,
} from '../spreadsheetDocument'
import { readSheetPartPaths, readSheetTables } from '../spreadsheetTables'
import { writeWorkbookThroughOriginal } from '../xlsxPassthrough'
import {
  buildCrossSheetFormulaWorkbook,
  buildMultiSheetWorkbook,
  buildSharedFormulaWorkbook,
  buildSharedStringsWorkbook,
  buildStyledWorkbook,
  buildWorkbookWithOwnedParts,
  buildWorkbookWithSharedImage,
  STYLES_XML,
} from './styledWorkbook'

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

  it('declines (so the caller falls back) for a non-OOXML buffer or an untracked document', async () => {
    const doc = await load(original)
    expect(await writeWorkbookThroughOriginal(new TextEncoder().encode('a,b').buffer as ArrayBuffer, doc)).toBeNull()

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

// ---------------------------------------------------------------------------
// SHEET-2 — an Excel table's filterColumn/sortState must survive the
// save-through-original path (`writeWorkbookThroughOriginal`, not just the
// fresh-workbook `graftTables` path covered in `spreadsheetTables.test.ts`)
// when the edit never touches that table's own columns.
// ---------------------------------------------------------------------------

const TABLE_WITH_FILTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Items" displayName="Items" ref="A1:C3" totalsRowShown="0">
  <autoFilter ref="A1:C3"><filterColumn colId="1"><filters><filter val="45000"/></filters></filterColumn></autoFilter>
  <sortState ref="A2:C3"><sortCondition ref="B2:B3"/></sortState>
  <tableColumns count="3">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Due"/>
    <tableColumn id="3" name="Cost"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`

async function buildWorkbookWithFilteredTable(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await buildStyledWorkbook())
  zip.file('xl/tables/table1.xml', TABLE_WITH_FILTER_XML)
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheetXml.replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rId9"/></tableParts></worksheet>'),
  )
  const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    contentTypes.replace(
      '</Types>',
      '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('writeWorkbookThroughOriginal — Excel table filter/sort state', () => {
  it("keeps an untouched table's filterColumn and sortState when an unrelated cell is edited", async () => {
    const withTable = await buildWorkbookWithFilteredTable()
    const doc = setCellValue(await load(withTable), 0, 1, 0, 'Recycled paper') // edits A2 — nothing to do with the table's columns
    const bytes = (await writeWorkbookThroughOriginal(withTable, doc))!
    expect(bytes).not.toBeNull()
    const zip = await JSZip.loadAsync(bytes)
    const tableXml = await zip.file('xl/tables/table1.xml')!.async('string')
    expect(tableXml).toContain('<filterColumn colId="1">')
    expect(tableXml).toContain('<filter val="45000"/>')
    expect(tableXml).toContain('<sortState ref="A2:C3"><sortCondition ref="B2:B3"/></sortState>')
  })

  it("re-anchors the filterColumn and drops sortState when a column insert moves the table's columns", async () => {
    const withTable = await buildWorkbookWithFilteredTable()
    let doc = await load(withTable)
    doc = insertColumnAt(doc, 0, 1) // insert a new column at B — Due (colId="1") moves to colId="2"
    const bytes = (await writeWorkbookThroughOriginal(withTable, doc))!
    expect(bytes).not.toBeNull()
    const zip = await JSZip.loadAsync(bytes)
    const tableXml = await zip.file('xl/tables/table1.xml')!.async('string')
    expect(tableXml).toContain('<filterColumn colId="2">')
    expect(tableXml).not.toContain('sortState')
  })
})

describe('writeWorkbookThroughOriginal — structural sheet changes', () => {
  let multi: ArrayBuffer
  beforeEach(async () => {
    multi = await buildMultiSheetWorkbook()
  })

  async function savedWorkbookXml(bytes: Uint8Array): Promise<string> {
    const zip = await JSZip.loadAsync(bytes)
    return zip.file('xl/workbook.xml')!.async('string')
  }

  it('adds a new worksheet part, sheet entry and content-type override', async () => {
    const doc = addSheet(await load(multi), 'Extra')
    const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    expect(bytes).not.toBeNull()

    const zip = await JSZip.loadAsync(bytes)
    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    expect(workbook).toContain('name="Extra"')
    // A brand-new part path distinct from the two originals.
    expect(zip.file('xl/worksheets/sheet3.xml')).not.toBeNull()
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/xl/worksheets/sheet3.xml')

    // Re-opens with SheetJS...
    const wb = XLSX.read(bytes)
    expect(wb.SheetNames).toEqual(['Budget', 'Notes', 'Extra'])
    // ...and with Atlas's own reader.
    const reopened = await load(bytes.buffer as ArrayBuffer)
    expect(reopened.sheets.map((s) => s.name)).toEqual(['Budget', 'Notes', 'Extra'])
  })

  it('deletes a sheet: its part, rels, content-type override and scoped defined names are gone; later localSheetIds are renumbered', async () => {
    const doc = deleteSheet(await load(multi), 0) // delete "Budget" — "Notes" becomes index 0
    const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    expect(bytes).not.toBeNull()

    const zip = await JSZip.loadAsync(bytes)
    expect(zip.file('xl/worksheets/sheet1.xml')).toBeNull()
    expect(zip.file('xl/worksheets/sheet2.xml')).not.toBeNull()
    expect(await zip.file('[Content_Types].xml')!.async('string')).not.toContain('sheet1.xml')

    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    // "BudgetHeader" (localSheetId=0, scoped to the deleted sheet) is gone.
    expect(workbook).not.toContain('BudgetHeader')
    // "Costs" referenced Budget!... and Budget no longer exists — dropped too.
    expect(workbook).not.toContain('name="Costs"')
    // "NotesRef" was localSheetId=1; Notes is now the only (index-0) sheet.
    expect(workbook).toContain('<definedName name="NotesRef" localSheetId="0">Notes!$A$1</definedName>')

    const reopened = await load(bytes.buffer as ArrayBuffer)
    expect(reopened.sheets.map((s) => s.name)).toEqual(['Notes'])
  })

  it('reorders sheets to match the document order', async () => {
    let doc = await load(multi)
    const [budget, notes] = doc.sheets
    doc = { sheets: [notes, budget] }
    const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    const workbook = await savedWorkbookXml(bytes)
    const order = Array.from(workbook.matchAll(/<sheet name="([^"]+)"/g)).map((m) => m[1])
    expect(order).toEqual(['Notes', 'Budget'])

    const reopened = await load(bytes.buffer as ArrayBuffer)
    expect(reopened.sheets.map((s) => s.name)).toEqual(['Notes', 'Budget'])
  })

  it('updates docProps/app.xml worksheet titles when adding a sheet, and leaves it alone when absent', async () => {
    // A real Excel-produced docProps/app.xml shape (namespaced vt:vector/vt:variant children).
    const appXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      `<Application>Microsoft Excel</Application>` +
      `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
      `<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Budget</vt:lpstr><vt:lpstr>Notes</vt:lpstr></vt:vector></TitlesOfParts>` +
      `</Properties>`
    const zip = await JSZip.loadAsync(multi)
    zip.file('docProps/app.xml', appXml)
    const withAppXml = await zip.generateAsync({ type: 'arraybuffer' })

    const doc = addSheet(await load(withAppXml), 'Extra')
    const bytes = (await writeWorkbookThroughOriginal(withAppXml, doc))!
    const saved = await JSZip.loadAsync(bytes)
    const savedApp = await saved.file('docProps/app.xml')!.async('string')
    expect(savedApp).toContain('<vt:lpstr>Budget</vt:lpstr><vt:lpstr>Notes</vt:lpstr><vt:lpstr>Extra</vt:lpstr>')
    expect(savedApp).toContain('size="3"')
    expect(savedApp).toContain('<vt:i4>3</vt:i4>')
  })

  it('renaming a sheet updates a single-area defined name that points at it', async () => {
    const doc = renameSheet(await load(multi), 0, 'Plan')
    const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    const workbook = await savedWorkbookXml(bytes)
    expect(workbook).toContain('<definedName name="Costs">Plan!$C$2:$C$3</definedName>')
    expect(workbook).toContain('<definedName name="BudgetHeader" localSheetId="0">Plan!$A$1:$C$1</definedName>')
  })

  it('re-anchors conditional formatting, data validation, a hyperlink and the autoFilter through a row insert', async () => {
    let doc = await load(multi)
    doc = insertRowAt(doc, 0, 1) // insert a blank row between the header and "Paper"
    const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    const zip = await JSZip.loadAsync(bytes)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')

    // C2:C3/B2:B3 sat entirely AT/AFTER the insertion point, so they shift
    // down whole (matching `insertRowAt`'s own merge-cell semantics); the
    // sheet-level autoFilter (A1:C3) STRADDLES the insertion point, so it
    // instead grows by one row to include it.
    expect(sheet).toContain('<conditionalFormatting sqref="C3:C4">')
    expect(sheet).toContain('sqref="B3:B4"')
    expect(sheet).toContain('<hyperlink ref="A3"')
    expect(sheet).toContain('<autoFilter ref="A1:C4"/>')
  })

  it('re-anchors through a row delete, and drops a range fully consumed by the delete', async () => {
    let doc = await load(multi)
    doc = deleteRowAt(doc, 0, 1) // delete the "Paper" row — the CF/DV range (rows 2:3) shrinks to row 2 only
    const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    const zip = await JSZip.loadAsync(bytes)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')

    // The single surviving row collapses the range to one cell — Atlas
    // writes a bare `C2`, not the equivalent-but-redundant `C2:C2`.
    expect(sheet).toContain('<conditionalFormatting sqref="C2">')
    expect(sheet).toContain('<autoFilter ref="A1:C2"/>')
  })
})

// ---------------------------------------------------------------------
// Cell formulas referencing a renamed/deleted sheet or shifted cells
// (USR-17 follow-up — see `formulaRefs.ts`).
// ---------------------------------------------------------------------

  describe('writeWorkbookThroughOriginal — cell formulas referencing a renamed/deleted sheet or shifted cells', () => {
    let crossSheet: ArrayBuffer
    beforeEach(async () => {
      crossSheet = await buildCrossSheetFormulaWorkbook()
    })

    it('rewrites a cloned formula\'s cross-sheet reference on a rename', async () => {
      const doc = renameSheet(await load(crossSheet), 1, 'Info') // Notes -> Info
      const bytes = (await writeWorkbookThroughOriginal(crossSheet, doc))!
      expect(bytes).not.toBeNull()
      const zip = await JSZip.loadAsync(bytes)
      const budget = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(budget).toContain('<f>Info!A1</f>')
    })

    it('turns a cloned formula\'s cross-sheet reference into #REF! when the target sheet is deleted, clearing its stale cached numeric value', async () => {
      // Notes!A2 is `<c r="A2"><f>Budget!C2</f><v>12.5</v></c>` — the old
      // cached `12.5` must not survive alongside the now-broken formula.
      const doc = deleteSheet(await load(crossSheet), 0) // delete Budget — Notes becomes the only sheet
      const bytes = (await writeWorkbookThroughOriginal(crossSheet, doc))!
      expect(bytes).not.toBeNull()
      const zip = await JSZip.loadAsync(bytes)
      const notes = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
      expect(notes).toContain('<f>#REF!</f>')
      expect(notes).not.toContain('12.5')
      // The cached value becomes the error itself (what Excel writes), not
      // nothing: a formula cell with no <v> is skipped entirely by SheetJS,
      // so the formula would look lost when the file is reopened.
      expect(notes).toContain('t="e"')
      expect(notes).toContain('<v>#REF!</v>')
    })

    it('re-anchors a cloned formula\'s cross-sheet coordinates through a row insert on the TARGET sheet', async () => {
      let doc = await load(crossSheet)
      doc = insertRowAt(doc, 0, 1) // insert a blank row into Budget, between the header and "Paper"
      const bytes = (await writeWorkbookThroughOriginal(crossSheet, doc))!
      const zip = await JSZip.loadAsync(bytes)
      const notes = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
      // Notes!A2 referenced Budget!C2 ("Paper"'s cost) — Budget's row 2 slid down to row 3.
      expect(notes).toContain('<f>Budget!C3</f>')
    })

    it('re-anchors a cloned formula\'s UNQUALIFIED (own-sheet) reference through its own row insert', async () => {
      let doc = await load(crossSheet)
      doc = insertRowAt(doc, 0, 1) // insert a blank row into Budget itself
      const bytes = (await writeWorkbookThroughOriginal(crossSheet, doc))!
      const zip = await JSZip.loadAsync(bytes)
      const budget = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      // Budget!D3 referenced its own B2 ("Paper"'s flag) — that row slid down to row 3.
      expect(budget).toContain('<f>B3</f>')
    })

    it('still fixes a stale sheet name on a freshly-typed formula, without touching its coordinates', async () => {
      let doc = await load(crossSheet)
      doc = setCellValue(doc, 0, 0, 3, '=Notes!A1') // user types this fresh into Budget!D1
      doc = renameSheet(doc, 1, 'Info') // Notes -> Info, decided AFTER the formula was typed
      const bytes = (await writeWorkbookThroughOriginal(crossSheet, doc))!
      const zip = await JSZip.loadAsync(bytes)
      const budget = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(budget).toContain('<f>Info!A1</f>')
    })

    it('does NOT re-anchor a freshly-typed formula\'s coordinates even when the referenced sheet had rows inserted', async () => {
      let doc = await load(crossSheet)
      doc = insertRowAt(doc, 0, 1) // insert a blank row into Budget — its rowSources now record a shift
      // The user types this AFTER the insert, so "C2" already means the CURRENT (post-insert) C2.
      doc = setCellValue(doc, 1, 0, 0, '=Budget!C2')
      const bytes = (await writeWorkbookThroughOriginal(crossSheet, doc))!
      const zip = await JSZip.loadAsync(bytes)
      const notes = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
      expect(notes).toContain('<f>Budget!C2</f>')
    })

    it('clears a cloned formula\'s stale cached STRING value when its cross-sheet reference becomes #REF!', async () => {
      // Budget!D2 is `<c r="D2" t="str"><f>Notes!A1</f><v>See Budget</v></c>` —
      // deleting Notes turns the formula into `#REF!`; the old cached "See
      // Budget" text (and its now-meaningless t="str") must not survive
      // alongside it.
      const doc = deleteSheet(await load(crossSheet), 1) // delete Notes — Budget becomes the only sheet
      const bytes = (await writeWorkbookThroughOriginal(crossSheet, doc))!
      expect(bytes).not.toBeNull()
      const zip = await JSZip.loadAsync(bytes)
      const budget = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(budget).toContain('<f>#REF!</f>')
      expect(budget).not.toContain('See Budget')
      expect(budget).not.toContain('D2" t="str"')
    })
  })

  // ---------------------------------------------------------------------
  // A shared-formula group's `ref` SPAN attribute is re-anchored right
  // alongside the master cell's own formula text (USR-17 follow-up).
  // ---------------------------------------------------------------------

  describe('writeWorkbookThroughOriginal — shared-formula ref span re-anchoring', () => {
    it('re-anchors the ref span (and the master formula text) through a row insert above the group', async () => {
      const original = await buildSharedFormulaWorkbook()
      let doc = await load(original)
      doc = insertRowAt(doc, 0, 0) // insert a blank row above row 1 — everything slides down by one
      const bytes = (await writeWorkbookThroughOriginal(original, doc))!
      expect(bytes).not.toBeNull()
      const zip = await JSZip.loadAsync(bytes)
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheet).toContain('<f t="shared" ref="B3:B5" si="0">A3*2</f>')
    })

    it('shrinks the ref span when a row insert/delete happens INSIDE the group', async () => {
      const original = await buildSharedFormulaWorkbook()
      let doc = await load(original)
      doc = deleteRowAt(doc, 0, 2) // delete row 3 (a follower member of the group, A3/B3)
      const bytes = (await writeWorkbookThroughOriginal(original, doc))!
      expect(bytes).not.toBeNull()
      const zip = await JSZip.loadAsync(bytes)
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      // The master's own row (row 2) never moved, so its formula text is
      // unchanged — only the group's ref span shrinks from B2:B4 to B2:B3.
      expect(sheet).toContain('<f t="shared" ref="B2:B3" si="0">A2*2</f>')
    })
  })

  // ---------------------------------------------------------------------
  // Generalized defined-name rewriting: multi-area, function-wrapped, and
  // whole-row/whole-column ranges (USR-17 follow-up).
  // ---------------------------------------------------------------------

  async function withExtraDefinedNames(original: ArrayBuffer, extra: string): Promise<ArrayBuffer> {
    const zip = await JSZip.loadAsync(original)
    const workbookXml = await zip.file('xl/workbook.xml')!.async('string')
    expect(workbookXml).toContain('</definedNames>')
    zip.file('xl/workbook.xml', workbookXml.replace('</definedNames>', `${extra}</definedNames>`))
    return zip.generateAsync({ type: 'arraybuffer' })
  }

  describe('writeWorkbookThroughOriginal — generalized defined-name rewriting', () => {
    it('rewrites every area of a multi-area name independently, keeping the comma', async () => {
      const multi = await withExtraDefinedNames(
        await buildMultiSheetWorkbook(),
        '<definedName name="MultiArea">Budget!$A$1,Notes!$A$1</definedName>',
      )
      const doc = renameSheet(await load(multi), 0, 'Plan')
      const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
      const workbook = await (await JSZip.loadAsync(bytes)).file('xl/workbook.xml')!.async('string')
      expect(workbook).toContain('<definedName name="MultiArea">Plan!$A$1,Notes!$A$1</definedName>')
    })

    it('keeps a multi-area name with a #REF! spliced into just the deleted area', async () => {
      const multi = await withExtraDefinedNames(
        await buildMultiSheetWorkbook(),
        '<definedName name="MultiArea">Budget!$A$1,Notes!$A$1</definedName>',
      )
      const doc = deleteSheet(await load(multi), 0) // delete Budget
      const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
      const workbook = await (await JSZip.loadAsync(bytes)).file('xl/workbook.xml')!.async('string')
      expect(workbook).toContain('<definedName name="MultiArea">#REF!,Notes!$A$1</definedName>')
    })

    it('rewrites a sheet reference nested inside a function-wrapped name', async () => {
      const multi = await withExtraDefinedNames(
        await buildMultiSheetWorkbook(),
        '<definedName name="Wrapped">OFFSET(Budget!$A$1,0,0,3,1)</definedName>',
      )
      const doc = renameSheet(await load(multi), 0, 'Plan')
      const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
      const workbook = await (await JSZip.loadAsync(bytes)).file('xl/workbook.xml')!.async('string')
      expect(workbook).toContain('<definedName name="Wrapped">OFFSET(Plan!$A$1,0,0,3,1)</definedName>')
    })

    it('re-anchors a whole-row range (a Print_Titles-style name) through a row insert', async () => {
      const multi = await withExtraDefinedNames(
        await buildMultiSheetWorkbook(),
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">Budget!$1:$1</definedName>',
      )
      let doc = await load(multi)
      doc = insertRowAt(doc, 0, 0) // insert a blank row before Budget's own header row
      const bytes = (await writeWorkbookThroughOriginal(multi, doc))!
      const workbook = await (await JSZip.loadAsync(bytes)).file('xl/workbook.xml')!.async('string')
      expect(workbook).toContain('<definedName name="_xlnm.Print_Titles" localSheetId="0">Budget!$2:$2</definedName>')
    })
  })

  // ---------------------------------------------------------------------
  // A deleted sheet's own exclusively-owned parts are swept, not orphaned
  // (USR-17 follow-up).
  // ---------------------------------------------------------------------

  describe('writeWorkbookThroughOriginal — sweeps a deleted sheet\'s exclusively-owned parts', () => {
    it('removes the table, comments, VML note-shape drawing, drawing, chart AND its now-provably-unreferenced image', async () => {
      const original = await buildWorkbookWithOwnedParts()
      const doc = deleteSheet(await load(original), 1) // delete "Extra"
      const bytes = (await writeWorkbookThroughOriginal(original, doc))!
      expect(bytes).not.toBeNull()

      const zip = await JSZip.loadAsync(bytes)
      expect(zip.file('xl/worksheets/sheet2.xml')).toBeNull()
      expect(zip.file('xl/tables/table1.xml')).toBeNull()
      expect(zip.file('xl/comments1.xml')).toBeNull()
      expect(zip.file('xl/drawings/vmlDrawing1.vml')).toBeNull()
      expect(zip.file('xl/drawings/drawing1.xml')).toBeNull()
      expect(zip.file('xl/drawings/_rels/drawing1.xml.rels')).toBeNull()
      expect(zip.file('xl/charts/chart1.xml')).toBeNull()
      // Nothing left in the package references it any more (`Extra`'s own
      // drawing was its only reference) — reference-counting across the
      // WHOLE package proves it, so it is swept too (`sweepOrphanedMedia`).
      expect(zip.file('xl/media/image1.png')).toBeNull()

      const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
      expect(contentTypes).not.toContain('table1.xml')
      expect(contentTypes).not.toContain('comments1.xml')
      expect(contentTypes).not.toContain('drawing1.xml')
      expect(contentTypes).not.toContain('chart1.xml')

      const reopened = await load(bytes.buffer as ArrayBuffer)
      expect(reopened.sheets.map((s) => s.name)).toEqual(['Main'])
    })

    it('leaves an image alone when another surviving sheet\'s drawing still targets it', async () => {
      const original = await buildWorkbookWithSharedImage()
      const doc = deleteSheet(await load(original), 1) // delete "Extra" — "Main" keeps its own drawing on the same image
      const bytes = (await writeWorkbookThroughOriginal(original, doc))!
      expect(bytes).not.toBeNull()

      const zip = await JSZip.loadAsync(bytes)
      expect(zip.file('xl/drawings/drawing1.xml')).toBeNull() // Extra's own drawing is still swept
      expect(zip.file('xl/drawings/drawing2.xml')).not.toBeNull() // Main's own drawing survives (Main was not deleted)
      expect(zip.file('xl/media/image1.png')).not.toBeNull() // still referenced by Main's drawing
    })
  })

  // ---------------------------------------------------------------------
  // sharedStrings.xml hygiene: unused entries dropped, count/uniqueCount
  // recomputed (USR-17 follow-up).
  // ---------------------------------------------------------------------

  describe('writeWorkbookThroughOriginal — sharedStrings.xml hygiene', () => {
    it('drops an entry no cell references, compacting indices and recomputing count/uniqueCount', async () => {
      const original = await buildSharedStringsWorkbook()
      const doc = await load(original) // no edits at all — still exercises the compaction pass on every save
      const bytes = (await writeWorkbookThroughOriginal(original, doc))!
      expect(bytes).not.toBeNull()

      const zip = await JSZip.loadAsync(bytes)
      const sharedStrings = await zip.file('xl/sharedStrings.xml')!.async('string')
      expect(sharedStrings).not.toContain('Beta')
      expect(sharedStrings).toContain('<si><t>Alpha</t></si><si><t>Gamma</t></si>')
      expect(sharedStrings).toContain('count="2"')
      expect(sharedStrings).toContain('uniqueCount="2"')

      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>') // unchanged — already the lowest index
      expect(sheet).toContain('<c r="A2" t="s"><v>1</v></c>') // compacted from 2 down to 1

      const reopened = await load(bytes.buffer as ArrayBuffer)
      expect(reopened.sheets[0].rows[0][0]).toBe('Alpha')
      expect(reopened.sheets[0].rows[1][0]).toBe('Gamma')
    })
  })

describe('writeWorkbookThroughOriginal — structural sheet changes (column ranges)', () => {
  let multi: ArrayBuffer
  beforeEach(async () => {
    multi = await buildMultiSheetWorkbook()
  })

  it('re-anchors through a column insert/delete', async () => {
    let doc = await load(multi)
    doc = insertColumnAt(doc, 0, 0) // insert a column before A — everything shifts right by one
    let bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    let zip = await JSZip.loadAsync(bytes)
    let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<autoFilter ref="B1:D3"/>')
    expect(sheet).toContain('sqref="D2:D3"') // conditional formatting (was C2:C3)

    doc = deleteColumnAt(doc, 0, 0) // delete it straight back — ranges return to their original columns
    bytes = (await writeWorkbookThroughOriginal(multi, doc))!
    zip = await JSZip.loadAsync(bytes)
    sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<autoFilter ref="A1:C3"/>')
  })
})

// ---------------------------------------------------------------------------
// SHEET-3/4/5 — Excel-parity auto-typing on write (`buildCell`). Reuses
// `buildStyledWorkbook`'s A2 (no style — General), B2 (`s="3"`, numFmtId 164
// `dd/mm/yyyy` — an existing DATE format) and C2 (`s="4"`, numFmtId 44 — an
// existing, non-date/percent/General format) to cover both "General cell
// gets a fresh format" and "an already-formatted cell keeps its own format".
// ---------------------------------------------------------------------------

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`

/** A single-cell (A1) workbook whose only style is `@` (text format, builtin numFmtId 49) — SHEET-3's "already text-formatted cell" case, which `buildStyledWorkbook` has no equivalent of. */
async function buildTextFormattedWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file('xl/_rels/workbook.xml.rels', rels(`<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`))
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`,
  )
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('writeWorkbookThroughOriginal — SHEET-3/4/5 auto-typing', () => {
  it('still types a plain leading-zero string as the number 7 (Excel default, unchanged)', async () => {
    const doc = setCellValue(await load(original), 0, 1, 0, '007') // A2 — no style, General
    const { sheet } = await saved(original, doc)
    expect(sheet).toContain('<c r="A2"><v>7</v></c>')
  })

  it('forces a leading-apostrophe value to text, without storing the apostrophe', async () => {
    const doc = setCellValue(await load(original), 0, 1, 0, "'007")
    const { sheet } = await saved(original, doc)
    expect(sheet).toContain('<t xml:space="preserve">007</t>')
    expect(sheet).not.toContain("'007")
  })

  it('forces ANY value to text in a cell already styled `@` (text format), apostrophe never needed', async () => {
    const textWb = await buildTextFormattedWorkbook()
    const doc = setCellValue(await load(textWb), 0, 0, 0, '007')
    const bytes = (await writeWorkbookThroughOriginal(textWb, doc))!
    const zip = await JSZip.loadAsync(bytes)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<c r="A1" s="1" t="inlineStr">')
    expect(sheet).toContain('<t xml:space="preserve">007</t>')
  })

  /** The `<xf>` entries inside `<cellXfs>` ONLY (a cell's `s` index refers into this list, not `<cellStyleXfs>`, which sits earlier in the document and would otherwise shift every index off by one). */
  function cellXfsOf(stylesXml: string): string[] {
    const block = /<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/.exec(stylesXml)
    expect(block).not.toBeNull()
    return Array.from(block![1].matchAll(/<xf [^>]*\/>/g)).map((m) => m[0])
  }

  it('types "50%" as 0.5 with a fresh percent format when the cell was General', async () => {
    const doc = setCellValue(await load(original), 0, 1, 0, '50%') // A2 — General
    const { zip, sheet } = await saved(original, doc)
    const cellMatch = /<c r="A2"( s="(\d+)")?><v>0\.5<\/v><\/c>/.exec(sheet)
    expect(cellMatch).not.toBeNull()
    const styleIndex = cellMatch![2]
    expect(styleIndex).toBeDefined()
    const styles = await zip.file('xl/styles.xml')!.async('string')
    // Builtin `0%` (numFmtId 9) needs no new `<numFmt>` entry at all.
    expect(cellXfsOf(styles)[Number(styleIndex)]).toContain('numFmtId="9"')
  })

  it('types "12.5%" with the two-decimal builtin percent format (0.00%)', async () => {
    const doc = setCellValue(await load(original), 0, 1, 0, '12.5%')
    const { zip, sheet } = await saved(original, doc)
    const cellMatch = /<c r="A2"( s="(\d+)")?><v>0\.125<\/v><\/c>/.exec(sheet)
    expect(cellMatch).not.toBeNull()
    const styles = await zip.file('xl/styles.xml')!.async('string')
    expect(cellXfsOf(styles)[Number(cellMatch![2])]).toContain('numFmtId="10"')
  })

  it('types an ISO date as an Excel serial with a fresh yyyy-mm-dd format when the cell was General', async () => {
    const doc = setCellValue(await load(original), 0, 1, 0, '2024-03-14') // A2 — General
    const { zip, sheet } = await saved(original, doc)
    const cellMatch = /<c r="A2" s="(\d+)"><v>45365<\/v><\/c>/.exec(sheet) // 45365 is 2024-03-14's real Excel serial
    expect(cellMatch).not.toBeNull()
    const styles = await zip.file('xl/styles.xml')!.async('string')
    expect(cellXfsOf(styles)[Number(cellMatch![1])]).toContain('numFmtId="165"')
    // STYLES_XML already uses numFmtId 164 for the pre-existing "dd/mm/yyyy"
    // date format — the new one must pick a genuinely free id (165) rather
    // than collide with it, and the original entry must stay untouched.
    expect(styles).toContain('<numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>')
    expect(styles).toContain('<numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>')
  })

  it('respects an existing DATE format instead of assigning a new one', async () => {
    // B2 already carries s="3" -> numFmtId 164 "dd/mm/yyyy" in STYLES_XML.
    const doc = setCellValue(await load(original), 0, 1, 1, '2024-03-14')
    const { zip, sheet } = await saved(original, doc)
    expect(sheet).toContain('<c r="B2" s="3"><v>45365</v></c>')
    // No new numFmt/cellXfs entries were added for this save.
    expect(await zip.file('xl/styles.xml')!.async('string')).toBe(STYLES_XML)
  })

  it('respects an existing non-percent/date format when a percent-looking value is typed', async () => {
    // C2 already carries s="4" -> numFmtId 44 (accounting), neither General nor percent/date.
    const doc = setCellValue(await load(original), 0, 1, 2, '50%')
    const { zip, sheet } = await saved(original, doc)
    expect(sheet).toContain('<c r="C2" s="4"><v>0.5</v></c>')
    expect(await zip.file('xl/styles.xml')!.async('string')).toBe(STYLES_XML)
  })

  it('leaves an unparseable/ambiguous value as plain text (not 50%, not a date)', async () => {
    let doc = setCellValue(await load(original), 0, 1, 0, '150%%')
    let result = await saved(original, doc)
    expect(result.sheet).toContain('<t xml:space="preserve">150%%</t>')

    doc = setCellValue(await load(original), 0, 1, 0, '2024-13-40') // not a real calendar date
    result = await saved(original, doc)
    expect(result.sheet).toContain('<t xml:space="preserve">2024-13-40</t>')
  })

  it('does not reinterpret a FORMULA result as a percent/date (only user-typed values)', async () => {
    // A formula computing to the text "50%" must stay a plain string cell, not a re-parsed percent.
    let doc = await load(original)
    doc = setCellValue(doc, 0, 1, 0, '="50%"')
    const { sheet } = await saved(original, doc)
    expect(sheet).toContain('<f>"50%"</f>')
    expect(sheet).toMatch(/<c r="A2"[^>]*t="str"[^>]*>[\s\S]*?<v>50%<\/v>/)
  })
})

// ---------------------------------------------------------------------------
// SHEET-9 — `ignoredErrors`' `sqref` must shift through a row/column
// insert/delete exactly like `conditionalFormatting`/`dataValidation` do.
// ---------------------------------------------------------------------------

async function buildWorkbookWithIgnoredErrors(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await buildMultiSheetWorkbook())
  const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet1.replace(
      '</worksheet>',
      '<ignoredErrors><ignoredError numberStoredAsText="1" sqref="B2:B3"/></ignoredErrors></worksheet>',
    ),
  )
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('writeWorkbookThroughOriginal — SHEET-9 ignoredErrors sqref', () => {
  it('shifts ignoredErrors sqref through a row insert, like conditionalFormatting/dataValidation', async () => {
    const withIgnored = await buildWorkbookWithIgnoredErrors()
    let doc = await load(withIgnored)
    doc = insertRowAt(doc, 0, 0) // insert above row 1 — everything shifts down by one
    const bytes = (await writeWorkbookThroughOriginal(withIgnored, doc))!
    const zip = await JSZip.loadAsync(bytes)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('sqref="B3:B4"') // was B2:B3
    expect(sheet).toContain('sqref="C3:C4"') // conditionalFormatting shifted the same way (was C2:C3, a ROW insert)
  })

  it('drops ignoredErrors entirely when its whole range is deleted', async () => {
    const withIgnored = await buildWorkbookWithIgnoredErrors()
    let doc = await load(withIgnored)
    doc = deleteRowAt(doc, 0, 1) // deletes row 2 and row 3 in turn -- consumes B2:B3 entirely
    doc = deleteRowAt(doc, 0, 1)
    const bytes = (await writeWorkbookThroughOriginal(withIgnored, doc))!
    const zip = await JSZip.loadAsync(bytes)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).not.toContain('ignoredError')
  })
})
