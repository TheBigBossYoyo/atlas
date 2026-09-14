/**
 * T4/DAT-10 remainder — frozen-pane reading via JSZip + DOMParser, since
 * SheetJS itself never parses `<pane>` (see spreadsheetPanes.ts's header).
 */
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'

import { readFrozenPanes } from '../spreadsheetPanes'

const WORKBOOK_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Frozen" sheetId="1" r:id="rId1"/>
    <sheet name="Plain" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`

function frozenSheetXml(xSplit: number, ySplit: number, state = 'frozen'): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="${xSplit}" ySplit="${ySplit}" topLeftCell="B2" activePane="bottomRight" state="${state}"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>`
}

const PLAIN_SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>`

async function buildFixtureZip(sheet1Xml: string): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file('xl/workbook.xml', WORKBOOK_XML)
  zip.file('xl/_rels/workbook.xml.rels', RELS_XML)
  zip.file('xl/worksheets/sheet1.xml', sheet1Xml)
  zip.file('xl/worksheets/sheet2.xml', PLAIN_SHEET_XML)
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('readFrozenPanes', () => {
  it('reads xSplit/ySplit from a state="frozen" pane, keyed by sheet name', async () => {
    const buffer = await buildFixtureZip(frozenSheetXml(1, 2))

    const result = await readFrozenPanes(buffer)

    expect(result).toEqual({ Frozen: { cols: 1, rows: 2 } })
  })

  it('supports a rows-only freeze (xSplit=0)', async () => {
    const buffer = await buildFixtureZip(frozenSheetXml(0, 3))

    const result = await readFrozenPanes(buffer)

    expect(result).toEqual({ Frozen: { cols: 0, rows: 3 } })
  })

  it('ignores a plain "split" pane (not a freeze)', async () => {
    const buffer = await buildFixtureZip(frozenSheetXml(1, 1, 'split'))

    const result = await readFrozenPanes(buffer)

    expect(result).toEqual({})
  })

  it('resolves an empty map for a sheet with no <pane> at all', async () => {
    const buffer = await buildFixtureZip(PLAIN_SHEET_XML)

    const result = await readFrozenPanes(buffer)

    expect(result).toEqual({})
  })

  it('resolves an empty map (never rejects) for a non-zip buffer', async () => {
    const bytes = new TextEncoder().encode('not a zip at all')
    const result = await readFrozenPanes(bytes.buffer as ArrayBuffer)
    expect(result).toEqual({})
  })

  it('resolves an empty map for a real xlsx that has no freeze panes at all', async () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a', 'b']]), 'Sheet1')
    const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer

    const result = await readFrozenPanes(buffer)

    expect(result).toEqual({})
  })
})
