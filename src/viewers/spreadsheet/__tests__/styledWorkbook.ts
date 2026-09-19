/**
 * USR-17 — a styled single-sheet workbook (bold header, date and currency
 * number formats, a formula with a cached value, conditional formatting,
 * column widths, a defined name, media and a stale calc chain) used by the
 * save-through-original tests, unit and e2e alike.
 */
import JSZip from 'jszip'

export const SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:C3"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="3" width="12"/></cols>
  <sheetData>
    <row r="1" ht="30" customHeight="1" s="2" customFormat="1">
      <c r="A1" s="2" t="inlineStr"><is><t>Item</t></is></c>
      <c r="B1" s="2" t="inlineStr"><is><t>Due</t></is></c>
      <c r="C1" s="2" t="inlineStr"><is><t>Cost</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Paper</t></is></c>
      <c r="B2" s="3"><v>45000</v></c>
      <c r="C2" s="4"><v>12.5</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Ink</t></is></c>
      <c r="B3" s="3"><v>45010</v></c>
      <c r="C3" s="4"><f>C2*2</f><v>25</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="C2:C3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>20</formula></cfRule></conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`

export const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="44" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`

const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** A styled single-sheet workbook with a formula, a date format, conditional formatting and a stale calc chain. */
export async function buildStyledWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Costs">Budget!$C$2:$C$3</definedName></definedNames></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels(
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${REL}/calcChain" Target="calcChain.xml"/>`,
    ),
  )
  zip.file('xl/worksheets/sheet1.xml', SHEET_XML)
  zip.file('xl/styles.xml', STYLES_XML)
  zip.file(
    'xl/calcChain.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C3" i="1"/></calcChain>`,
  )
  zip.file('xl/media/image1.png', new Uint8Array([137, 80, 78, 71]))
  return zip.generateAsync({ type: 'arraybuffer' })
}

// ---------------------------------------------------------------------------
// Two-sheet workbook — structural changes (add/delete/rename/reorder sheets)
// and range re-anchoring (conditional formatting, data validation, a
// hyperlink and the sheet-level autoFilter, all through a row insert/delete).
// ---------------------------------------------------------------------------

export const MULTI_SHEET1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:C3"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Item</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Flag</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Cost</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Paper</t></is></c>
      <c r="B2" t="inlineStr"><is><t>y</t></is></c>
      <c r="C2"><v>12.5</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Ink</t></is></c>
      <c r="B3" t="inlineStr"><is><t>n</t></is></c>
      <c r="C3"><v>25</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:C3"/>
  <conditionalFormatting sqref="C2:C3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>20</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="B2:B3"><formula1>"y,n"</formula1></dataValidation></dataValidations>
  <hyperlinks><hyperlink ref="A2" location="Notes!A1" display="Notes!A1"/></hyperlinks>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`

export const MULTI_SHEET2_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:A1"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>See Budget</t></is></c></row></sheetData>
</worksheet>`

/**
 * A two-sheet workbook ("Budget", "Notes") with a sheet-level autoFilter, a
 * conditional format, a data validation and a hyperlink on "Budget" (all
 * `sqref`/`ref`-anchored ranges the passthrough must shift on a row/column
 * insert or delete), plus three defined names covering every
 * structural-change case: a workbook-scope single-area name into "Budget"
 * (rename/re-anchor), a name local to "Budget" (`localSheetId="0"`, dropped
 * when "Budget" is deleted), and a name local to "Notes"
 * (`localSheetId="1"`, renumbered to `0` when "Budget" is deleted ahead of it).
 */
export async function buildMultiSheetWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}">` +
      `<sheets><sheet name="Budget" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId4"/></sheets>` +
      `<definedNames>` +
      `<definedName name="Costs">Budget!$C$2:$C$3</definedName>` +
      `<definedName name="BudgetHeader" localSheetId="0">Budget!$A$1:$C$1</definedName>` +
      `<definedName name="NotesRef" localSheetId="1">Notes!$A$1</definedName>` +
      `</definedNames></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels(
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId4" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>`,
    ),
  )
  zip.file('xl/worksheets/sheet1.xml', MULTI_SHEET1_XML)
  zip.file('xl/worksheets/sheet2.xml', MULTI_SHEET2_XML)
  zip.file('xl/styles.xml', STYLES_XML)
  return zip.generateAsync({ type: 'arraybuffer' })
}

