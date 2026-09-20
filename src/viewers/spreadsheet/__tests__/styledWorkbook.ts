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
    // `Default Extension="png"` matters here, not just for realism: a real
    // Excel-produced package always declares it once any part uses that
    // extension (OPC requires every part to resolve to SOME content type,
    // §10.1.2.2.1) — omitting it left `xl/media/image1.png` below with none
    // at all, a genuine (if minor) spec violation in this fixture that
    // `scripts/validate-office-file.mjs` catches.
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>` +
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

// ---------------------------------------------------------------------------
// Cross-sheet formulas — a formula referencing ANOTHER sheet (qualified) and
// one referencing its own sheet (unqualified), each both as a cell the user
// never touched (needs full re-anchoring on save) and, in tests, as one the
// user retypes fresh (needs only its sheet name fixed, never its
// coordinates) — see `formulaRefs.ts`.
// ---------------------------------------------------------------------------

export const CROSS_SHEET_BUDGET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:D3"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Item</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Flag</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Cost</t></is></c>
      <c r="D1" t="inlineStr"><is><t>Ref</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Paper</t></is></c>
      <c r="B2" t="inlineStr"><is><t>y</t></is></c>
      <c r="C2"><v>12.5</v></c>
      <c r="D2" t="str"><f>Notes!A1</f><v>See Budget</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Ink</t></is></c>
      <c r="B3" t="inlineStr"><is><t>n</t></is></c>
      <c r="C3"><v>25</v></c>
      <c r="D3" t="str"><f>B2</f><v>y</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:C3"/>
  <conditionalFormatting sqref="C2:C3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>20</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="B2:B3"><formula1>"y,n"</formula1></dataValidation></dataValidations>
  <hyperlinks><hyperlink ref="A2" location="Notes!A1" display="Notes!A1"/></hyperlinks>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`

export const CROSS_SHEET_NOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:A2"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>See Budget</t></is></c></row>
    <row r="2"><c r="A2"><f>Budget!C2</f><v>12.5</v></c></row>
  </sheetData>
</worksheet>`

/** `buildMultiSheetWorkbook`'s two sheets, with a cross-sheet and an own-sheet formula added to "Budget" and "Notes" respectively. */
export async function buildCrossSheetFormulaWorkbook(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await buildMultiSheetWorkbook())
  zip.file('xl/worksheets/sheet1.xml', CROSS_SHEET_BUDGET_XML)
  zip.file('xl/worksheets/sheet2.xml', CROSS_SHEET_NOTES_XML)
  return zip.generateAsync({ type: 'arraybuffer' })
}

// ---------------------------------------------------------------------------
// A second sheet ("Extra") owning a table, comments (+ VML note-shape
// drawing) and a drawing embedding a chart and an image — every part kind
// `removeSheetOwnedParts` in `xlsxPassthrough.ts` sweeps (or, for the image,
// deliberately leaves orphaned) when that sheet is deleted.
// ---------------------------------------------------------------------------

export async function buildWorkbookWithOwnedParts(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>` +
      `<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>` +
      `<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>` +
      `<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` +
      `<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}">` +
      `<sheets><sheet name="Main" sheetId="1" r:id="rId1"/><sheet name="Extra" sheetId="2" r:id="rId2"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels(
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>`,
    ),
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Main</t></is></c></row></sheetData></worksheet>`,
  )
  zip.file(
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}">` +
      `<dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
      `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c></row></sheetData>` +
      `<tableParts count="1"><tablePart r:id="rId10"/></tableParts>` +
      `<legacyDrawing r:id="rId12"/><drawing r:id="rId13"/></worksheet>`,
  )
  zip.file(
    'xl/worksheets/_rels/sheet2.xml.rels',
    rels(
      `<Relationship Id="rId10" Type="${REL}/table" Target="../tables/table1.xml"/>` +
        `<Relationship Id="rId11" Type="${REL}/comments" Target="../comments1.xml"/>` +
        `<Relationship Id="rId12" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` +
        `<Relationship Id="rId13" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/>`,
    ),
  )
  zip.file(
    'xl/tables/table1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:A1"><tableColumns count="1"><tableColumn id="1" name="Column1"/></tableColumns></table>`,
  )
  zip.file(
    'xl/comments1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Atlas</author></authors><commentList><comment ref="A1" authorId="0"><text><t>Note</t></text></comment></commentList></comments>`,
  )
  zip.file('xl/drawings/vmlDrawing1.vml', '<xml>placeholder vml note shape</xml>')
  zip.file(
    'xl/drawings/drawing1.xml',
    rels(`<Relationship Id="rId1" Type="${REL}/chart" Target="../charts/chart1.xml"/><Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.png"/>`),
  )
  zip.file(
    'xl/drawings/_rels/drawing1.xml.rels',
    rels(
      `<Relationship Id="rId1" Type="${REL}/chart" Target="../charts/chart1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.png"/>`,
    ),
  )
  zip.file('xl/charts/chart1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><chartSpace xmlns="placeholder"/>`)
  zip.file('xl/media/image1.png', new Uint8Array([137, 80, 78, 71]))
  return zip.generateAsync({ type: 'arraybuffer' })
}

/**
 * Same as `buildWorkbookWithOwnedParts`, except "Main" (the SURVIVING sheet)
 * has its own drawing that ALSO targets `xl/media/image1.png` — the image is
 * genuinely shared across two sheets' drawings, not just shaped like it
 * might be. Deleting "Extra" must leave it alone.
 */
export async function buildWorkbookWithSharedImage(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await buildWorkbookWithOwnedParts())

  const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    contentTypes.replace(
      '</Types>',
      '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>',
    ),
  )

  const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', sheet1.replace('</worksheet>', `<drawing xmlns:r="${REL}" r:id="rId1"/></worksheet>`))
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels(`<Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing2.xml"/>`))
  zip.file('xl/drawings/drawing2.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="placeholder"/>`)
  zip.file(
    'xl/drawings/_rels/drawing2.xml.rels',
    rels(`<Relationship Id="rId1" Type="${REL}/image" Target="../media/image1.png"/>`),
  )

  return zip.generateAsync({ type: 'arraybuffer' })
}

// ---------------------------------------------------------------------------
// A single-sheet workbook whose `xl/sharedStrings.xml` has one entry
// ("Beta") no cell references at all, exercising `compactSharedStrings`.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// A shared-formula group (`<f t="shared" ref="B2:B4" si="0">A2*2</f>` on the
// master cell B2, `<f t="shared" si="0"/>` follower cells on B3/B4) —
// exercises re-anchoring the group's `ref` SPAN attribute through a
// row insert/delete, not just the master's own formula TEXT.
// ---------------------------------------------------------------------------

export const SHARED_FORMULA_SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B4"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Base</t></is></c><c r="B1" t="inlineStr"><is><t>Doubled</t></is></c></row>
    <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B4" si="0">A2*2</f><v>4</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>
    <row r="4"><c r="A4"><v>4</v></c><c r="B4"><f t="shared" si="0"/><v>8</v></c></row>
  </sheetData>
</worksheet>`

export async function buildSharedFormulaWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file('xl/_rels/workbook.xml.rels', rels(`<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`))
  zip.file('xl/worksheets/sheet1.xml', SHARED_FORMULA_SHEET_XML)
  return zip.generateAsync({ type: 'arraybuffer' })
}

export async function buildSharedStringsWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels(`<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`),
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
      `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c></row></sheetData></worksheet>`,
  )
  zip.file(
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="3">` +
      `<si><t>Alpha</t></si><si><t>Beta</t></si><si><t>Gamma</t></si></sst>`,
  )
  return zip.generateAsync({ type: 'arraybuffer' })
}

