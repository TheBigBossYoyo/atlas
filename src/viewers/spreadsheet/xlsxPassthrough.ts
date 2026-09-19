/**
 * USR-17 — saving an .xlsx/.xlsm THROUGH the file it was opened from.
 *
 * `spreadsheetWrite.ts` builds a brand-new workbook from the editable model,
 * which is the only option for a CSV or a format change — but it can only
 * carry what the model itself holds (values, formulas, merges, widths), so
 * every cell style, number format, conditional format, chart, image, filter
 * and pivot in a real workbook was lost on save.
 *
 * This module instead keeps the ORIGINAL package and rewrites only each
 * worksheet's `<sheetData>`:
 *
 *  - a cell the user never touched is copied element-for-element (its type,
 *    style, cached formula value and even rich text stay byte-identical);
 *  - an edited cell is written from the model, keeping the style (`s`) of the
 *    cell that was there before;
 *  - rows and columns carry their original attributes (height, custom format,
 *    hidden) through inserts and deletes via the model's `rowSources`/
 *    `colSources`, so styling follows the row it belonged to;
 *  - `calcChain.xml` is dropped and the workbook is marked "recalculate on
 *    load", since cached formula values are only as good as this app's own
 *    formula engine;
 *  - Excel tables are rewritten in place (range, column names).
 *
 * Everything else in the package is passed through untouched.
 *
 * Deliberate limits: structural sheet changes (adding or deleting a sheet)
 * and any non-OOXML target fall back to the fresh-workbook writer, and
 * references that live outside the model — conditional formatting, data
 * validation, defined names — are not re-anchored when rows or columns are
 * inserted or deleted.
 */
import JSZip from 'jszip'

import {
  childElements,
  descendantElements,
  firstChildElement,
  parseXmlPart,
  serializeXmlPart,
  xmlSafeText,
} from '../../office/ooxmlDom'
import { readSheetParts } from './spreadsheetPanes'
import { cellKey, type EditableSheet, type SpreadsheetDocument } from './spreadsheetDocument'
import { encodeCol, rewriteTableXml, tableHeaderNames } from './spreadsheetTables'

/** Worksheet children that must come AFTER `<mergeCells>` (CT_Worksheet order). */
const AFTER_MERGE_CELLS: ReadonlySet<string> = new Set([
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
])

function encodeCell(row: number, col: number): string {
  return `${encodeCol(col)}${row + 1}`
}

function isNumericText(text: string): boolean {
  const trimmed = text.trim()
  return trimmed !== '' && Number.isFinite(Number(trimmed))
}

type OriginalSheet = {
  readonly doc: XMLDocument
  readonly root: Element
  readonly sheetData: Element
  readonly rows: Map<number, Element>
  readonly cells: Map<string, Element>
}

function readOriginalSheet(xml: string): OriginalSheet | null {
  const doc = parseXmlPart(xml)
  const root = doc.documentElement
  const sheetData = firstChildElement(root, 'sheetData')
  if (!sheetData) return null

  const rows = new Map<number, Element>()
  const cells = new Map<string, Element>()
  for (const rowElement of childElements(sheetData, 'row')) {
    const rowIndex = Number(rowElement.getAttribute('r')) - 1
    if (!Number.isFinite(rowIndex) || rowIndex < 0) continue
    rows.set(rowIndex, rowElement)
    for (const cell of childElements(rowElement, 'c')) {
      const reference = cell.getAttribute('r')
      const match = reference ? /^([A-Z]+)(\d+)$/i.exec(reference) : null
      if (!match) continue
      let col = 0
      for (const character of match[1].toUpperCase()) col = col * 26 + (character.charCodeAt(0) - 64)
      cells.set(cellKey(Number(match[2]) - 1, col - 1), cell)
    }
  }
  return { doc, root, sheetData, rows, cells }
}

/** Builds the `<c>` element for a cell that was edited (or is new). */
function buildCell(
  doc: XMLDocument,
  namespace: string | null,
  address: string,
  text: string,
  formula: string | undefined,
  style: string | null,
): Element {
  const cell = doc.createElementNS(namespace, 'c')
  cell.setAttribute('r', address)
  if (style !== null) cell.setAttribute('s', style)

  if (formula !== undefined) {
    const formulaElement = doc.createElementNS(namespace, 'f')
    formulaElement.textContent = xmlSafeText(formula)
    cell.appendChild(formulaElement)
  }

  if (text === '' && formula === undefined) return cell

  if (isNumericText(text)) {
    const value = doc.createElementNS(namespace, 'v')
    value.textContent = String(Number(text.trim()))
    cell.appendChild(value)
    return cell
  }

  if (formula !== undefined) {
    cell.setAttribute('t', 'str')
    const value = doc.createElementNS(namespace, 'v')
    value.textContent = xmlSafeText(text)
    cell.appendChild(value)
    return cell
  }

  // Inline strings avoid touching sharedStrings.xml (Excel reads both).
  cell.setAttribute('t', 'inlineStr')
  const inlineString = doc.createElementNS(namespace, 'is')
  const textElement = doc.createElementNS(namespace, 't')
  textElement.setAttribute('xml:space', 'preserve')
  textElement.textContent = xmlSafeText(text)
  inlineString.appendChild(textElement)
  cell.appendChild(inlineString)
  return cell
}

function rebuildSheetData(sheet: EditableSheet, original: OriginalSheet): void {
  const { doc, sheetData } = original
  const namespace = sheetData.namespaceURI
  const fresh = doc.createElementNS(namespace, 'sheetData')
  const edited = sheet.editedCells

  for (let row = 0; row < sheet.rows.length; row++) {
    const sourceRow = sheet.rowSources?.[row] ?? null
    const template = sourceRow === null ? undefined : original.rows.get(sourceRow)
    const rowElement = doc.createElementNS(namespace, 'row')
    rowElement.setAttribute('r', String(row + 1))
    if (template) {
      for (const attribute of Array.from(template.attributes)) {
        // `spans` is an optimization hint; a stale one is worse than none.
        if (attribute.name !== 'r' && attribute.name !== 'spans') {
          rowElement.setAttribute(attribute.name, attribute.value)
        }
      }
    }

    for (let col = 0; col < sheet.colCount; col++) {
      const sourceCol = sheet.colSources?.[col] ?? null
      const originalCell =
        sourceRow === null || sourceCol === null ? undefined : original.cells.get(cellKey(sourceRow, sourceCol))
      const text = sheet.rows[row]?.[col] ?? ''
      const formula = sheet.formulas[row]?.[col]
      const address = encodeCell(row, col)

      // Untouched cells are copied verbatim — that is what keeps dates dates,
      // percentages percentages, and rich text rich.
      if (originalCell && edited !== undefined && !edited.has(cellKey(row, col))) {
        const clone = originalCell.cloneNode(true) as Element
        clone.setAttribute('r', address)
        rowElement.appendChild(clone)
        continue
      }

      const style = originalCell?.getAttribute('s') ?? null
      if (text === '' && formula === undefined) {
        // Keep an empty-but-styled cell so its formatting survives.
        if (style !== null) rowElement.appendChild(buildCell(doc, namespace, address, '', undefined, style))
        continue
      }
      rowElement.appendChild(buildCell(doc, namespace, address, text, formula, style))
    }

    if (rowElement.children.length > 0 || template) fresh.appendChild(rowElement)
  }

  sheetData.parentNode?.replaceChild(fresh, sheetData)
}

function updateDimension(root: Element, sheet: EditableSheet): void {
  const dimension = firstChildElement(root, 'dimension')
  if (!dimension) return
  dimension.setAttribute('ref', `A1:${encodeCol(Math.max(sheet.colCount, 1) - 1)}${Math.max(sheet.rows.length, 1)}`)
}

function updateMergeCells(doc: XMLDocument, root: Element, sheet: EditableSheet): void {
  const existing = firstChildElement(root, 'mergeCells')
  if (sheet.merges.length === 0) {
    if (existing) root.removeChild(existing)
    return
  }

  const element = doc.createElementNS(root.namespaceURI, 'mergeCells')
  element.setAttribute('count', String(sheet.merges.length))
  for (const merge of sheet.merges) {
    const mergeCell = doc.createElementNS(root.namespaceURI, 'mergeCell')
    mergeCell.setAttribute('ref', `${encodeCell(merge.r0, merge.c0)}:${encodeCell(merge.r1, merge.c1)}`)
    element.appendChild(mergeCell)
  }

  if (existing) {
    root.replaceChild(element, existing)
    return
  }
  const anchor = Array.from(root.children).find((child) => AFTER_MERGE_CELLS.has(child.localName))
  root.insertBefore(element, anchor ?? null)
}

/** Re-anchors `<cols>` entries when columns were inserted or deleted. */
function updateColumns(doc: XMLDocument, root: Element, sheet: EditableSheet): void {
  const cols = firstChildElement(root, 'cols')
  const sources = sheet.colSources
  if (!cols || !sources || sources.every((source, index) => source === index)) return

  const definitions = childElements(cols, 'col')
  const fresh = doc.createElementNS(root.namespaceURI, 'cols')
  sources.forEach((source, index) => {
    if (source === null) return
    const definition = definitions.find(
      (col) => Number(col.getAttribute('min')) - 1 <= source && source <= Number(col.getAttribute('max')) - 1,
    )
    if (!definition) return
    const clone = definition.cloneNode(false) as Element
    clone.setAttribute('min', String(index + 1))
    clone.setAttribute('max', String(index + 1))
    fresh.appendChild(clone)
  })

  if (fresh.children.length === 0) root.removeChild(cols)
  else root.replaceChild(fresh, cols)
}

function markRecalculateOnLoad(workbookXml: string): string {
  const doc = parseXmlPart(workbookXml)
  const root = doc.documentElement
  let calcPr = firstChildElement(root, 'calcPr')
  if (!calcPr) {
    calcPr = doc.createElementNS(root.namespaceURI, 'calcPr')
    root.insertBefore(calcPr, firstChildElement(root, 'extLst'))
  }
  calcPr.setAttribute('fullCalcOnLoad', '1')
  return serializeXmlPart(doc)
}

function renameSheets(workbookXml: string, names: ReadonlyArray<string>): string {
  const doc = parseXmlPart(workbookXml)
  const sheets = descendantElements(doc, 'sheet').filter((sheet) => sheet.parentElement?.localName === 'sheets')
  let changed = false
  sheets.forEach((sheet, index) => {
    const name = names[index]
    if (name !== undefined && sheet.getAttribute('name') !== name) {
      sheet.setAttribute('name', name)
      changed = true
    }
  })
  return changed ? serializeXmlPart(doc) : workbookXml
}

/** Drops the (now stale) calculation chain; Excel rebuilds it on the next recalculation. */
function removeCalcChain(zip: JSZip, contentTypes: string | undefined): string | undefined {
  if (!zip.file('xl/calcChain.xml')) return contentTypes
  zip.remove('xl/calcChain.xml')
  if (contentTypes === undefined) return undefined

  const doc = parseXmlPart(contentTypes)
  for (const override of descendantElements(doc, 'Override')) {
    if (override.getAttribute('PartName') === '/xl/calcChain.xml') override.parentNode?.removeChild(override)
  }
  return serializeXmlPart(doc)
}

/**
 * Rewrites the original workbook with the document's current values.
 * Returns `null` when the original is not an OOXML workbook this can patch,
 * or when sheets were added/removed, so the caller falls back to
 * `writeWorkbookBytes`.
 */
export async function writeWorkbookThroughOriginal(
  original: ArrayBuffer,
  document: SpreadsheetDocument,
): Promise<Uint8Array | null> {
  const sourcePaths = document.sheets.map((sheet) => sheet.sourcePath)
  if (sourcePaths.some((path) => path === undefined)) return null

  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(original)
  } catch {
    return null
  }

  const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
  if (!workbookXml) return null
  const relsXml = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? null
  const parts = readSheetParts(workbookXml, relsXml)
  // A sheet added or deleted here changes the workbook's structure; that still
  // goes through the fresh-workbook writer.
  if (parts.length !== document.sheets.length) return null
  if (sourcePaths.some((path, index) => path !== parts[index].path)) return null

  for (const [sheetIndex, sheet] of document.sheets.entries()) {
    const partPath = sourcePaths[sheetIndex]!
    const sheetXml = await zip.file(partPath)?.async('string')
    if (!sheetXml) return null
    const parsed = readOriginalSheet(sheetXml)
    if (!parsed) return null

    rebuildSheetData(sheet, parsed)
    updateDimension(parsed.root, sheet)
    updateMergeCells(parsed.doc, parsed.root, sheet)
    updateColumns(parsed.doc, parsed.root, sheet)
    zip.file(partPath, serializeXmlPart(parsed.doc))

    for (const [tableIndex, table] of (sheet.tables ?? []).entries()) {
      if (table.partPath === undefined) continue
      zip.file(table.partPath, rewriteTableXml(table, tableIndex + 1, tableHeaderNames(table, sheet.rows)))
    }
  }

  const contentTypes = removeCalcChain(zip, await zip.file('[Content_Types].xml')?.async('string'))
  if (contentTypes !== undefined) zip.file('[Content_Types].xml', contentTypes)
  zip.file(
    'xl/workbook.xml',
    markRecalculateOnLoad(renameSheets(workbookXml, document.sheets.map((sheet) => sheet.name))),
  )

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
