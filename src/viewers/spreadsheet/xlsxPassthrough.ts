/**
 * USR-17 — saving an .xlsx/.xlsm THROUGH the file it was opened from.
 *
 * `spreadsheetWrite.ts` builds a brand-new workbook from the editable model,
 * which is the only option for a CSV or a format change — but it can only
 * carry what the model itself holds (values, formulas, merges, widths), so
 * every cell style, number format, conditional format, chart, image, filter
 * and pivot in a real workbook was lost on save.
 *
 * This module instead keeps the ORIGINAL package and rewrites only what
 * changed:
 *
 *  - a cell the user never touched is copied element-for-element (its type,
 *    style, cached formula value and even rich text stay byte-identical);
 *  - an edited cell is written from the model, keeping the style (`s`) of the
 *    cell that was there before;
 *  - rows and columns carry their original attributes (height, custom format,
 *    hidden) through inserts and deletes via the model's `rowSources`/
 *    `colSources`, so styling follows the row it belonged to;
 *  - conditional formatting, data validation, hyperlinks and the sheet-level
 *    autoFilter are re-anchored through the same inserts/deletes
 *    (`spreadsheetRangeShift.ts`), the way Excel itself keeps them aligned;
 *  - a sheet added, deleted, renamed or reordered in Atlas is reflected in
 *    `workbook.xml`, its relationships and `[Content_Types].xml`, instead of
 *    falling back to a from-scratch workbook for the WHOLE file just because
 *    one sheet's structure changed;
 *  - defined names are kept pointing at the right cells: a name scoped to a
 *    deleted sheet is dropped and every later sheet's `localSheetId` is
 *    renumbered; a name's own range is re-anchored the same way a
 *    conditional format's is; a rename updates the `Sheet!` prefix in a
 *    single-area name (see `rewriteDefinedNameFormula`'s own header for the
 *    exact — deliberately narrow — pattern this covers);
 *  - `calcChain.xml` is dropped and the workbook is marked "recalculate on
 *    load", since cached formula values are only as good as this app's own
 *    formula engine;
 *  - Excel tables are rewritten in place (range, column names).
 *
 * Everything else in the package is passed through untouched.
 *
 * Deliberate limits (see the findings register for the full list): a
 * non-OOXML target still falls back to the fresh-workbook writer; a cell
 * FORMULA referencing a renamed/reordered/deleted sheet (anywhere other than
 * inside a defined name) is not rewritten, nor is a multi-area or
 * function-wrapped defined name, or a whole-row/whole-column range (e.g. a
 * Print_Titles `$1:$1`); a deleted sheet's own table/drawing parts are
 * orphaned in the zip (unreferenced, but still present) rather than swept.
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
import { cellKey, type EditableSheet, type SpreadsheetDocument } from './spreadsheetDocument'
import { encodeCol, rewriteTableXml, tableHeaderNames } from './spreadsheetTables'
import { remapSqref, type IndexSources } from './spreadsheetRangeShift'
import { loadWorkbookZip } from './spreadsheetZipBudget'

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

const SPREADSHEETML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const WORKSHEET_REL_TYPE = `${RELATIONSHIPS_NS}/worksheet`
const WORKSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const WORKBOOK_RELS_PATH = 'xl/_rels/workbook.xml.rels'
const CONTENT_TYPES_PATH = '[Content_Types].xml'
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

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

/** A brand-new, otherwise-empty worksheet part for a sheet added in Atlas (no original to read from). */
function blankOriginalSheet(): OriginalSheet {
  const xml =
    `${XML_DECLARATION}<worksheet xmlns="${SPREADSHEETML_NS}" xmlns:r="${RELATIONSHIPS_NS}">` +
    '<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>'
  // Every row's `sourceRow`/`sourceCol` is `null` for a sheet with no
  // `rowSources`/`colSources` at all (see `rebuildSheetData`), so every cell
  // is built fresh from the model — this template only needs to be a valid,
  // empty worksheet for that output to attach to.
  return readOriginalSheet(xml)!
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

/**
 * Re-anchors every `sqref`-bearing conditional format / data validation, and
 * every `ref`-bearing hyperlink / sheet-level autoFilter, through this
 * sheet's row/column inserts and deletes (USR-17 finding). An element whose
 * entire area was deleted is dropped outright rather than left pointing at
 * cells that no longer exist.
 */
function updateShiftedRanges(root: Element, sheet: EditableSheet): void {
  const rowSources: IndexSources = sheet.rowSources
  const colSources: IndexSources = sheet.colSources
  if (!rowSources && !colSources) return // no structural changes recorded for this sheet — nothing to re-anchor

  for (const cf of descendantElements(root, 'conditionalFormatting')) {
    const sqref = cf.getAttribute('sqref')
    const next = sqref ? remapSqref(sqref, rowSources, colSources) : null
    if (next) cf.setAttribute('sqref', next)
    else cf.parentNode?.removeChild(cf)
  }

  const dataValidations = firstChildElement(root, 'dataValidations')
  if (dataValidations) {
    let survivors = 0
    for (const validation of childElements(dataValidations, 'dataValidation')) {
      const sqref = validation.getAttribute('sqref')
      const next = sqref ? remapSqref(sqref, rowSources, colSources) : null
      if (next) {
        validation.setAttribute('sqref', next)
        survivors += 1
      } else {
        dataValidations.removeChild(validation)
      }
    }
    if (survivors === 0) dataValidations.parentNode?.removeChild(dataValidations)
    else dataValidations.setAttribute('count', String(survivors))
  }

  const hyperlinks = firstChildElement(root, 'hyperlinks')
  if (hyperlinks) {
    let survivors = 0
    for (const link of childElements(hyperlinks, 'hyperlink')) {
      const ref = link.getAttribute('ref')
      const next = ref ? remapSqref(ref, rowSources, colSources) : null
      // A dropped hyperlink whose Target was an external relationship leaves
      // that relationship unused in the worksheet's own .rels part — Excel
      // tolerates an unreferenced relationship, so this is left as-is rather
      // than also scrubbing xl/worksheets/_rels/sheetN.xml.rels.
      if (next) {
        link.setAttribute('ref', next)
        survivors += 1
      } else {
        hyperlinks.removeChild(link)
      }
    }
    if (survivors === 0) hyperlinks.parentNode?.removeChild(hyperlinks)
  }

  // The sheet-level autoFilter (not one owned by an Excel table — those are
  // re-anchored by `rewriteTableXml` from the live, model-tracked table
  // range instead).
  const autoFilter = firstChildElement(root, 'autoFilter')
  if (autoFilter) {
    const ref = autoFilter.getAttribute('ref')
    const next = ref ? remapSqref(ref, rowSources, colSources) : null
    if (next) autoFilter.setAttribute('ref', next)
    else autoFilter.parentNode?.removeChild(autoFilter)
  }
}

function markRecalculateOnLoad(doc: XMLDocument, root: Element): void {
  let calcPr = firstChildElement(root, 'calcPr')
  if (!calcPr) {
    calcPr = doc.createElementNS(root.namespaceURI, 'calcPr')
    root.insertBefore(calcPr, firstChildElement(root, 'extLst'))
  }
  calcPr.setAttribute('fullCalcOnLoad', '1')
}

/**
 * Drops the (now stale) calculation chain from the zip AND the
 * `xl/_rels/workbook.xml.rels` relationship that points at it — leaving
 * that relationship behind (the original bug here) is a dangling `r:id`
 * target: the part it named no longer exists in the package, which is
 * exactly the "relationship target missing" defect that makes Excel show
 * an "unreadable content" repair prompt. Returns whether a calc chain
 * existed at all, so the caller can also scrub its `[Content_Types].xml`
 * override. Excel rebuilds the chain on the next recalculation.
 */
function removeCalcChain(zip: JSZip, relsDoc: XMLDocument): boolean {
  if (!zip.file('xl/calcChain.xml')) return false
  zip.remove('xl/calcChain.xml')
  for (const rel of descendantElements(relsDoc, 'Relationship')) {
    if (resolveWorkbookRelTarget(rel.getAttribute('Target') ?? '') === 'xl/calcChain.xml') {
      rel.parentNode?.removeChild(rel)
    }
  }
  return true
}

// ---------------------------------------------------------------------------
// Relationship / workbook-structure helpers
// ---------------------------------------------------------------------------

/** Resolves a `xl/_rels/workbook.xml.rels` `Target` (relative to `xl/`) to its full zip-entry path. */
function resolveWorkbookRelTarget(target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  return `xl/${target}`
}

/** Inverse of `resolveWorkbookRelTarget`, for writing a fresh `Target`. */
function relativeToXl(path: string): string {
  return path.startsWith('xl/') ? path.slice(3) : `/${path}`
}

type OriginalSheetEntry = {
  readonly element: Element
  readonly name: string
  readonly rId: string
  readonly path: string | undefined
  readonly originalIndex: number
}

function readOriginalSheetEntries(workbookRoot: Element, relsDoc: XMLDocument | null): OriginalSheetEntry[] {
  const sheetsContainer = firstChildElement(workbookRoot, 'sheets')
  if (!sheetsContainer) return []

  const relTargets = new Map<string, string>()
  if (relsDoc) {
    for (const rel of descendantElements(relsDoc, 'Relationship')) {
      const id = rel.getAttribute('Id')
      const target = rel.getAttribute('Target')
      if (id && target) relTargets.set(id, resolveWorkbookRelTarget(target))
    }
  }

  return childElements(sheetsContainer, 'sheet').map((element, originalIndex) => {
    const rId = element.getAttribute('r:id') ?? element.getAttributeNS(RELATIONSHIPS_NS, 'id') ?? ''
    return { element, name: element.getAttribute('name') ?? '', rId, path: relTargets.get(rId), originalIndex }
  })
}

function nextNumericSuffix(ids: Iterable<string>, prefix: string): number {
  let max = 0
  for (const id of ids) {
    const match = new RegExp(`^${prefix}(\\d+)$`).exec(id)
    if (match) max = Math.max(max, Number(match[1]))
  }
  return max + 1
}

function nextWorksheetPartPath(zip: JSZip): string {
  const existing = new Set(Object.keys(zip.files))
  let n = 1
  while (existing.has(`xl/worksheets/sheet${n}.xml`)) n += 1
  return `xl/worksheets/sheet${n}.xml`
}

/** Adds an `<Override>` for a new part. Order among `<Override>` siblings is not significant to the OOXML schema. */
function addContentTypeOverride(doc: XMLDocument, partPath: string, contentType: string): void {
  const override = doc.createElementNS(doc.documentElement.namespaceURI, 'Override')
  override.setAttribute('PartName', `/${partPath}`)
  override.setAttribute('ContentType', contentType)
  doc.documentElement.appendChild(override)
}

function removeContentTypeOverrideForPart(doc: XMLDocument, partPath: string): void {
  for (const override of descendantElements(doc, 'Override')) {
    if (override.getAttribute('PartName') === `/${partPath}`) override.parentNode?.removeChild(override)
  }
}

// ---------------------------------------------------------------------------
// Defined names
// ---------------------------------------------------------------------------

type SheetChange = {
  /** `undefined` when the sheet still exists but kept its name. */
  readonly newName: string | undefined
  readonly deleted: boolean
  readonly rowSources: IndexSources
  readonly colSources: IndexSources
}

function decodeDollarCell(ref: string): { readonly r: number; readonly c: number; readonly colAbs: boolean; readonly rowAbs: boolean } | null {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref.trim())
  if (!match) return null
  let c = 0
  for (const character of match[2].toUpperCase()) c = c * 26 + (character.charCodeAt(0) - 64)
  return { r: Number(match[4]) - 1, c: c - 1, colAbs: match[1] === '$', rowAbs: match[3] === '$' }
}

function encodeDollarCell(row: number, col: number, colAbs: boolean, rowAbs: boolean): string {
  return `${colAbs ? '$' : ''}${encodeCol(col)}${rowAbs ? '$' : ''}${row + 1}`
}

const SINGLE_AREA_NAME = /^(?:'((?:[^']|'')*)'|([A-Za-z_][\w.]*))!(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$/

/**
 * Rewrites one `<definedName>`'s formula text when — and only when — it is a
 * single-area reference into exactly one sheet (`Sheet1!$A$1:$B$2`, quoted
 * or not): the sheet name is swapped for its new one (rename) and the range
 * is re-anchored through that sheet's row/column inserts and deletes, the
 * same way a conditional format's `sqref` is. Anything else — a multi-area
 * reference, a name wrapping a function, a whole-row/whole-column range like
 * `$1:$1` (no column letter/row digit pair to parse), or a name that does
 * not reference a changed sheet at all — is returned unchanged. Returns
 * `null` when the referenced sheet was deleted, or the range was fully
 * consumed by a delete: the caller drops the whole `<definedName>` then.
 */
function rewriteDefinedNameFormula(text: string, changesByOriginalName: ReadonlyMap<string, SheetChange>): string | null {
  const match = SINGLE_AREA_NAME.exec(text.trim())
  if (!match) return text

  const sheetName = (match[1] ?? match[2]).replace(/''/g, "'")
  const change = changesByOriginalName.get(sheetName)
  if (!change) return text
  if (change.deleted) return null

  const start = decodeDollarCell(match[3])
  const end = match[4] ? decodeDollarCell(match[4]) : start
  if (!start || !end) return text

  const range = {
    r0: Math.min(start.r, end.r),
    c0: Math.min(start.c, end.c),
    r1: Math.max(start.r, end.r),
    c1: Math.max(start.c, end.c),
  }
  const remapped = remapSqref(`${encodeCol(range.c0)}${range.r0 + 1}:${encodeCol(range.c1)}${range.r1 + 1}`, change.rowSources, change.colSources)
  if (!remapped) return null

  const [remappedStart, remappedEnd] = remapped.split(':')
  const remappedStartCell = decodeDollarCell(remappedStart)!
  const newName = change.newName ?? sheetName
  const needsQuote = !/^[A-Za-z_][\w.]*$/.test(newName)
  const prefix = needsQuote ? `'${newName.replace(/'/g, "''")}'` : newName
  const startOut = encodeDollarCell(remappedStartCell.r, remappedStartCell.c, start.colAbs, start.rowAbs)
  if (!match[4]) return `${prefix}!${startOut}`
  const remappedEndCell = decodeDollarCell(remappedEnd ?? remappedStart)!
  const endOut = encodeDollarCell(remappedEndCell.r, remappedEndCell.c, end.colAbs, end.rowAbs)
  return `${prefix}!${startOut}:${endOut}`
}

/**
 * Applies every defined-name edit implied by the sheet structure change:
 * `localSheetId` renumbering (or removal, for a name scoped to a deleted
 * sheet) and, where the formula matches `rewriteDefinedNameFormula`'s narrow
 * pattern, a rename/re-anchor of the formula text itself.
 */
function updateDefinedNames(
  workbookRoot: Element,
  changesByOriginalName: ReadonlyMap<string, SheetChange>,
  originalIndexToNewIndex: ReadonlyMap<number, number>,
): void {
  const definedNames = firstChildElement(workbookRoot, 'definedNames')
  if (!definedNames) return

  for (const nameEl of childElements(definedNames, 'definedName')) {
    const localSheetId = nameEl.getAttribute('localSheetId')
    if (localSheetId !== null) {
      const newIndex = originalIndexToNewIndex.get(Number(localSheetId))
      if (newIndex === undefined) {
        nameEl.parentNode?.removeChild(nameEl)
        continue
      }
      nameEl.setAttribute('localSheetId', String(newIndex))
    }

    const rewritten = rewriteDefinedNameFormula(nameEl.textContent ?? '', changesByOriginalName)
    if (rewritten === null) {
      nameEl.parentNode?.removeChild(nameEl)
      continue
    }
    if (rewritten !== nameEl.textContent) nameEl.textContent = rewritten
  }

  if (definedNames.children.length === 0) definedNames.parentNode?.removeChild(definedNames)
}

// ---------------------------------------------------------------------------
// docProps/app.xml (best-effort; see module header)
// ---------------------------------------------------------------------------

/**
 * Updates `docProps/app.xml`'s worksheet-title list when it is present and
 * looks like a plain, single-sheet-type workbook — its `<TitlesOfParts>`
 * vector holds exactly as many entries as the ORIGINAL workbook had sheets.
 * A workbook whose titles vector mixes in named ranges or other part kinds
 * (a different count) is left untouched rather than guessed at.
 */
async function updateAppPropsTitles(
  zip: JSZip,
  originalSheetCount: number,
  newSheetNames: ReadonlyArray<string>,
): Promise<void> {
  const path = 'docProps/app.xml'
  const xml = await zip.file(path)?.async('string')
  if (!xml) return

  let doc: XMLDocument
  try {
    doc = parseXmlPart(xml)
  } catch {
    return
  }
  const root = doc.documentElement
  const titlesVector = firstChildElement(firstChildElement(root, 'TitlesOfParts') ?? root, 'vector')
  if (!titlesVector || titlesVector.parentElement?.localName !== 'TitlesOfParts') return
  const titleEls = childElements(titlesVector, 'lpstr')
  if (titleEls.length !== originalSheetCount) return // mixed content or unexpected shape — leave it alone

  // Build with the SAME prefix `<vt:vector>` itself uses (conventionally
  // "vt", but read from the source rather than assumed) — creating the
  // element with a bare, unprefixed qualified name would serialize it
  // against whatever default namespace is in scope at that point in the
  // document (`Properties`' own, not `vt`), which is a different element
  // entirely even though `Element.localName` alone can't tell them apart.
  const qualifiedName = titlesVector.prefix ? `${titlesVector.prefix}:lpstr` : 'lpstr'
  while (titlesVector.firstChild) titlesVector.removeChild(titlesVector.firstChild)
  for (const name of newSheetNames) {
    const lpstr = doc.createElementNS(titlesVector.namespaceURI, qualifiedName)
    lpstr.textContent = name
    titlesVector.appendChild(lpstr)
  }
  titlesVector.setAttribute('size', String(newSheetNames.length))

  // `HeadingPairs` carries the Worksheets COUNT as the `vt:i4` right after
  // the `<vt:lpstr>Worksheets</vt:lpstr>` variant marker.
  const headingVector = firstChildElement(firstChildElement(root, 'HeadingPairs') ?? root, 'vector')
  if (headingVector?.parentElement?.localName === 'HeadingPairs') {
    const variants = childElements(headingVector, 'variant')
    for (let i = 0; i < variants.length; i++) {
      const marker = firstChildElement(variants[i], 'lpstr')
      if (marker?.textContent === 'Worksheets') {
        const countEl = firstChildElement(variants[i + 1], 'i4')
        if (countEl) countEl.textContent = String(newSheetNames.length)
        break
      }
    }
  }

  zip.file(path, serializeXmlPart(doc))
}

// ---------------------------------------------------------------------------
// Save
// ---------------------------------------------------------------------------

/**
 * Rewrites the original workbook with the document's current values,
 * structure and re-anchored ranges. Returns `null` when the original is not
 * an OOXML workbook this can patch at all (non-zip, no readable
 * `workbook.xml`/relationships, or a `sourcePath` the original package no
 * longer has), so the caller falls back to `writeWorkbookBytes`.
 */
export async function writeWorkbookThroughOriginal(
  original: ArrayBuffer,
  document: SpreadsheetDocument,
): Promise<Uint8Array | null> {
  // A document with NO sourcePath-tracked sheet at all isn't confirmably one
  // this buffer was loaded from (a CSV-origin document, or one built by
  // `spreadsheetWrite.ts` itself) — bail to the fresh-workbook writer rather
  // than guess. A document with SOME tracked and some untracked sheets (the
  // common case once a sheet has been added) is exactly what the rest of
  // this function exists to patch.
  if (document.sheets.length === 0 || document.sheets.every((sheet) => sheet.sourcePath === undefined)) return null

  let zip: JSZip
  try {
    zip = await loadWorkbookZip(original)
  } catch {
    return null
  }

  const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
  if (!workbookXml) return null
  let workbookDoc: XMLDocument
  try {
    workbookDoc = parseXmlPart(workbookXml)
  } catch {
    return null
  }
  const workbookRoot = workbookDoc.documentElement
  const sheetsContainer = firstChildElement(workbookRoot, 'sheets')
  if (!sheetsContainer) return null

  const relsXml = (await zip.file(WORKBOOK_RELS_PATH)?.async('string')) ?? null
  let relsDoc: XMLDocument | null = null
  if (relsXml) {
    try {
      relsDoc = parseXmlPart(relsXml)
    } catch {
      relsDoc = null
    }
  }
  if (!relsDoc) return null // no relationships to resolve worksheet parts through — can't safely restructure

  const originalSheets = readOriginalSheetEntries(workbookRoot, relsDoc)
  const pathToOriginal = new Map(originalSheets.filter((s) => s.path).map((s) => [s.path!, s]))

  // Every sourcePath the document claims must actually exist in this
  // original package (a mismatch means the buffer isn't the file this
  // document was loaded from — bail rather than silently mis-restructure).
  for (const sheet of document.sheets) {
    if (sheet.sourcePath !== undefined && !pathToOriginal.has(sheet.sourcePath)) return null
  }

  const originalIndexToNewIndex = new Map<number, number>()
  document.sheets.forEach((sheet, newIndex) => {
    if (sheet.sourcePath !== undefined) {
      originalIndexToNewIndex.set(pathToOriginal.get(sheet.sourcePath)!.originalIndex, newIndex)
    }
  })

  const changesByOriginalName = new Map<string, SheetChange>()
  for (const original of originalSheets) {
    const newIndex = originalIndexToNewIndex.get(original.originalIndex)
    if (newIndex === undefined) {
      changesByOriginalName.set(original.name, { newName: undefined, deleted: true, rowSources: undefined, colSources: undefined })
      continue
    }
    const newSheet = document.sheets[newIndex]
    changesByOriginalName.set(original.name, {
      newName: newSheet.name !== original.name ? newSheet.name : undefined,
      deleted: false,
      rowSources: newSheet.rowSources,
      colSources: newSheet.colSources,
    })
  }

  // ---- relationships / content-types / new-sheet bookkeeping ----
  const relsRoot = relsDoc.documentElement
  let nextRIdNumber = nextNumericSuffix(descendantElements(relsDoc, 'Relationship').map((r) => r.getAttribute('Id') ?? ''), 'rId')
  let nextSheetIdNumber = nextNumericSuffix(originalSheets.map((s) => s.element.getAttribute('sheetId') ?? ''), '')

  const contentTypesXml = await zip.file(CONTENT_TYPES_PATH)?.async('string')
  let contentTypesDoc: XMLDocument | null = null
  if (contentTypesXml !== undefined) {
    try {
      contentTypesDoc = parseXmlPart(contentTypesXml)
    } catch {
      contentTypesDoc = null
    }
  }

  // Rebuild `<sheets>` in the document's own order — this is what expresses
  // an add/delete/reorder, all in one pass.
  const freshSheets = workbookDoc.createElementNS(workbookRoot.namespaceURI, 'sheets')
  const perSheetOriginal = new Map<number, OriginalSheet>() // newIndex -> parsed worksheet
  const perSheetPartPath = new Map<number, string>() // newIndex -> part path to write back

  for (const [newIndex, sheet] of document.sheets.entries()) {
    if (sheet.sourcePath !== undefined) {
      const originalEntry = pathToOriginal.get(sheet.sourcePath)!
      const clonedSheetEl = originalEntry.element.cloneNode(false) as Element
      clonedSheetEl.setAttribute('name', sheet.name)
      freshSheets.appendChild(clonedSheetEl)

      const sheetXml = await zip.file(originalEntry.path!)?.async('string')
      if (!sheetXml) return null
      let parsed: OriginalSheet | null
      try {
        parsed = readOriginalSheet(sheetXml)
      } catch {
        return null
      }
      if (!parsed) return null
      perSheetOriginal.set(newIndex, parsed)
      perSheetPartPath.set(newIndex, originalEntry.path!)
      continue
    }

    // A sheet added in Atlas: allocate a fresh part/relationship/sheetId.
    const partPath = nextWorksheetPartPath(zip)
    const rId = `rId${nextRIdNumber++}`
    const sheetId = nextSheetIdNumber++

    const sheetEl = workbookDoc.createElementNS(workbookRoot.namespaceURI, 'sheet')
    sheetEl.setAttribute('name', sheet.name)
    sheetEl.setAttribute('sheetId', String(sheetId))
    sheetEl.setAttributeNS(RELATIONSHIPS_NS, 'r:id', rId)
    freshSheets.appendChild(sheetEl)

    const relationshipEl = relsDoc.createElementNS(PACKAGE_RELS_NS, 'Relationship')
    relationshipEl.setAttribute('Id', rId)
    relationshipEl.setAttribute('Type', WORKSHEET_REL_TYPE)
    relationshipEl.setAttribute('Target', relativeToXl(partPath))
    relsRoot.appendChild(relationshipEl)

    if (contentTypesDoc) addContentTypeOverride(contentTypesDoc, partPath, WORKSHEET_CONTENT_TYPE)

    perSheetOriginal.set(newIndex, blankOriginalSheet())
    perSheetPartPath.set(newIndex, partPath)
  }
  workbookRoot.replaceChild(freshSheets, sheetsContainer)

  // Sheets present in the original but not in the new document: remove
  // their part, their worksheet-level rels (if any), the workbook
  // relationship, and their content-type override.
  for (const original of originalSheets) {
    if (originalIndexToNewIndex.has(original.originalIndex) || !original.path) continue
    zip.remove(original.path)
    const slash = original.path.lastIndexOf('/')
    const sheetRelsPath = `${original.path.slice(0, slash)}/_rels/${original.path.slice(slash + 1)}.rels`
    if (zip.file(sheetRelsPath)) zip.remove(sheetRelsPath)
    for (const rel of descendantElements(relsDoc, 'Relationship')) {
      if (rel.getAttribute('Id') === original.rId) rel.parentNode?.removeChild(rel)
    }
    if (contentTypesDoc) removeContentTypeOverrideForPart(contentTypesDoc, original.path)
  }

  // ---- per-sheet content: cells, dimension, merges, columns, ranges, tables ----
  for (const [newIndex, sheet] of document.sheets.entries()) {
    const parsed = perSheetOriginal.get(newIndex)!
    const partPath = perSheetPartPath.get(newIndex)!

    rebuildSheetData(sheet, parsed)
    updateDimension(parsed.root, sheet)
    updateMergeCells(parsed.doc, parsed.root, sheet)
    updateColumns(parsed.doc, parsed.root, sheet)
    updateShiftedRanges(parsed.root, sheet)
    zip.file(partPath, serializeXmlPart(parsed.doc))

    for (const [tableIndex, table] of (sheet.tables ?? []).entries()) {
      if (table.partPath === undefined) continue
      zip.file(table.partPath, rewriteTableXml(table, tableIndex + 1, tableHeaderNames(table, sheet.rows)))
    }
  }

  updateDefinedNames(workbookRoot, changesByOriginalName, originalIndexToNewIndex)
  markRecalculateOnLoad(workbookDoc, workbookRoot)
  if (removeCalcChain(zip, relsDoc) && contentTypesDoc) removeContentTypeOverrideForPart(contentTypesDoc, 'xl/calcChain.xml')

  if (contentTypesDoc) zip.file(CONTENT_TYPES_PATH, serializeXmlPart(contentTypesDoc))
  zip.file(WORKBOOK_RELS_PATH, serializeXmlPart(relsDoc))
  zip.file('xl/workbook.xml', serializeXmlPart(workbookDoc))

  await updateAppPropsTitles(
    zip,
    originalSheets.length,
    document.sheets.map((s) => s.name),
  )

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
