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
 *    renumbered; every name's formula text — single-area, multi-area
 *    (`Sheet1!$A$1,Sheet1!$C$3`), function-wrapped (`OFFSET(Sheet1!$A$1,...)`)
 *    or a whole-row/whole-column range (a `Print_Titles` `$1:$1`) — is
 *    rewritten through a sheet rename/delete and re-anchored through a
 *    row/column insert or delete by the same engine a cell formula goes
 *    through (`formulaRefs.ts`), not just the narrow single-area case;
 *  - a CELL FORMULA referencing a renamed, reordered or deleted sheet
 *    (`=Sheet2!A1`, `='My Sheet'!A1:B2`, a 3-D `=SUM(Sheet1:Sheet3!A1)`), or
 *    one referencing cells moved by a row/column insert or delete — on its
 *    own sheet or another one — is rewritten the way Excel itself rewrites
 *    it (`formulaRefs.ts`); a reference to a deleted sheet, or a range fully
 *    consumed by a delete, becomes `#REF!`, matching Excel's own documented
 *    behavior. This applies fully to a formula CLONED byte-for-byte from the
 *    original file; a formula the user actually typed in Atlas only has its
 *    SHEET NAME fixed (never its coordinates, which are already expressed
 *    against the current, saved layout) — see that module's header for why;
 *  - a deleted sheet's own EXCLUSIVELY-owned parts — its table(s), comments
 *    (legacy VML note shapes and modern threaded comments alike) and its own
 *    drawing (one level further, any chart embedded in it) — are removed
 *    from the package, not left as dead weight (`removeSheetOwnedParts`);
 *  - `xl/sharedStrings.xml` is compacted after every save: an entry no cell
 *    anywhere in the package references any more is dropped, the survivors'
 *    indices renumbered contiguously (and every affected cell's `<v>`
 *    updated to match), and `count`/`uniqueCount` recomputed
 *    (`compactSharedStrings`);
 *  - `calcChain.xml` is dropped and the workbook is marked "recalculate on
 *    load", since cached formula values are only as good as this app's own
 *    formula engine;
 *  - Excel tables are rewritten in place (range, column names);
 *  - a shared-formula group's `ref` SPAN attribute
 *    (`<f t="shared" ref="C3:C10" si="0">`) is re-anchored right alongside
 *    the master cell's own formula TEXT when a row/column insert or delete
 *    moves some of its member cells — leaving the span stale would make
 *    Excel apply the (unrelated) old formula to the wrong range on open;
 *  - a formula rewritten into (or partly into) `#REF!` by a structural edit
 *    has its now-stale cached `<v>` cleared, not left showing a value that
 *    visibly contradicts its own formula (see `rewriteClonedFormula`);
 *  - an `xl/media/*` image is swept once it is PROVABLY unreferenced by
 *    every remaining relationship in the package (not just the deleted
 *    sheet's own drawing) after a sheet delete removes that sheet's own
 *    drawing part (`sweepOrphanedMedia`).
 *
 * Everything else in the package is passed through untouched.
 *
 * Deliberate limits (see the findings register for the full list, and
 * `formulaRefs.ts`'s own header for the formula-syntax engine's specific
 * coverage): a non-OOXML target still falls back to the fresh-workbook
 * writer; a 3-D formula reference's row/column shift is taken from its
 * FIRST sheet only, on the assumption (implicit in a 3-D reference itself)
 * that every sheet it spans is laid out the same way — a row/column
 * inserted or deleted on a sheet in the MIDDLE of the span (not its first
 * sheet) never shifts the reference at all. This is left as-is rather than
 * attempting a per-sheet shift a single A1-style reference cannot even
 * represent (one reference, applied uniformly to every sheet in the span):
 * real Excel has the same well-documented limitation (a 3-D reference is
 * one of the cases Microsoft itself flags as unreliable across
 * insert/delete/sort operations on the sheets it spans), so matching "leave
 * it alone unless the first sheet moved" is the conservative, Excel-shaped
 * choice, not an oversight — see `formulaRefs.ts`'s own header and
 * `rewriteReference`'s `shiftSource` comment.
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
import { rewriteFormulaReferences, type SheetChange } from './formulaRefs'
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

// SHEET-3/4/5 — Excel-parity auto-typing on write (the xlsx passthrough
// path only; `spreadsheetWrite.ts`'s fresh-workbook writer has its own,
// deliberately simpler `typeCellText` for the formats that never carry a
// `styles.xml` to consult at all).
//
// A handful of a builtin numFmt codes ECMA-376 Part 1 §18.8.30 fixes by ID
// (never spelled out in `styles.xml` itself) — just enough to classify a
// cell's EXISTING format, not a full builtin table. Anything not listed
// here that still resolves to a real numFmtId (there are ~50 more, mostly
// currency/accounting/time variants irrelevant to this classification)
// falls through `classifyFormatCode`'s own `'other'` catch-all, which is
// the conservative "respect it, don't touch" outcome either way.
const BUILTIN_NUMFMT_CODES: ReadonlyMap<number, string> = new Map([
  [0, 'General'],
  [9, '0%'],
  [10, '0.00%'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [22, 'm/d/yy h:mm'],
  [49, '@'],
])

type NumFmtKind = 'general' | 'text' | 'percent' | 'date' | 'other'

/** Classifies a resolved `formatCode` string (never a numFmtId) for the auto-typing decision below. */
function classifyFormatCode(code: string): NumFmtKind {
  if (code === '' || code.toLowerCase() === 'general') return 'general'
  if (code === '@') return 'text'
  if (code.includes('%')) return 'percent'
  // Date detection: strip bracketed locale/color tags (`[$-409]`, `[Red]`)
  // and quoted literals before looking for date tokens, so a literal string
  // like `"m" units` doesn't false-positive.
  const stripped = code.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')
  if (/(yy|mm|dd|m\/d|d\/m|mmm)/i.test(stripped)) return 'date'
  return 'other'
}

/** Parsed once per save from `xl/styles.xml`; `null` when that part is missing/unreadable or has no `<cellXfs>` at all — every helper below then degrades to "don't touch the format", never throws. */
type StylesContext = {
  readonly doc: XMLDocument
  readonly root: Element
  numFmtsEl: Element | null
  readonly cellXfsEl: Element
  readonly xfs: Element[]
  readonly numFmtCodeById: Map<number, string>
  nextCustomNumFmtId: number
  dirty: boolean
}

function loadStylesContext(stylesXml: string | undefined): StylesContext | null {
  if (!stylesXml) return null
  let doc: XMLDocument
  try {
    doc = parseXmlPart(stylesXml)
  } catch {
    return null
  }
  const root = doc.documentElement
  const cellXfsEl = firstChildElement(root, 'cellXfs')
  if (!cellXfsEl) return null // no cellXfs at all — nothing to index cell styles into

  const numFmtsEl = firstChildElement(root, 'numFmts')
  const numFmtCodeById = new Map<number, string>()
  if (numFmtsEl) {
    for (const el of childElements(numFmtsEl, 'numFmt')) {
      const id = Number(el.getAttribute('numFmtId'))
      const code = el.getAttribute('formatCode') ?? ''
      if (Number.isFinite(id)) numFmtCodeById.set(id, code)
    }
  }
  let nextCustomNumFmtId = 164 // first ID the OOXML spec reserves for custom (non-builtin) formats
  for (const id of numFmtCodeById.keys()) {
    if (id >= nextCustomNumFmtId) nextCustomNumFmtId = id + 1
  }

  return {
    doc,
    root,
    numFmtsEl,
    cellXfsEl,
    xfs: childElements(cellXfsEl, 'xf'),
    numFmtCodeById,
    nextCustomNumFmtId,
    dirty: false,
  }
}

function resolveNumFmtCode(ctx: StylesContext, numFmtId: number): string {
  return ctx.numFmtCodeById.get(numFmtId) ?? BUILTIN_NUMFMT_CODES.get(numFmtId) ?? ''
}

/** What format an existing cell (by its `s` style-index attribute, or `null` for none = default General) already carries. */
function classifyStyle(ctx: StylesContext | null, styleIndexStr: string | null): NumFmtKind {
  if (!ctx || styleIndexStr === null) return 'general'
  const idx = Number(styleIndexStr)
  const xf = Number.isFinite(idx) ? ctx.xfs[idx] : undefined
  if (!xf) return 'general'
  const numFmtId = Number(xf.getAttribute('numFmtId') ?? '0')
  if (!Number.isFinite(numFmtId) || numFmtId === 0) return 'general'
  const code = resolveNumFmtCode(ctx, numFmtId)
  if (code === '') return 'other' // an unresolvable numFmtId — conservative: leave it alone
  return classifyFormatCode(code)
}

/** True when `candidate`'s non-numFmt attributes (font/fill/border/alignment refs) already match `base` (or match the all-default xf when there is no base) — lets a new percent/date style reuse an existing `cellXfs` entry instead of growing the array on every save that happens to hit this path. */
function sameXfExceptNumFmt(candidate: Element, base: Element | undefined): boolean {
  const attrs = ['fontId', 'fillId', 'borderId', 'xfId', 'applyFont', 'applyFill', 'applyBorder', 'applyAlignment']
  for (const name of attrs) {
    const a = candidate.getAttribute(name) ?? ''
    const b = base?.getAttribute(name) ?? ''
    if (a !== b) return false
  }
  return true
}

/** Finds (or creates) a custom `<numFmt>` entry for `code` in `<numFmts>`, returning its numFmtId. `<numFmts>` must be the FIRST child of `<styleSheet>` per CT_Stylesheet's fixed element order — created there when it doesn't exist yet. */
function ensureCustomNumFmtId(ctx: StylesContext, code: string): number {
  for (const [id, existingCode] of ctx.numFmtCodeById) {
    if (existingCode === code) return id
  }
  const id = ctx.nextCustomNumFmtId++
  ctx.numFmtCodeById.set(id, code)
  let numFmtsEl = ctx.numFmtsEl
  if (!numFmtsEl) {
    numFmtsEl = ctx.doc.createElementNS(ctx.root.namespaceURI, 'numFmts')
    ctx.root.insertBefore(numFmtsEl, ctx.root.firstChild)
    ctx.numFmtsEl = numFmtsEl
  }
  const el = ctx.doc.createElementNS(ctx.root.namespaceURI, 'numFmt')
  el.setAttribute('numFmtId', String(id))
  el.setAttribute('formatCode', code)
  numFmtsEl.appendChild(el)
  numFmtsEl.setAttribute('count', String(numFmtsEl.children.length))
  ctx.dirty = true
  return id
}

/** Finds (or creates) a `cellXfs` entry using `numFmtId`, cloning `baseStyleIndexStr`'s other attributes (font/fill/border) so assigning a percent/date format doesn't discard an edited cell's existing look. Returns the new `s` index as a string, ready to set on the `<c>` element. */
function ensureXfWithNumFmt(ctx: StylesContext, baseStyleIndexStr: string | null, numFmtId: number): string {
  const baseIdx = baseStyleIndexStr !== null ? Number(baseStyleIndexStr) : NaN
  const base = Number.isFinite(baseIdx) ? ctx.xfs[baseIdx] : undefined

  for (let i = 0; i < ctx.xfs.length; i++) {
    if (Number(ctx.xfs[i].getAttribute('numFmtId') ?? '0') !== numFmtId) continue
    if (sameXfExceptNumFmt(ctx.xfs[i], base)) return String(i)
  }

  const fresh = base ? (base.cloneNode(true) as Element) : ctx.doc.createElementNS(ctx.root.namespaceURI, 'xf')
  fresh.setAttribute('numFmtId', String(numFmtId))
  fresh.setAttribute('applyNumberFormat', '1')
  if (!base) {
    fresh.setAttribute('fontId', '0')
    fresh.setAttribute('fillId', '0')
    fresh.setAttribute('borderId', '0')
    fresh.setAttribute('xfId', '0')
  }
  ctx.cellXfsEl.appendChild(fresh)
  ctx.xfs.push(fresh)
  ctx.cellXfsEl.setAttribute('count', String(ctx.xfs.length))
  ctx.dirty = true
  return String(ctx.xfs.length - 1)
}

/** Reuses/creates a percent `cellXfs` entry — builtin `0%`(9) or `0.00%`(10), matching how many decimal places the user actually typed (`"50%"` vs `"12.3%"`). */
function ensurePercentStyle(ctx: StylesContext | null, baseStyle: string | null, hasDecimal: boolean): string | null {
  if (!ctx) return null
  return ensureXfWithNumFmt(ctx, baseStyle, hasDecimal ? 10 : 9)
}

/** Reuses/creates a date `cellXfs` entry with a custom `yyyy-mm-dd` numFmt — the same unambiguous shape this only ever recognizes on input (see `buildCell`), so the cell displays back exactly what the user typed. */
function ensureDateStyle(ctx: StylesContext | null, baseStyle: string | null): string | null {
  if (!ctx) return null
  const numFmtId = ensureCustomNumFmtId(ctx, 'yyyy-mm-dd')
  return ensureXfWithNumFmt(ctx, baseStyle, numFmtId)
}

const PERCENT_TEXT = /^-?\d+(\.\d+)?%$/
const ISO_DATE_TEXT = /^(\d{4})-(\d{2})-(\d{2})$/

/**
 * Excel's own serial-date numbering: whole days since the fictitious
 * 1899-12-30 "day zero" (one day before 1900-01-01, which is itself day 1) —
 * the classic epoch trick that also reproduces Excel's well-known "1900 was
 * a leap year" bug for any real date on or after 1900-03-01 without special
 * cased for it, since the phantom Feb-29-1900 falls before this range.
 * Returns `null` for a string that parses as digits but isn't a real
 * calendar date (`2024-02-30`, `2024-13-01`) — `Date.UTC` itself normalizes
 * those into a DIFFERENT valid date instead of rejecting them, so the
 * round-trip-through-UTC-components check below is what actually catches it.
 */
function excelSerialFromIsoDate(match: RegExpExecArray): number | null {
  const year = Number(match[1])
  const month = Number(match[2])
  const day = Number(match[3])
  const asUtc = Date.UTC(year, month - 1, day)
  const roundTrip = new Date(asUtc)
  if (roundTrip.getUTCFullYear() !== year || roundTrip.getUTCMonth() !== month - 1 || roundTrip.getUTCDate() !== day) {
    return null
  }
  const epoch = Date.UTC(1899, 11, 30)
  return Math.round((asUtc - epoch) / 86_400_000)
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

/**
 * Builds the `<c>` element for a cell that was edited (or is new).
 *
 * SHEET-3/4/5 auto-typing (a user-typed value only — `formula !== undefined`
 * skips all of it and keeps the prior plain numeric-vs-text rule, since a
 * formula's text is a COMPUTED result, never something the user typed):
 *
 *  - A leading `'` (Excel's own "force text" quote prefix) or a cell already
 *    styled `@` (text format) makes the value text NO MATTER what it looks
 *    like — `'007` and `007` in an `@`-formatted cell both save as the text
 *    `"007"`, apostrophe never stored. Plain `007` with neither still becomes
 *    the number `7`, unchanged: that already matches Excel's own default.
 *  - An unambiguous `N%`/`N.N%` becomes a numeric value (`/100`) with a
 *    percent format; an unambiguous ISO `YYYY-MM-DD` becomes an Excel date
 *    serial with a date format — but ONLY when the target cell's own
 *    existing format is General. A cell that already carries some other
 *    explicit format keeps that format (`ensure*Style` is never called), so
 *    an existing choice is never second-guessed — just the parsed numeric
 *    value is written under it, mirroring how Excel itself won't silently
 *    swap a cell's chosen format out from under a typed value.
 */
function buildCell(
  doc: XMLDocument,
  namespace: string | null,
  address: string,
  text: string,
  formula: string | undefined,
  style: string | null,
  stylesCtx: StylesContext | null,
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

  // A formula's cached value is a COMPUTED result, not something the user
  // typed — none of the auto-typing below applies to it; it keeps the prior
  // plain numeric-vs-text rule exactly (a numeric result gets no `t`
  // attribute at all, matching a plain numeric cell; anything else is
  // `t="str"`).
  if (formula !== undefined) {
    if (isNumericText(text)) {
      const value = doc.createElementNS(namespace, 'v')
      value.textContent = String(Number(text.trim()))
      cell.appendChild(value)
      return cell
    }
    cell.setAttribute('t', 'str')
    const value = doc.createElementNS(namespace, 'v')
    value.textContent = xmlSafeText(text)
    cell.appendChild(value)
    return cell
  }

  const formatKind = classifyStyle(stylesCtx, style)
  let raw = text
  let forceText = formatKind === 'text'
  if (raw.startsWith("'")) {
    raw = raw.slice(1)
    forceText = true
  }

  // Percent/date auto-typing needs `stylesCtx` to assign (or find) a format
  // for the General case — without it (missing/unreadable `xl/styles.xml`,
  // see `loadStylesContext`), skip straight to the plain numeric/text rule
  // below rather than writing an unformatted numeric value that would
  // display as a bare `0.5` or `45365` instead of the percent/date the user
  // actually typed.
  if (!forceText && stylesCtx) {
    const trimmed = raw.trim()
    const percentMatch = PERCENT_TEXT.exec(trimmed)
    if (percentMatch) {
      const numeric = Number(trimmed.slice(0, -1)) / 100
      if (Number.isFinite(numeric)) {
        if (formatKind === 'general') {
          const newStyle = ensurePercentStyle(stylesCtx, style, percentMatch[1] !== undefined)
          if (newStyle !== null) cell.setAttribute('s', newStyle)
        }
        const value = doc.createElementNS(namespace, 'v')
        value.textContent = String(numeric)
        cell.appendChild(value)
        return cell
      }
    }

    const dateMatch = ISO_DATE_TEXT.exec(trimmed)
    if (dateMatch) {
      const serial = excelSerialFromIsoDate(dateMatch)
      if (serial !== null) {
        if (formatKind === 'general') {
          const newStyle = ensureDateStyle(stylesCtx, style)
          if (newStyle !== null) cell.setAttribute('s', newStyle)
        }
        const value = doc.createElementNS(namespace, 'v')
        value.textContent = String(serial)
        cell.appendChild(value)
        return cell
      }
    }
  }

  if (!forceText && isNumericText(raw)) {
    const value = doc.createElementNS(namespace, 'v')
    value.textContent = String(Number(raw.trim()))
    cell.appendChild(value)
    return cell
  }

  // Inline strings avoid touching sharedStrings.xml (Excel reads both).
  cell.setAttribute('t', 'inlineStr')
  const inlineString = doc.createElementNS(namespace, 'is')
  const textElement = doc.createElementNS(namespace, 't')
  textElement.setAttribute('xml:space', 'preserve')
  textElement.textContent = xmlSafeText(raw)
  inlineString.appendChild(textElement)
  cell.appendChild(inlineString)
  return cell
}

/**
 * Rewrites one cloned (untouched) cell's `<f>` text in place, when it has
 * one, through this sheet's own row/column history and every sheet's
 * rename/delete (USR-17 follow-up, see `formulaRefs.ts`). Byte-for-byte
 * copied text is exactly what needs this: it is still expressed in the
 * ORIGINAL file's numbering and sheet names.
 */
function rewriteClonedFormula(
  clone: Element,
  sheet: EditableSheet,
  changesByOriginalName: ReadonlyMap<string, SheetChange>,
): void {
  const formulaElement = firstChildElement(clone, 'f')
  if (!formulaElement) return // no formula on this cell at all

  // A shared-formula MASTER cell also carries the group's `ref` SPAN
  // attribute (`<f t="shared" ref="B2:B10" si="0">`), which is a plain
  // sqref-style range on THIS sheet — a shared-formula group can never span
  // more than one sheet — so it re-anchors through the sheet's own
  // row/column history exactly like a conditional format's `sqref`
  // (`updateShiftedRanges`/`remapSqref`), not through `formulaRefs.ts`
  // (that engine only rewrites reference syntax INSIDE a formula's text).
  // Previously only the master's own formula TEXT was rewritten, leaving
  // the `ref` span pointing at the ORIGINAL rows/columns — a stale span
  // makes Excel apply the (unrelated) old formula to the wrong range on
  // open. The master cell being cloned here always still exists in the
  // saved sheet, so its own row/column is always inside the remapped span
  // and `remapSqref` is never asked to drop the whole thing (its `null`
  // result — "fully consumed by a delete" — is therefore unreachable here;
  // the attribute is simply left as-is in that theoretical case rather than
  // ever emitting an invalid, empty `ref`).
  if (formulaElement.getAttribute('t') === 'shared') {
    const ref = formulaElement.getAttribute('ref')
    if (ref) {
      const nextRef = remapSqref(ref, sheet.rowSources, sheet.colSources)
      if (nextRef) formulaElement.setAttribute('ref', nextRef)
    }
  }

  const text = formulaElement.textContent
  if (!text) return // a shared/array-formula FOLLOWER has no text of its own to rewrite

  const rewritten = rewriteFormulaReferences(text, {
    changesByOriginalName,
    ownRowSources: sheet.rowSources,
    ownColSources: sheet.colSources,
    remapCoordinates: true,
  })
  formulaElement.textContent = xmlSafeText(rewritten)

  // A structural edit (a sheet delete, or a row/column delete that fully
  // consumed a referenced range) can turn part or all of this formula into
  // `#REF!`. In Excel, a `#REF!` anywhere inside an expression propagates
  // through the whole formula unless it is caught by IFERROR/IFNA — which
  // this module has no way to detect from the outside — so a cached `<v>`
  // computed before the edit is no longer trustworthy either way. A bug
  // hunt found exactly this going wrong: a stale `99` still on display
  // under a formula that now reads `#REF!`, which is worse than a blank
  // cell because it looks like a real, current answer instead of an
  // obviously broken one. Atlas's own formula evaluator has no cross-sheet
  // support at all (a deliberate scope cut — see this module's header), so
  // recomputing the correct value here is not an option; clearing the
  // stale cache and relying on the "recalculate on load" flag this module
  // already sets unconditionally (`markRecalculateOnLoad`) is the honest,
  // conservative choice — Excel fills in the real value the moment the
  // file is opened, so the emptied `<v>` is never actually shown to a user
  // in Excel.
  if (rewritten.includes('#REF!') && !text.includes('#REF!')) {
    // Write the error as the cached value (`t="e"`, `<v>#REF!</v>`) — what
    // Excel itself stores for a formula whose result is an error. Simply
    // dropping the `<v>` would be valid OOXML but invisible on reopen:
    // SheetJS, which every reader here goes through, skips a formula cell
    // that has no cached value entirely, so the formula would look lost.
    const staleValue = firstChildElement(clone, 'v')
    if (staleValue) clone.removeChild(staleValue)
    const errorValue = clone.ownerDocument.createElementNS(clone.namespaceURI, 'v')
    errorValue.textContent = '#REF!'
    clone.appendChild(errorValue)
    clone.setAttribute('t', 'e')
  }
}

function rebuildSheetData(
  sheet: EditableSheet,
  original: OriginalSheet,
  changesByOriginalName: ReadonlyMap<string, SheetChange>,
  stylesCtx: StylesContext | null,
): void {
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
        rewriteClonedFormula(clone, sheet, changesByOriginalName)
        rowElement.appendChild(clone)
        continue
      }

      const style = originalCell?.getAttribute('s') ?? null
      if (text === '' && formula === undefined) {
        // Keep an empty-but-styled cell so its formatting survives.
        if (style !== null) rowElement.appendChild(buildCell(doc, namespace, address, '', undefined, style, stylesCtx))
        continue
      }
      // A formula the user typed (or a formula cell otherwise rebuilt from
      // the live model) is already expressed against the CURRENT, saved
      // layout — only a stale sheet NAME needs fixing here, never a
      // coordinate (see `formulaRefs.ts`'s module header).
      const rewrittenFormula =
        formula === undefined ? undefined : rewriteFormulaReferences(formula, { changesByOriginalName, remapCoordinates: false })
      rowElement.appendChild(buildCell(doc, namespace, address, text, rewrittenFormula, style, stylesCtx))
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

  // SHEET-9 — `<ignoredErrors>` (the "number stored as text"/"formula
  // omits cells" markers Excel writes so it doesn't keep flagging the same
  // cell) has its own `sqref`-bearing `<ignoredError>` children, same shape
  // as `dataValidation`, but this element was never re-anchored through a
  // row/column insert or delete at all — left stale while every sibling
  // above (`conditionalFormatting`, `dataValidation`) already tracked the
  // sheet's growth/shrink.
  const ignoredErrors = firstChildElement(root, 'ignoredErrors')
  if (ignoredErrors) {
    let survivors = 0
    for (const ignoredError of childElements(ignoredErrors, 'ignoredError')) {
      const sqref = ignoredError.getAttribute('sqref')
      const next = sqref ? remapSqref(sqref, rowSources, colSources) : null
      if (next) {
        ignoredError.setAttribute('sqref', next)
        survivors += 1
      } else {
        ignoredErrors.removeChild(ignoredError)
      }
    }
    if (survivors === 0) ignoredErrors.parentNode?.removeChild(ignoredErrors)
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

const SHARED_STRINGS_PATH = 'xl/sharedStrings.xml'

/**
 * Drops every `xl/sharedStrings.xml` entry no longer referenced by any cell
 * in the saved package, compacting the remaining entries' indices and
 * recomputing `count` (total `t="s"` cell references across the workbook)
 * and `uniqueCount` (surviving `<si>` entries) — hygiene an edit or a
 * row/column delete otherwise leaves behind. Atlas itself never ADDS a
 * shared-string reference (every new or edited cell is written as an
 * inline string instead — see `buildCell`), so across successive saves this
 * table only ever accumulates stale entries, never grows for a reason.
 *
 * Every `<c t="s">` cell across every sheet being saved has its `<v>` index
 * rewritten in place when its entry's position moved. A cell whose `<v>`
 * is missing or not a plain integer is left alone entirely (a malformed
 * shared-string reference is not this function's problem to fix) — its
 * index still counts as "used" via whatever numeric value can be read, so a
 * borderline case never causes a used entry to be dropped.
 */
async function compactSharedStrings(zip: JSZip, sheets: ReadonlyArray<OriginalSheet>): Promise<void> {
  const xml = await zip.file(SHARED_STRINGS_PATH)?.async('string')
  if (!xml) return
  let doc: XMLDocument
  try {
    doc = parseXmlPart(xml)
  } catch {
    return
  }
  const root = doc.documentElement
  const entries = childElements(root, 'si')
  if (entries.length === 0) return

  const cellsByIndex = new Map<number, Element[]>()
  let totalReferences = 0
  for (const sheet of sheets) {
    for (const cell of descendantElements(sheet.root, 'c')) {
      if (cell.getAttribute('t') !== 's') continue
      const valueEl = firstChildElement(cell, 'v')
      const index = valueEl?.textContent === undefined || valueEl.textContent === null ? NaN : Number(valueEl.textContent)
      if (!Number.isInteger(index)) continue
      totalReferences += 1
      const list = cellsByIndex.get(index)
      if (list) list.push(cell)
      else cellsByIndex.set(index, [cell])
    }
  }

  const usedOldIndices = Array.from(cellsByIndex.keys())
    .filter((index) => index >= 0 && index < entries.length)
    .sort((a, b) => a - b)

  if (usedOldIndices.length === entries.length && usedOldIndices.every((value, i) => value === i)) {
    // Every entry is still used and already contiguous — only `count` (a pure reference-count tally) can possibly be stale.
    if (root.getAttribute('count') !== String(totalReferences)) root.setAttribute('count', String(totalReferences))
    zip.file(SHARED_STRINGS_PATH, serializeXmlPart(doc))
    return
  }

  const remap = new Map<number, number>()
  usedOldIndices.forEach((oldIndex, newIndex) => remap.set(oldIndex, newIndex))
  for (const [oldIndex, cells] of cellsByIndex) {
    const newIndex = remap.get(oldIndex)
    if (newIndex === undefined || newIndex === oldIndex) continue
    for (const cell of cells) {
      const valueEl = firstChildElement(cell, 'v')
      if (valueEl) valueEl.textContent = String(newIndex)
    }
  }

  // `<sst>` is `si*, extLst?` — reinsert the surviving entries (reordered to
  // their new, compacted indices) before any `extLst`, not blindly appended,
  // so a package that happens to have one keeps valid element order.
  const extLst = firstChildElement(root, 'extLst')
  for (const entry of entries) root.removeChild(entry)
  for (const oldIndex of usedOldIndices) root.insertBefore(entries[oldIndex], extLst)

  root.setAttribute('count', String(totalReferences))
  root.setAttribute('uniqueCount', String(usedOldIndices.length))
  zip.file(SHARED_STRINGS_PATH, serializeXmlPart(doc))
}

// ---------------------------------------------------------------------------
// A deleted sheet's own exclusively-owned parts (USR-17 follow-up)
// ---------------------------------------------------------------------------

/** Resolves a relationship `Target` against the directory of the part that owns the relationship (NOT `xl/workbook.xml.rels`'s own `xl/`-relative convention — see `resolveWorkbookRelTarget` for that one). */
function resolveRelativeTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const segments = baseDir.split('/').filter(Boolean)
  for (const part of target.split('/')) {
    if (part === '..') segments.pop()
    else if (part !== '.' && part !== '') segments.push(part)
  }
  return segments.join('/')
}

function relsPathForPart(partPath: string): string {
  const slash = partPath.lastIndexOf('/')
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`
}

function removePart(zip: JSZip, contentTypesDoc: XMLDocument | null, path: string): void {
  if (zip.file(path)) zip.remove(path)
  if (contentTypesDoc) removeContentTypeOverrideForPart(contentTypesDoc, path)
}

/** Relationship-`Type` suffixes exclusively owned by ONE worksheet — never shared with a surviving sheet, so safe to delete outright along with it. */
const SHEET_OWNED_REL_SUFFIXES = ['/table', '/comments', '/vmlDrawing', '/threadedComment', '/drawing']
/** One level further: parts a swept DRAWING itself exclusively owns. Does NOT include `/image` — a drawing's image could in principle be a part another, surviving sheet's drawing also targets (Excel does not guarantee one image part per drawing). An image is instead handled by `sweepOrphanedMedia`, which can PROVE exclusivity (or the lack of it) by reference-counting across the whole package, something this function's narrower one-drawing-at-a-time pass cannot do. */
const DRAWING_OWNED_REL_SUFFIXES = ['/chart']

/**
 * Removes every part that exists ONLY to serve the worksheet being deleted:
 * its table(s), its comments (legacy `<legacyDrawing>`/VML note shapes AND
 * modern threaded comments), and its own drawing part — one level further,
 * any chart THAT drawing embeds too. None of these can legitimately be
 * referenced by any other, surviving part of the package (a table and a
 * comments part belong to exactly one sheet by the OOXML schema itself; a
 * drawing and the charts inside it are likewise 1:1 with the sheet/drawing
 * that placed them). `xl/media/*` images are the one exception — see
 * `DRAWING_OWNED_REL_SUFFIXES` and `sweepOrphanedMedia`, which runs
 * separately (after every sheet delete has already gone through here) once
 * it can see the WHOLE package's surviving relationships.
 */
async function removeSheetOwnedParts(
  zip: JSZip,
  contentTypesDoc: XMLDocument | null,
  sheetDir: string,
  sheetRelsXml: string,
): Promise<void> {
  let relsDoc: XMLDocument
  try {
    relsDoc = parseXmlPart(sheetRelsXml)
  } catch {
    return
  }

  for (const rel of descendantElements(relsDoc, 'Relationship')) {
    const type = rel.getAttribute('Type') ?? ''
    const target = rel.getAttribute('Target')
    if (!target || !SHEET_OWNED_REL_SUFFIXES.some((suffix) => type.endsWith(suffix))) continue
    const partPath = resolveRelativeTarget(sheetDir, target)

    if (type.endsWith('/drawing')) {
      const drawingRelsXml = await zip.file(relsPathForPart(partPath))?.async('string')
      if (drawingRelsXml) {
        try {
          const drawingRelsDoc = parseXmlPart(drawingRelsXml)
          const drawingDir = partPath.slice(0, partPath.lastIndexOf('/'))
          for (const drawingRel of descendantElements(drawingRelsDoc, 'Relationship')) {
            const drawingRelType = drawingRel.getAttribute('Type') ?? ''
            const drawingTarget = drawingRel.getAttribute('Target')
            if (!drawingTarget || !DRAWING_OWNED_REL_SUFFIXES.some((suffix) => drawingRelType.endsWith(suffix))) continue
            const chartPath = resolveRelativeTarget(drawingDir, drawingTarget)
            removePart(zip, contentTypesDoc, chartPath)
            removePart(zip, contentTypesDoc, relsPathForPart(chartPath))
          }
        } catch {
          // Malformed drawing rels — still remove the drawing part itself below.
        }
      }
    }

    removePart(zip, contentTypesDoc, partPath)
    removePart(zip, contentTypesDoc, relsPathForPart(partPath))
  }
}

/**
 * Sweeps every `xl/media/*` part that is PROVABLY unreferenced by anything
 * left in the package, after `removeSheetOwnedParts` has already removed a
 * deleted sheet's own drawing (and that drawing's own rels file, which is
 * what used to point at the image). The module's own header used to
 * document leaving `xl/media/*` orphaned outright, because one drawing's
 * own rels file can never prove ITS image isn't also targeted by another,
 * surviving drawing — that reasoning is correct for a single drawing in
 * isolation, but doesn't hold once every remaining relationship in the
 * package is counted: OOXML has no part that is "provisionally" referenced,
 * so an image no surviving `.rels` file anywhere points at can only be
 * genuinely orphaned. This scans every `.rels` part still in the zip — a
 * drawing's own, a worksheet's (a sheet can target a background image
 * directly via `<picture r:id="...">`, with no drawing involved at all), a
 * chart's, a legacy VML drawing's, ... — not just the deleted sheet's own,
 * so an image still used by any of them is never touched. An external
 * (`TargetMode="External"`) relationship is skipped: its `Target` is a URL
 * or an out-of-package path, never an `xl/media/*` part.
 *
 * The caller only runs this when a sheet was actually deleted this save
 * (`anySheetDeleted`): an image the ORIGINAL file already left unreferenced,
 * with no sheet delete involved, is none of this save's business — this
 * module otherwise keeps everything it did not itself change untouched, and
 * a general "clean up anything unreferenced" pass would break that rule for
 * no reason tied to the edit being made.
 */
async function sweepOrphanedMedia(zip: JSZip, contentTypesDoc: XMLDocument | null): Promise<void> {
  const mediaPaths = Object.keys(zip.files).filter((path) => path.startsWith('xl/media/') && !zip.files[path].dir)
  if (mediaPaths.length === 0) return

  const referenced = new Set<string>()
  const relsPaths = Object.keys(zip.files).filter((path) => path.endsWith('.rels') && !zip.files[path].dir)
  for (const relsPath of relsPaths) {
    const xml = await zip.file(relsPath)?.async('string')
    if (!xml) continue
    let relsDoc: XMLDocument
    try {
      relsDoc = parseXmlPart(xml)
    } catch {
      continue // malformed rels — cannot prove anything it points at is unreferenced, so nothing here counts as "used"; harmless, since a WELL-FORMED rels file elsewhere still protects any image genuinely still in use
    }
    // A `.rels` part lives at `<dir>/_rels/<file>.rels`; its own `Target`s resolve against `<dir>`, the directory ABOVE `_rels`.
    const relsDir = relsPath.slice(0, relsPath.lastIndexOf('/_rels/'))
    for (const rel of descendantElements(relsDoc, 'Relationship')) {
      const target = rel.getAttribute('Target')
      if (!target || rel.getAttribute('TargetMode') === 'External') continue
      const resolved = resolveRelativeTarget(relsDir, target)
      if (resolved.startsWith('xl/media/')) referenced.add(resolved)
    }
  }

  for (const path of mediaPaths) {
    if (!referenced.has(path)) removePart(zip, contentTypesDoc, path)
  }
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

/**
 * Applies every defined-name edit implied by the sheet structure change:
 * `localSheetId` renumbering (or removal, for a name scoped to a deleted
 * sheet), and a rename/delete/re-anchor of the formula text itself via the
 * same reference-syntax engine a cell formula goes through
 * (`formulaRefs.ts`) — this covers a multi-area name
 * (`Sheet1!$A$1,Sheet1!$C$3`), a function-wrapped one
 * (`OFFSET(Sheet1!$A$1,...)`) and a whole-row/whole-column range (a
 * `Print_Titles` `$1:$1`), not just the single-area case. A name whose
 * formula collapses entirely to `#REF!` (its one area's sheet was deleted,
 * or its one range was fully consumed by a delete) is dropped outright, the
 * same way Excel drops a name that no longer resolves to anything at all —
 * a name with only SOME of its areas turned to `#REF!` is kept, `#REF!`
 * spliced in for just that area, matching what Excel itself does.
 */
function updateDefinedNames(
  workbookRoot: Element,
  changesByOriginalName: ReadonlyMap<string, SheetChange>,
  originalIndexToNewIndex: ReadonlyMap<number, number>,
  originalSheets: ReadonlyArray<OriginalSheetEntry>,
): void {
  const definedNames = firstChildElement(workbookRoot, 'definedNames')
  if (!definedNames) return

  for (const nameEl of childElements(definedNames, 'definedName')) {
    const localSheetId = nameEl.getAttribute('localSheetId')
    let ownChange: SheetChange | undefined
    if (localSheetId !== null) {
      const newIndex = originalIndexToNewIndex.get(Number(localSheetId))
      if (newIndex === undefined) {
        nameEl.parentNode?.removeChild(nameEl)
        continue
      }
      nameEl.setAttribute('localSheetId', String(newIndex))
      // A scoped name's OWN sheet, for re-anchoring an unqualified reference within it.
      ownChange = changesByOriginalName.get(originalSheets[Number(localSheetId)]?.name ?? '')
    }

    const original = nameEl.textContent ?? ''
    const rewritten = rewriteFormulaReferences(original, {
      changesByOriginalName,
      ownRowSources: ownChange?.rowSources,
      ownColSources: ownChange?.colSources,
      remapCoordinates: true, // a definedName's stored text is always in the ORIGINAL file's coordinate space — Atlas has no UI to edit one directly
    })
    if (rewritten === '#REF!') {
      nameEl.parentNode?.removeChild(nameEl)
      continue
    }
    if (rewritten !== original) nameEl.textContent = rewritten
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

  // SHEET-4/5 — loaded once for the whole save so every sheet's percent/date
  // auto-typing shares (and reuses) the same `cellXfs`/`numFmts` entries
  // instead of each sheet growing its own duplicate ones. Missing/unreadable
  // `xl/styles.xml` degrades to `null`, which turns the auto-typing off
  // entirely (`buildCell` still writes the parsed numeric value in that
  // case — see its own header — just never assigns a format for it).
  const stylesXml = await zip.file('xl/styles.xml')?.async('string')
  const stylesCtx = loadStylesContext(stylesXml)

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
  // relationship, their content-type override, and — one level further —
  // every part exclusively owned by them (tables, comments, drawings/charts;
  // see `removeSheetOwnedParts`).
  let anySheetDeleted = false
  for (const original of originalSheets) {
    if (originalIndexToNewIndex.has(original.originalIndex) || !original.path) continue
    anySheetDeleted = true
    const slash = original.path.lastIndexOf('/')
    const sheetDir = original.path.slice(0, slash)
    const sheetRelsPath = `${sheetDir}/_rels/${original.path.slice(slash + 1)}.rels`
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async('string')
    if (sheetRelsXml) await removeSheetOwnedParts(zip, contentTypesDoc, sheetDir, sheetRelsXml)

    zip.remove(original.path)
    if (zip.file(sheetRelsPath)) zip.remove(sheetRelsPath)
    for (const rel of descendantElements(relsDoc, 'Relationship')) {
      if (rel.getAttribute('Id') === original.rId) rel.parentNode?.removeChild(rel)
    }
    if (contentTypesDoc) removeContentTypeOverrideForPart(contentTypesDoc, original.path)
  }

  // ---- per-sheet content: cells, dimension, merges, columns, ranges, tables ----
  // Serialization is deferred until after `compactSharedStrings` below, which
  // needs every sheet's final `<c t="s">` cells still live as DOM nodes (not
  // yet flattened to XML text) to know which shared-string entries survived.
  const worksheetOutputs: { readonly partPath: string; readonly parsed: OriginalSheet }[] = []
  for (const [newIndex, sheet] of document.sheets.entries()) {
    const parsed = perSheetOriginal.get(newIndex)!
    const partPath = perSheetPartPath.get(newIndex)!

    rebuildSheetData(sheet, parsed, changesByOriginalName, stylesCtx)
    updateDimension(parsed.root, sheet)
    updateMergeCells(parsed.doc, parsed.root, sheet)
    updateColumns(parsed.doc, parsed.root, sheet)
    updateShiftedRanges(parsed.root, sheet)
    worksheetOutputs.push({ partPath, parsed })

    for (const [tableIndex, table] of (sheet.tables ?? []).entries()) {
      if (table.partPath === undefined) continue
      zip.file(table.partPath, rewriteTableXml(table, tableIndex + 1, tableHeaderNames(table, sheet.rows)))
    }
  }

  await compactSharedStrings(
    zip,
    worksheetOutputs.map((w) => w.parsed),
  )
  for (const { partPath, parsed } of worksheetOutputs) zip.file(partPath, serializeXmlPart(parsed.doc))

  updateDefinedNames(workbookRoot, changesByOriginalName, originalIndexToNewIndex, originalSheets)
  markRecalculateOnLoad(workbookDoc, workbookRoot)
  if (removeCalcChain(zip, relsDoc) && contentTypesDoc) removeContentTypeOverrideForPart(contentTypesDoc, 'xl/calcChain.xml')

  // Only relevant when a sheet delete just removed a drawing above — a
  // package's own, pre-existing unreferenced media (nothing to do with
  // THIS save) is left alone, matching this module's own "keeps the
  // ORIGINAL package and rewrites only what changed" rule; sweeping it
  // unconditionally would also punish an already-orphaned image that has
  // nothing to do with a sheet delete. Runs last, after every deleted
  // sheet's own drawing (and that drawing's rels) has already been removed
  // above, so this sees the package's FINAL set of surviving
  // relationships — anything in `xl/media/*` none of them still points at
  // is genuinely orphaned (see `sweepOrphanedMedia`).
  if (anySheetDeleted) await sweepOrphanedMedia(zip, contentTypesDoc)

  if (contentTypesDoc) zip.file(CONTENT_TYPES_PATH, serializeXmlPart(contentTypesDoc))
  zip.file(WORKBOOK_RELS_PATH, serializeXmlPart(relsDoc))
  zip.file('xl/workbook.xml', serializeXmlPart(workbookDoc))
  // SHEET-4/5 — only rewritten when a percent/date auto-type actually added
  // a `numFmt`/`cellXfs` entry this save; every other save leaves
  // `xl/styles.xml` byte-identical (still copied through untouched, since
  // it's outside this function's own zip writes).
  if (stylesCtx?.dirty) zip.file('xl/styles.xml', serializeXmlPart(stylesCtx.doc))

  await updateAppPropsTitles(
    zip,
    originalSheets.length,
    document.sheets.map((s) => s.name),
  )

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
