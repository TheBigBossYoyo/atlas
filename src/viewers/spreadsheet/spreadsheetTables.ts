/**
 * Excel tables (ListObjects) — USR-17.
 *
 * SheetJS neither parses nor writes `xl/tables/*.xml`, and Atlas saves a
 * FRESH workbook (see `spreadsheetWrite.ts`), so before this module every
 * save silently turned an Excel table back into a plain cell range. This
 * module:
 *
 *  1. reads each sheet's table parts from the original OOXML zip at load
 *     (`readSheetTables`, same zip/relationship walk as `spreadsheetPanes.ts`),
 *  2. keeps each table's range in step with row/column inserts and deletes
 *     (`tablesAfter*`, called by `spreadsheetDocument.ts` next to its merge
 *     bookkeeping),
 *  3. grafts the tables back into the written package on save
 *     (`graftTables`), rewriting `ref`, the autoFilter range and the column
 *     list, and naming columns from the header cells exactly as Excel
 *     requires (`tableHeaderNames`).
 *
 * Scope: OOXML (.xlsx/.xlsm) only. Structured references inside formulas
 * (`Table1[Qty]`) are kept as formula text but not evaluated by Atlas's own
 * formula engine; filter/sort state is dropped on save (column indexes may
 * have moved) while the table itself, its style and its totals row survive.
 */
import { readSheetParts } from './spreadsheetPanes'
import { loadWorkbookZip } from './spreadsheetZipBudget'

export type SheetTable = {
  /** `displayName` — what formulas and Excel's UI refer to. */
  readonly name: string
  readonly r0: number
  readonly c0: number
  readonly r1: number
  readonly c1: number
  readonly headerRow: boolean
  readonly totalsRow: boolean
  /** Original table part XML, rewritten and grafted back on save. */
  readonly xml: string
  /** Where that part lives in the original package — set when the table was read from a file. */
  readonly partPath?: string
  /** Per current column: index of the original `<tableColumn>` it came from, or `null` for a column inserted in Atlas. */
  readonly columnSources: ReadonlyArray<number | null>
}

const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const TABLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

// ---------------------------------------------------------------------------
// Cell references
// ---------------------------------------------------------------------------

function decodeCell(ref: string): { r: number; c: number } | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref.trim())
  if (!match) return null
  let c = 0
  for (const ch of match[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64)
  return { r: Number(match[2]) - 1, c: c - 1 }
}

export function encodeCol(col: number): string {
  let title = ''
  let n = col
  while (n >= 0) {
    title = String.fromCharCode(65 + (n % 26)) + title
    n = Math.floor(n / 26) - 1
  }
  return title
}

function encodeRange(r0: number, c0: number, r1: number, c1: number): string {
  return `${encodeCol(c0)}${r0 + 1}:${encodeCol(c1)}${r1 + 1}`
}

export function decodeRange(ref: string): { r0: number; c0: number; r1: number; c1: number } | null {
  const [start, end = start] = ref.split(':')
  const s = decodeCell(start)
  const e = decodeCell(end)
  if (!s || !e) return null
  return { r0: Math.min(s.r, e.r), c0: Math.min(s.c, e.c), r1: Math.max(s.r, e.r), c1: Math.max(s.c, e.c) }
}

/** Resolves a relationship `Target` against the directory of the part that owns the relationship. */
function resolvePartPath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const segments = baseDir.split('/').filter(Boolean)
  for (const part of target.split('/')) {
    if (part === '..') segments.pop()
    else if (part !== '.' && part !== '') segments.push(part)
  }
  return segments.join('/')
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/')
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`
}

// ---------------------------------------------------------------------------
// Load
// ---------------------------------------------------------------------------

export function parseTableXml(xml: string): SheetTable | null {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) return null
  const root = doc.documentElement
  const range = decodeRange(root.getAttribute('ref') ?? '')
  if (!range) return null
  const columnCount = doc.getElementsByTagName('tableColumn').length
  const width = range.c1 - range.c0 + 1
  return {
    name: root.getAttribute('displayName') ?? root.getAttribute('name') ?? 'Table',
    ...range,
    headerRow: root.getAttribute('headerRowCount') !== '0',
    totalsRow: Number(root.getAttribute('totalsRowCount') ?? '0') > 0,
    xml,
    columnSources: Array.from({ length: width }, (_, i) => (i < columnCount ? i : null)),
  }
}

/** Worksheet part paths in workbook order (USR-17 save-through-original); `[]` for a non-OOXML file. */
export async function readSheetPartPaths(buffer: ArrayBuffer): Promise<string[]> {
  try {
    // USR-17: same declared-size zip-bomb budget as `spreadsheetPanes.ts` — see `spreadsheetZipBudget.ts`.
    const zip = await loadWorkbookZip(buffer)
    const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
    if (!workbookXml) return []
    const relsXml = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? null
    return readSheetParts(workbookXml, relsXml).map((part) => part.path)
  } catch {
    return []
  }
}

/** Sheet name -> tables, read from an OOXML workbook buffer. Resolves `{}` (never rejects) for any other format. */
export async function readSheetTables(buffer: ArrayBuffer): Promise<Readonly<Record<string, ReadonlyArray<SheetTable>>>> {
  try {
    // USR-17: same declared-size zip-bomb budget as `spreadsheetPanes.ts` — see `spreadsheetZipBudget.ts`.
    const zip = await loadWorkbookZip(buffer)
    const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
    if (!workbookXml) return {}
    const relsXml = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? null
    const result: Record<string, SheetTable[]> = {}

    for (const { name, path } of readSheetParts(workbookXml, relsXml)) {
      const sheetRels = await zip.file(relsPathFor(path))?.async('string')
      if (!sheetRels) continue
      const relsDoc = new DOMParser().parseFromString(sheetRels, 'application/xml')
      const sheetDir = path.slice(0, path.lastIndexOf('/'))
      for (const rel of Array.from(relsDoc.getElementsByTagName('Relationship'))) {
        const target = rel.getAttribute('Target')
        if (rel.getAttribute('Type') !== TABLE_REL_TYPE || !target) continue
        const partPath = resolvePartPath(sheetDir, target)
        const tableXml = await zip.file(partPath)?.async('string')
        const table = tableXml ? parseTableXml(tableXml) : null
        if (table) (result[name] ??= []).push({ ...table, partPath })
      }
    }
    return result
  } catch {
    return {}
  }
}

// ---------------------------------------------------------------------------
// Edits (pure; mirror spreadsheetDocument.ts's merge bookkeeping)
// ---------------------------------------------------------------------------

/** Smallest table Excel accepts: the header row (if any) plus one data row. */
function isViable(table: SheetTable): boolean {
  const minRows = (table.headerRow ? 1 : 0) + (table.totalsRow ? 1 : 0) + 1
  return table.r1 - table.r0 + 1 >= minRows && table.c1 >= table.c0
}

export function tablesAfterRowInsert(tables: ReadonlyArray<SheetTable>, at: number): ReadonlyArray<SheetTable> {
  return tables.map((t) => {
    if (at <= t.r0) return { ...t, r0: t.r0 + 1, r1: t.r1 + 1 }
    if (at <= t.r1) return { ...t, r1: t.r1 + 1 }
    return t
  })
}

export function tablesAfterRowDelete(tables: ReadonlyArray<SheetTable>, at: number): ReadonlyArray<SheetTable> {
  return tables.flatMap((t) => {
    if (at < t.r0) return [{ ...t, r0: t.r0 - 1, r1: t.r1 - 1 }]
    if (at > t.r1) return [t]
    // Deleting the header row un-tables the range (the data itself stays).
    if (t.headerRow && at === t.r0) return []
    const next = { ...t, r1: t.r1 - 1, totalsRow: t.totalsRow && at !== t.r1 }
    return isViable(next) ? [next] : []
  })
}

export function tablesAfterColumnInsert(tables: ReadonlyArray<SheetTable>, at: number): ReadonlyArray<SheetTable> {
  return tables.map((t) => {
    if (at <= t.c0) return { ...t, c0: t.c0 + 1, c1: t.c1 + 1 }
    if (at > t.c1) return t
    const columnSources = [...t.columnSources]
    columnSources.splice(at - t.c0, 0, null)
    return { ...t, c1: t.c1 + 1, columnSources }
  })
}

export function tablesAfterColumnDelete(tables: ReadonlyArray<SheetTable>, at: number): ReadonlyArray<SheetTable> {
  return tables.flatMap((t) => {
    if (at < t.c0) return [{ ...t, c0: t.c0 - 1, c1: t.c1 - 1 }]
    if (at > t.c1) return [t]
    if (t.c0 === t.c1) return []
    const columnSources = t.columnSources.filter((_, i) => i !== at - t.c0)
    return [{ ...t, c1: t.c1 - 1, columnSources }]
  })
}

/** True when (row, col) is a header cell of one of `tables`. */
export function isTableHeaderCell(tables: ReadonlyArray<SheetTable> | undefined, row: number, col: number): boolean {
  return (tables ?? []).some((t) => t.headerRow && row === t.r0 && col >= t.c0 && col <= t.c1)
}

// ---------------------------------------------------------------------------
// Save
// ---------------------------------------------------------------------------

/**
 * Column names for a table, taken from its header cells as Excel requires
 * (header cell text must equal the column name): blank headers become
 * `ColumnN`, and case-insensitive duplicates get a numeric suffix.
 */
export function tableHeaderNames(table: SheetTable, rows: ReadonlyArray<ReadonlyArray<string>>): string[] {
  const used = new Set<string>()
  const names: string[] = []
  for (let c = table.c0; c <= table.c1; c++) {
    const index = c - table.c0
    const raw = table.headerRow ? (rows[table.r0]?.[c] ?? '').trim() : ''
    const base = raw || `Column${index + 1}`
    let name = base
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base}${n}`
    used.add(name.toLowerCase())
    names.push(name)
  }
  return names
}

function removeAll(root: Element, tagName: string): void {
  for (const el of Array.from(root.getElementsByTagName(tagName))) el.parentNode?.removeChild(el)
}

/** Rewrites a table part for the table's current range and column list. */
export function rewriteTableXml(table: SheetTable, tableId: number, names: ReadonlyArray<string>): string {
  const doc = new DOMParser().parseFromString(table.xml, 'application/xml')
  const root = doc.documentElement
  const ns = root.namespaceURI
  root.setAttribute('id', String(tableId))
  root.setAttribute('ref', encodeRange(table.r0, table.c0, table.r1, table.c1))
  if (!table.totalsRow) root.removeAttribute('totalsRowCount')

  // Column indexes may have moved, so filter/sort state cannot be trusted.
  removeAll(root, 'sortState')
  removeAll(root, 'filterColumn')
  for (const autoFilter of Array.from(root.getElementsByTagName('autoFilter'))) {
    const lastDataRow = table.r1 - (table.totalsRow ? 1 : 0)
    autoFilter.setAttribute('ref', encodeRange(table.r0, table.c0, lastDataRow, table.c1))
  }

  const container = root.getElementsByTagName('tableColumns')[0]
  if (container) {
    const originals = Array.from(container.getElementsByTagName('tableColumn'))
    let nextId = originals.reduce((max, el) => Math.max(max, Number(el.getAttribute('id')) || 0), 0) + 1
    while (container.firstChild) container.removeChild(container.firstChild)
    table.columnSources.forEach((source, i) => {
      const original = source !== null ? originals[source] : undefined
      const column = original ?? doc.createElementNS(ns, 'tableColumn')
      if (!original) column.setAttribute('id', String(nextId++))
      column.setAttribute('name', names[i] ?? `Column${i + 1}`)
      container.appendChild(column)
    })
    container.setAttribute('count', String(table.columnSources.length))
  }

  const serialized = new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '')
  return XML_DECLARATION + serialized
}

type TableSheetInput = {
  readonly name: string
  readonly rows: ReadonlyArray<ReadonlyArray<string>>
  readonly tables?: ReadonlyArray<SheetTable>
}

function insertBeforeClosing(xml: string, closingTag: string, fragment: string, preferBefore?: string): string {
  const preferred = preferBefore ? xml.lastIndexOf(preferBefore) : -1
  const at = preferred >= 0 ? preferred : xml.lastIndexOf(closingTag)
  return at < 0 ? xml : `${xml.slice(0, at)}${fragment}${xml.slice(at)}`
}

/**
 * Adds every sheet's tables to a freshly written OOXML package (SheetJS
 * output). Returns the input bytes untouched when there is nothing to graft.
 */
export async function graftTables(bytes: Uint8Array, sheets: ReadonlyArray<TableSheetInput>): Promise<Uint8Array> {
  if (!sheets.some((s) => (s.tables?.length ?? 0) > 0)) return bytes

  // USR-17: `bytes` is Atlas's own just-written SheetJS output, not user
  // input, but routing it through the same declared-size guard as the two
  // readers above costs nothing for a normal file and keeps one consistent
  // check across every `JSZip.loadAsync` in this module.
  const zip = await loadWorkbookZip(bytes)
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
  if (!workbookXml) return bytes
  const parts = readSheetParts(workbookXml, (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? null)

  let tableNo = 0
  const overrides: string[] = []
  for (const sheet of sheets) {
    const part = parts.find((p) => p.name === sheet.name)
    const sheetEntry = part ? zip.file(part.path) : null
    if (!part || !sheetEntry || !sheet.tables?.length) continue

    const relsPath = relsPathFor(part.path)
    let relsXml =
      (await zip.file(relsPath)?.async('string')) ??
      `${XML_DECLARATION}<Relationships xmlns="${PACKAGE_RELS_NS}"></Relationships>`
    const tableParts: string[] = []
    for (const table of sheet.tables) {
      tableNo += 1
      const rId = `rIdAtlasTable${tableNo}`
      zip.file(`xl/tables/table${tableNo}.xml`, rewriteTableXml(table, tableNo, tableHeaderNames(table, sheet.rows)))
      relsXml = insertBeforeClosing(
        relsXml,
        '</Relationships>',
        `<Relationship Id="${rId}" Type="${TABLE_REL_TYPE}" Target="../tables/table${tableNo}.xml"/>`,
      )
      tableParts.push(`<tablePart r:id="${rId}"/>`)
      overrides.push(`<Override PartName="/xl/tables/table${tableNo}.xml" ContentType="${TABLE_CONTENT_TYPE}"/>`)
    }
    zip.file(relsPath, relsXml)

    let sheetXml = await sheetEntry.async('string')
    if (!/<worksheet[^>]*xmlns:r=/.test(sheetXml)) {
      sheetXml = sheetXml.replace(
        '<worksheet',
        '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      )
    }
    const fragment = `<tableParts count="${tableParts.length}">${tableParts.join('')}</tableParts>`
    zip.file(part.path, insertBeforeClosing(sheetXml, '</worksheet>', fragment, '<extLst'))
  }

  const contentTypes = await zip.file('[Content_Types].xml')?.async('string')
  if (contentTypes) zip.file('[Content_Types].xml', insertBeforeClosing(contentTypes, '</Types>', overrides.join('')))

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
