/**
 * Write-back (save) for the editable spreadsheet document model (wave 3).
 *
 * Builds a FRESH `XLSX.WorkBook` from a `SpreadsheetDocument` — it does not
 * hold onto or mutate the original parsed workbook. That means per-cell
 * styling (fonts, fills, borders, number formats) on cells the user never
 * touched is NOT preserved across a save: SheetJS's `aoa_to_sheet` only ever
 * produces plain value cells, and `useSpreadsheetWorkbook`'s worker-friendly
 * parse path (`parseWorkbookBuffer`) never retains the original `WorkBook`
 * object for exactly the reason a large file is parsed off-thread in the
 * first place (a `WorkBook` with full per-cell style refs is a much heavier,
 * less structured-clone-friendly object than the flat `ParsedSheet[]` this
 * app actually needs for rendering). What DOES survive a save: every cell's
 * value/formula, merged ranges, column widths, and row heights (T4's own
 * fidelity fixes) — everything the editable document model itself tracks.
 * This is a deliberate, documented scope cut, not an oversight.
 *
 * Cell typing on write: a cell's current display text is re-parsed as a
 * number when the whole (trimmed) string parses as one via `Number(...)`
 * (matching common spreadsheet auto-typing behavior), otherwise it's written
 * as a plain string; an empty string produces no cell at all (an genuinely
 * blank cell, not an empty-string one). This is a real, documented fidelity
 * loss for a heavily-formatted numeric display (e.g. `"$1,234.50"` re-reads
 * as text, not the original currency-formatted number 1234.5) since the
 * editable model only ever carries display TEXT, not the original SheetJS
 * cell object/number-format code.
 *
 * Formula cells: `f` is set to the formula text (no leading `=`) and `v`/`t`
 * are derived from the SAME re-typing rule above, from the cell's current
 * (already-recalculated, see `spreadsheetDocument.ts`) display text — i.e.
 * the formula is written alongside its best-known current value rather than
 * with no cached value at all. A cell left with literally no cached value
 * (`t` set, `v` omitted) round-trips through SheetJS's own XML writer as a
 * declared-but-unpopulated `t="e"` (error) cell with no `<v>` at all — which
 * a real spreadsheet application reads as "recalculate me", but displays
 * with an error indicator until it does, a worse first impression than a
 * plausible placeholder value. Writing the current display value avoids
 * that: Excel (or Atlas, reopening its own file) shows a correct-looking
 * value immediately, and still recalculates from `f` wherever it can.
 *
 * Per-format limitation, verified directly against this build: writing
 * `bookType: 'xls'` (BIFF8) or `'xlsb'` (BIFF12) never serializes a cell's
 * `.f` at all — only its cached value survives. A formula saved to `.xls`/
 * `.xlsb` therefore round-trips as a plain value, not a live formula, on
 * the very next open (in Atlas or anywhere else); the XML-based
 * xlsx/xlsm/ods/fods writers do not have this limitation.
 */
import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import Papa from 'papaparse'

import type { EditableSheet, SpreadsheetDocument } from './spreadsheetDocument'
import { graftTables, tableHeaderNames } from './spreadsheetTables'

/** Extension -> SheetJS write `bookType`. `.xltx`/`.xltm` (templates) have no distinct write format in this build, so they're written as plain `.xlsx`/`.xlsm` bytes — the OOXML content is otherwise identical. */
const EXTENSION_TO_BOOK_TYPE: Readonly<Record<string, XLSX.BookType>> = {
  xlsx: 'xlsx',
  xlsm: 'xlsm',
  xltx: 'xlsx',
  xltm: 'xlsm',
  xlsb: 'xlsb',
  xls: 'xls',
  ods: 'ods',
  fods: 'fods',
}

/** Resolves the SheetJS `bookType` to write for a given file extension (no leading dot, case-insensitive). Falls back to `'xlsx'` for anything unrecognized. */
export function bookTypeForExtension(ext: string): XLSX.BookType {
  return EXTENSION_TO_BOOK_TYPE[ext.toLowerCase()] ?? 'xlsx'
}

type TypedCellValue = { readonly t: 'n'; readonly v: number } | { readonly t: 's'; readonly v: string }

/**
 * Re-types a cell's current display text: a whole numeric string becomes a
 * number cell, anything else (including an empty string, handled by the
 * caller) a string cell.
 *
 * SHEET-3 — a leading `'` (Excel's own "force text" quote prefix) always
 * makes the value text, stripped of the quote itself (`'007` saves as the
 * text `"007"`, never the number `7`), matching `xlsxPassthrough.ts`'s own
 * `buildCell`. Only that half of SHEET-3 applies here: this writer builds a
 * brand-new workbook with no per-cell styles at all (see this module's own
 * header), so there is no existing `@` (text-formatted) cell to check for
 * the other half. `forFormula` skips the quote-prefix rule for a formula
 * cell's own (computed, never user-typed) display text.
 */
function typeCellText(text: string, forFormula = false): TypedCellValue {
  if (!forFormula && text.startsWith("'")) {
    return { t: 's', v: text.slice(1) }
  }
  const trimmed = text.trim()
  if (trimmed !== '' && Number.isFinite(Number(trimmed))) {
    return { t: 'n', v: Number(trimmed) }
  }
  return { t: 's', v: text }
}

function buildWorksheet(sheet: EditableSheet): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {}
  const rowCount = sheet.rows.length
  const colCount = sheet.colCount

  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const text = sheet.rows[r]?.[c] ?? ''
      const formula = sheet.formulas[r]?.[c]
      if (text === '' && formula === undefined) continue // a genuinely blank cell — no entry at all

      const addr = XLSX.utils.encode_cell({ r, c })
      const typed = typeCellText(text, formula !== undefined)
      ws[addr] = formula !== undefined ? { ...typed, f: formula } : typed
    }
  }

  // USR-17: an Excel table's header cells must be text equal to its column names.
  for (const table of sheet.tables ?? []) {
    if (!table.headerRow) continue
    tableHeaderNames(table, sheet.rows).forEach((name, i) => {
      ws[XLSX.utils.encode_cell({ r: table.r0, c: table.c0 + i })] = { t: 's', v: name }
    })
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(rowCount - 1, 0), c: Math.max(colCount - 1, 0) },
  })

  if (sheet.merges.length > 0) {
    ws['!merges'] = sheet.merges.map((m) => ({ s: { r: m.r0, c: m.c0 }, e: { r: m.r1, c: m.c1 } }))
  }
  if (sheet.colWidthsPx.some((w) => w !== undefined)) {
    ws['!cols'] = sheet.colWidthsPx.map((w) => (w !== undefined ? { wpx: w } : {}))
  }
  if (sheet.rowHeightsPx.some((h) => h !== undefined)) {
    ws['!rows'] = sheet.rowHeightsPx.map((h) => (h !== undefined ? { hpx: h } : {}))
  }

  return ws
}

/** Builds a fresh `XLSX.WorkBook` from the editable document — see module header for what is/isn't preserved. */
export function buildWorkbook(doc: SpreadsheetDocument): XLSX.WorkBook {
  const wb = XLSX.utils.book_new()
  for (const sheet of doc.sheets) {
    XLSX.utils.book_append_sheet(wb, buildWorksheet(sheet), sheet.name)
  }
  if (doc.sheets.some((s) => s.hidden)) {
    wb.Workbook = { Sheets: doc.sheets.map((s) => ({ Hidden: s.hidden ? 1 : 0 })) }
  }
  return wb
}

/**
 * Serializes the document to bytes for the given `bookType` (see
 * `bookTypeForExtension`). `XLSX.write`'s `type: 'array'` actually returns a
 * plain `ArrayBuffer` (not a `Uint8Array`, despite the option's name) — wrapped
 * here so callers get the `Uint8Array` `window.electronAPI.saveBinaryFile`
 * actually expects, and so the return type is honest.
 */
export function writeWorkbookBytes(doc: SpreadsheetDocument, bookType: XLSX.BookType): Uint8Array {
  const wb = buildWorkbook(doc)
  const buffer = XLSX.write(wb, { type: 'array', bookType }) as ArrayBuffer
  return new Uint8Array(buffer)
}

/** Book types whose package is OOXML, where Excel tables can be grafted back in (see `spreadsheetTables.ts`). */
const OOXML_BOOK_TYPES: ReadonlySet<XLSX.BookType> = new Set<XLSX.BookType>(['xlsx', 'xlsm'])

/**
 * SheetJS's `bookType: 'ods'` writer produces a package whose zip entry
 * order is `META-INF/manifest.xml, meta.xml, mimetype, styles.xml,
 * content.xml, manifest.rdf` (verified directly against this build's
 * `xlsx` version) — `mimetype` third, not first. The ODF Package spec
 * (OASIS ODF 1.2 part 3, §2.2) requires `mimetype` to be the FIRST entry in
 * the zip and stored uncompressed, or a conforming reader may refuse the
 * file outright rather than fall back to sniffing content; the same rule
 * `office/officePackage.ts`'s `writeOfficePackage` already honors for ODP
 * saves. Every `.ods` Atlas has ever saved (via `writeWorkbookBytesWithTables`,
 * the only path for `.ods` — `writeWorkbookThroughOriginal` only patches
 * `.xlsx`/`.xlsm`) shipped with this defect until this fix. Repacks the
 * archive with `mimetype` moved to the front and stored; every other part
 * keeps its bytes, just re-deflated.
 */
async function fixOdsPackaging(bytes: Uint8Array): Promise<Uint8Array> {
  const original = await JSZip.loadAsync(bytes)
  const mimetype = await original.file('mimetype')?.async('uint8array')
  if (mimetype === undefined) return bytes // defensive: nothing to reorder if SheetJS ever stops emitting one

  const repacked = new JSZip()
  repacked.file('mimetype', mimetype, { compression: 'STORE' })
  for (const [path, entry] of Object.entries(original.files)) {
    if (path === 'mimetype' || entry.dir) continue
    repacked.file(path, await entry.async('uint8array'))
  }
  return repacked.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** `writeWorkbookBytes` plus the sheets' Excel tables for OOXML targets (USR-17); other formats have no table concept. */
export async function writeWorkbookBytesWithTables(
  doc: SpreadsheetDocument,
  bookType: XLSX.BookType,
): Promise<Uint8Array> {
  const bytes = writeWorkbookBytes(doc, bookType)
  if (bookType === 'ods') return fixOdsPackaging(bytes)
  return OOXML_BOOK_TYPES.has(bookType) ? graftTables(bytes, doc.sheets) : bytes
}

/**
 * Serializes a single-sheet document (CsvViewer's use case) as delimited
 * text via Papa Parse's own writer, so quoting/escaping matches exactly what
 * `csvParse.ts` already reads back. `doc.sheets[0]` is used unconditionally
 * — CSV/TSV have no multi-sheet concept.
 *
 * Encoding/BOM (documented limitation): the delimiter is genuinely
 * preserved (the caller passes the same one the file was parsed with — see
 * `CsvViewer`'s `defaultSaveTarget`), but the *original file's byte-level*
 * encoding and BOM presence are not — `electron/lib/textDecoding.cjs`
 * normalizes every text-class file to a plain UTF-8 JS string at load time
 * (stripping any BOM in the process) with no metadata carried forward
 * recording what the original encoding/BOM actually was, so there's nothing
 * for this function to preserve even in principle. The renderer's own save
 * path (`window.electronAPI.saveFile`) always writes this string back as
 * plain UTF-8 with no BOM.
 */
export function documentToDelimitedText(doc: SpreadsheetDocument, delimiter: string): string {
  const sheet = doc.sheets[0]
  if (!sheet) return ''
  return Papa.unparse(
    sheet.rows.map((row) => [...row]),
    { delimiter },
  )
}
