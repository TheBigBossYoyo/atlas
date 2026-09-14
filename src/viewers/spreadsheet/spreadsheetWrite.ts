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
import Papa from 'papaparse'

import type { EditableSheet, SpreadsheetDocument } from './spreadsheetDocument'

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

/** Re-types a cell's current display text: a whole numeric string becomes a number cell, anything else (including an empty string, handled by the caller) a string cell. */
function typeCellText(text: string): TypedCellValue {
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
      const typed = typeCellText(text)
      ws[addr] = formula !== undefined ? { ...typed, f: formula } : typed
    }
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

/**
 * Serializes a single-sheet document (CsvViewer's use case) as delimited
 * text via Papa Parse's own writer, so quoting/escaping matches exactly what
 * `csvParse.ts` already reads back. `doc.sheets[0]` is used unconditionally
 * — CSV/TSV have no multi-sheet concept.
 */
export function documentToDelimitedText(doc: SpreadsheetDocument, delimiter: string): string {
  const sheet = doc.sheets[0]
  if (!sheet) return ''
  return Papa.unparse(
    sheet.rows.map((row) => [...row]),
    { delimiter },
  )
}
