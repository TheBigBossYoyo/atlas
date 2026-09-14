/**
 * Editable spreadsheet document model — spreadsheet editing + save (wave 3).
 *
 * A pure, immutable data model (every function returns a new
 * `SpreadsheetDocument`, never mutates its input) sitting between the
 * parsed-from-file `ParsedSheet[]` (`shared/spreadsheetGrid.ts`, read-only)
 * and the glide-data-grid-facing rendering hook. `useSpreadsheetEditor.ts`
 * wraps these operations with undo/redo history and wires them to a live
 * viewer; this module has no React/DOM dependency so it's cheap to unit-test
 * directly.
 *
 * Both SpreadsheetViewer (xlsx/ods, multi-sheet) and CsvViewer (csv/tsv,
 * always exactly one sheet) share this same model — CsvViewer just never
 * exposes the sheet-add/rename/delete operations in its UI.
 *
 * Formula recalculation (documented scope cut, matching the style of
 * `spreadsheetFormula.ts`'s own header): there is no dependency graph. Any
 * edit that could change a formula's result (a value edit, a formula edit, a
 * row/column insert or delete, a paste) triggers exactly one row-major
 * left-to-right recalculation pass over every formula cell in that sheet.
 * A formula that references a cell computed *later* in that same pass still
 * sees that cell's *previous* value, not its freshly-recalculated one — a
 * real spreadsheet's dependency-graph engine would converge in one pass
 * regardless of layout order; this one does not. Circular references are
 * not detected; they simply resolve to whatever the referenced cell's
 * current display text already is (no infinite loop, since a single pass
 * never re-visits a cell). Inserting/deleting a row or column also does NOT
 * shift other cells' formula references (e.g. deleting row 2 leaves a
 * `=A3` elsewhere still reading literal row 3) — implementing Excel's
 * reference-shifting semantics is a substantial, separate undertaking, out
 * of scope here.
 */
import { evaluateFormula } from './spreadsheetFormula'
import type { CellLookup } from './spreadsheetFormula'
import type { FrozenPanes, MergeRange, ParsedSheet } from '../shared/spreadsheetGrid'

export type EditableSheet = {
  readonly name: string
  readonly hidden: boolean
  readonly rows: ReadonlyArray<ReadonlyArray<string>>
  readonly formulas: ReadonlyArray<ReadonlyArray<string | undefined>>
  readonly colCount: number
  readonly merges: ReadonlyArray<MergeRange>
  readonly colWidthsPx: ReadonlyArray<number | undefined>
  readonly rowHeightsPx: ReadonlyArray<number | undefined>
  readonly freeze?: FrozenPanes
}

export type SpreadsheetDocument = {
  readonly sheets: ReadonlyArray<EditableSheet>
}

/** Rows/columns a freshly-added sheet starts with — a usable, small default grid (not Excel's impractical full extent). */
const NEW_SHEET_ROW_COUNT = 20
const NEW_SHEET_COL_COUNT = 10

function emptyRow(colCount: number): string[] {
  return new Array<string>(colCount).fill('')
}

function emptyFormulaRow(colCount: number): (string | undefined)[] {
  return new Array<string | undefined>(colCount).fill(undefined)
}

/**
 * Builds the initial editable document from parsed (read-only) sheets — the
 * load-time entry point. Runs `recalculateSheet` once up front so a formula
 * cell our evaluator can't handle but the *file* already cached a value for
 * (the overwhelmingly common case for anything beyond basic arithmetic —
 * `VLOOKUP`, `IF`, text functions, ...) still shows that cached value
 * (DAT-04 fidelity), while a formula our evaluator DOES support starts out
 * "live" from the first render, matching what happens after any later edit.
 */
export function createDocument(parsedSheets: ReadonlyArray<ParsedSheet>): SpreadsheetDocument {
  return {
    sheets: parsedSheets.map((sheet) =>
      recalculateSheet({
        name: sheet.name,
        hidden: sheet.hidden,
        rows: sheet.grid.rows,
        formulas: sheet.grid.formulas,
        colCount: sheet.grid.colCount,
        merges: sheet.grid.merges,
        colWidthsPx: sheet.grid.colWidthsPx,
        rowHeightsPx: sheet.grid.rowHeightsPx,
        ...(sheet.freeze ? { freeze: sheet.freeze } : {}),
      }),
    ),
  }
}

/** Builds a single-sheet document directly from CSV/TSV rows (no sheet/merge/width concept). */
export function createDocumentFromRows(rows: ReadonlyArray<ReadonlyArray<string>>, colCount: number): SpreadsheetDocument {
  return {
    sheets: [
      {
        name: 'Sheet1',
        hidden: false,
        rows,
        formulas: rows.map((row) => emptyFormulaRow(row.length)),
        colCount,
        merges: [],
        colWidthsPx: [],
        rowHeightsPx: [],
      },
    ],
  }
}

function replaceSheet(doc: SpreadsheetDocument, sheetIndex: number, sheet: EditableSheet): SpreadsheetDocument {
  return { sheets: doc.sheets.map((s, i) => (i === sheetIndex ? sheet : s)) }
}

/**
 * Recomputes every formula cell's display text in one row-major pass (see
 * module header for the single-pass/no-dependency-graph limitation).
 *
 * A formula our evaluator can compute (`result.ok`) always gets the live,
 * freshly-computed text — including on every subsequent edit, so a `=SUM(...)`
 * cell stays "live" the way a real spreadsheet's would. A formula our
 * evaluator can't handle at all (an unsupported function, a malformed
 * expression) instead KEEPS whatever display text the cell already has
 * (typically the original file's own cached value, e.g. a `VLOOKUP` result
 * Excel computed and Atlas's evaluator was never going to reproduce) rather
 * than clobbering it on every unrelated edit elsewhere in the sheet — it
 * only falls back to the literal `=<formula>` text when there is no cached
 * text at all (a brand-new formula, or one loaded from a file Atlas itself
 * saved with the cached value intentionally left empty — see
 * `spreadsheetWrite.ts`).
 *
 * Performance: `createDocument` calls this once per sheet on EVERY load,
 * including the 100k-row perf-guarantee fixtures (T2/DAT-07) that have no
 * formulas at all — so the common "no formulas in this sheet" case must not
 * pay for an O(rows) copy of every row it never needed. `rows` therefore
 * stays `null` (and the original `sheet` is returned unchanged) until a
 * formula cell is actually found; only then is the copy-on-write triggered.
 */
export function recalculateSheet(sheet: EditableSheet): EditableSheet {
  let rows: string[][] | null = null
  const lookup: CellLookup = (row, col) => (rows ?? sheet.rows)[row]?.[col] ?? ''

  for (let r = 0; r < sheet.formulas.length; r++) {
    const formulaRow = sheet.formulas[r]
    for (let c = 0; c < formulaRow.length; c++) {
      const formula = formulaRow[c]
      if (formula === undefined) continue
      if (rows === null) rows = sheet.rows.map((row) => [...row])

      const result = evaluateFormula(formula, lookup)
      if (result.ok) {
        rows[r][c] = result.text
      } else if (rows[r][c] === '') {
        rows[r][c] = `=${formula}`
      }
    }
  }

  return rows === null ? sheet : { ...sheet, rows }
}

/**
 * Sets one cell's raw user input. Input starting with `=` (and longer than
 * just `=`) is stored as a formula (`formulas[row][col]`, text after the
 * leading `=`) with its display text computed by `recalculateSheet`;
 * anything else is stored as a plain value with no formula. Growing the
 * grid by editing past its current bounds is not supported here — use
 * `insertRowAt`/`insertColumnAt` first.
 */
export function setCellValue(
  doc: SpreadsheetDocument,
  sheetIndex: number,
  row: number,
  col: number,
  rawInput: string,
): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet || row < 0 || row >= sheet.rows.length || col < 0 || col >= sheet.colCount) {
    return doc
  }

  const isFormula = rawInput.startsWith('=') && rawInput.length > 1
  const rows = sheet.rows.map((r) => [...r])
  const formulas = sheet.formulas.map((r) => [...r])

  formulas[row][col] = isFormula ? rawInput.slice(1) : undefined
  // For a formula cell this is immediately overwritten by `recalculateSheet`
  // below with the evaluated (or literal-formula-fallback) text; for a plain
  // value cell this IS the final stored display text.
  rows[row][col] = rawInput

  const updated = recalculateSheet({ ...sheet, rows, formulas })
  return replaceSheet(doc, sheetIndex, updated)
}

/** Inserts one empty row at `atIndex` (0-based; may equal `rows.length` to append at the end). */
export function insertRowAt(doc: SpreadsheetDocument, sheetIndex: number, atIndex: number): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet) return doc

  const rows = [...sheet.rows]
  rows.splice(atIndex, 0, emptyRow(sheet.colCount))
  const formulas = [...sheet.formulas]
  formulas.splice(atIndex, 0, emptyFormulaRow(sheet.colCount))
  const rowHeightsPx = [...sheet.rowHeightsPx]
  rowHeightsPx.splice(atIndex, 0, undefined)

  const merges = sheet.merges.map((m) =>
    m.r0 >= atIndex ? { ...m, r0: m.r0 + 1, r1: m.r1 + 1 } : m,
  )

  return replaceSheet(doc, sheetIndex, recalculateSheet({ ...sheet, rows, formulas, rowHeightsPx, merges }))
}

/** Deletes the row at `atIndex`. A no-op if it would leave the sheet with zero rows. */
export function deleteRowAt(doc: SpreadsheetDocument, sheetIndex: number, atIndex: number): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet || sheet.rows.length <= 1 || atIndex < 0 || atIndex >= sheet.rows.length) return doc

  const rows = sheet.rows.filter((_, i) => i !== atIndex)
  const formulas = sheet.formulas.filter((_, i) => i !== atIndex)
  const rowHeightsPx = sheet.rowHeightsPx.filter((_, i) => i !== atIndex)

  const merges = sheet.merges
    .filter((m) => !(m.r0 === atIndex && m.r1 === atIndex))
    .map((m) => ({
      r0: m.r0 > atIndex ? m.r0 - 1 : m.r0,
      r1: m.r1 > atIndex ? m.r1 - 1 : m.r1,
      c0: m.c0,
      c1: m.c1,
    }))

  return replaceSheet(doc, sheetIndex, recalculateSheet({ ...sheet, rows, formulas, rowHeightsPx, merges }))
}

/** Inserts one empty column at `atIndex` (0-based; may equal `colCount` to append at the end). */
export function insertColumnAt(doc: SpreadsheetDocument, sheetIndex: number, atIndex: number): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet) return doc

  const rows = sheet.rows.map((row) => {
    const next = [...row]
    next.splice(atIndex, 0, '')
    return next
  })
  const formulas = sheet.formulas.map((row) => {
    const next = [...row]
    next.splice(atIndex, 0, undefined)
    return next
  })
  const colWidthsPx = [...sheet.colWidthsPx]
  colWidthsPx.splice(atIndex, 0, undefined)

  const merges = sheet.merges.map((m) =>
    m.c0 >= atIndex ? { ...m, c0: m.c0 + 1, c1: m.c1 + 1 } : m,
  )

  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({ ...sheet, rows, formulas, colWidthsPx, colCount: sheet.colCount + 1, merges }),
  )
}

/** Deletes the column at `atIndex`. A no-op if it would leave the sheet with zero columns. */
export function deleteColumnAt(doc: SpreadsheetDocument, sheetIndex: number, atIndex: number): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet || sheet.colCount <= 1 || atIndex < 0 || atIndex >= sheet.colCount) return doc

  const rows = sheet.rows.map((row) => row.filter((_, i) => i !== atIndex))
  const formulas = sheet.formulas.map((row) => row.filter((_, i) => i !== atIndex))
  const colWidthsPx = sheet.colWidthsPx.filter((_, i) => i !== atIndex)

  const merges = sheet.merges
    .filter((m) => !(m.c0 === atIndex && m.c1 === atIndex))
    .map((m) => ({
      c0: m.c0 > atIndex ? m.c0 - 1 : m.c0,
      c1: m.c1 > atIndex ? m.c1 - 1 : m.c1,
      r0: m.r0,
      r1: m.r1,
    }))

  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({ ...sheet, rows, formulas, colWidthsPx, colCount: sheet.colCount - 1, merges }),
  )
}

/** Pastes a rectangular block of TSV-parsed values starting at (row, col), growing the sheet if the block runs past its current bounds. */
export function pasteRange(
  doc: SpreadsheetDocument,
  sheetIndex: number,
  startRow: number,
  startCol: number,
  values: ReadonlyArray<ReadonlyArray<string>>,
): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet || values.length === 0) return doc

  const neededRows = startRow + values.length
  const neededCols = startCol + Math.max(...values.map((r) => r.length), 0)

  const rows = sheet.rows.map((row) => [...row])
  const formulas = sheet.formulas.map((row) => [...row])
  while (rows.length < neededRows) {
    rows.push(emptyRow(Math.max(sheet.colCount, neededCols)))
    formulas.push(emptyFormulaRow(Math.max(sheet.colCount, neededCols)))
  }
  const colCount = Math.max(sheet.colCount, neededCols)
  if (colCount > sheet.colCount) {
    for (let r = 0; r < rows.length; r++) {
      while (rows[r].length < colCount) rows[r].push('')
      while (formulas[r].length < colCount) formulas[r].push(undefined)
    }
  }

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const value = values[r][c]
      const targetRow = startRow + r
      const targetCol = startCol + c
      const isFormula = value.startsWith('=') && value.length > 1
      formulas[targetRow][targetCol] = isFormula ? value.slice(1) : undefined
      rows[targetRow][targetCol] = value
    }
  }

  const colWidthsPx = [...sheet.colWidthsPx]
  while (colWidthsPx.length < colCount) colWidthsPx.push(undefined)
  const rowHeightsPx = [...sheet.rowHeightsPx]
  while (rowHeightsPx.length < rows.length) rowHeightsPx.push(undefined)

  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({ ...sheet, rows, formulas, colCount, colWidthsPx, rowHeightsPx }),
  )
}

function defaultSheetName(existing: ReadonlyArray<EditableSheet>): string {
  let n = existing.length + 1
  const names = new Set(existing.map((s) => s.name))
  while (names.has(`Sheet${n}`)) n += 1
  return `Sheet${n}`
}

/** Appends a new, blank sheet (small usable default grid — see `NEW_SHEET_ROW_COUNT`/`NEW_SHEET_COL_COUNT`). */
export function addSheet(doc: SpreadsheetDocument, name?: string): SpreadsheetDocument {
  const sheetName = name?.trim() || defaultSheetName(doc.sheets)
  const newSheet: EditableSheet = {
    name: sheetName,
    hidden: false,
    rows: Array.from({ length: NEW_SHEET_ROW_COUNT }, () => emptyRow(NEW_SHEET_COL_COUNT)),
    formulas: Array.from({ length: NEW_SHEET_ROW_COUNT }, () => emptyFormulaRow(NEW_SHEET_COL_COUNT)),
    colCount: NEW_SHEET_COL_COUNT,
    merges: [],
    colWidthsPx: [],
    rowHeightsPx: [],
  }
  return { sheets: [...doc.sheets, newSheet] }
}

/** Renames a sheet. A no-op for a blank name or a name already used by another sheet. */
export function renameSheet(doc: SpreadsheetDocument, sheetIndex: number, name: string): SpreadsheetDocument {
  const trimmed = name.trim()
  const sheet = doc.sheets[sheetIndex]
  if (!sheet || !trimmed) return doc
  if (doc.sheets.some((s, i) => i !== sheetIndex && s.name === trimmed)) return doc

  return replaceSheet(doc, sheetIndex, { ...sheet, name: trimmed })
}

/** Deletes a sheet. A no-op if it's the only sheet left (a workbook must keep at least one). */
export function deleteSheet(doc: SpreadsheetDocument, sheetIndex: number): SpreadsheetDocument {
  if (doc.sheets.length <= 1 || !doc.sheets[sheetIndex]) return doc
  return { sheets: doc.sheets.filter((_, i) => i !== sheetIndex) }
}
