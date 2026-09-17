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
import {
  tablesAfterColumnDelete,
  tablesAfterColumnInsert,
  tablesAfterRowDelete,
  tablesAfterRowInsert,
  type SheetTable,
} from './spreadsheetTables'

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
  /** Excel tables (USR-17), kept in step with row/column inserts and deletes. */
  readonly tables?: ReadonlyArray<SheetTable>
  /**
   * Save-through-the-original bookkeeping (USR-17, see `xlsxPassthrough.ts`):
   * the worksheet part this sheet was read from, where each row/column came
   * from in that part (`null` = added here), and which cells the user has
   * actually changed. Together they let a save keep every untouched cell —
   * with its number format, style, type and cached formula value — exactly as
   * the original file had it, instead of rewriting the workbook from display
   * text. Absent for CSV/TSV and for sheets Atlas created itself.
   */
  readonly sourcePath?: string
  readonly rowSources?: ReadonlyArray<number | null>
  readonly colSources?: ReadonlyArray<number | null>
  /** `"row:col"` (current indexes) of every cell edited since load. */
  readonly editedCells?: ReadonlySet<string>
}

/** Key for `editedCells`. */
export function cellKey(row: number, col: number): string {
  return `${row}:${col}`
}

/** Remaps `editedCells` after rows or columns move, dropping the ones that were deleted. */
function remapEditedCells(
  edited: ReadonlySet<string> | undefined,
  remap: (row: number, col: number) => readonly [number, number] | null,
): ReadonlySet<string> | undefined {
  if (!edited) return undefined
  const next = new Set<string>()
  for (const key of edited) {
    const [row, col] = key.split(':').map(Number)
    const moved = remap(row, col)
    if (moved) next.add(cellKey(moved[0], moved[1]))
  }
  return next
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
        ...(sheet.tables ? { tables: sheet.tables } : {}),
        ...(sheet.sourcePath !== undefined
          ? {
              sourcePath: sheet.sourcePath,
              rowSources: sheet.grid.rows.map((_, i) => i),
              colSources: Array.from({ length: sheet.grid.colCount }, (_, i) => i),
              editedCells: new Set<string>(),
            }
          : {}),
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

/** Spread into an updated sheet: its tables after a row/column insert or delete, or nothing when it has none. */
function shiftedTables(
  sheet: EditableSheet,
  shift: (tables: ReadonlyArray<SheetTable>, at: number) => ReadonlyArray<SheetTable>,
  atIndex: number,
): { readonly tables?: ReadonlyArray<SheetTable> } {
  return sheet.tables ? { tables: shift(sheet.tables, atIndex) } : {}
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
 * formula cell is actually found; only then is the copy-on-write triggered —
 * and even then, only the specific ROW a formula cell was found in is ever
 * cloned (`rowCopied`, below). Every other row keeps its exact original
 * array reference. This matters because `setCellValue` calls this after
 * EVERY keystroke-commit: a sheet with even a single formula cell anywhere
 * would otherwise pay an O(rows × cols) copy on every unrelated edit
 * elsewhere in a 100k-row sheet, not just the O(rows) the edit itself needs.
 */
export function recalculateSheet(sheet: EditableSheet): EditableSheet {
  let rows: string[][] | null = null
  const lookup: CellLookup = (row, col) => (rows ?? sheet.rows)[row]?.[col] ?? ''

  for (let r = 0; r < sheet.formulas.length; r++) {
    const formulaRow = sheet.formulas[r]
    let rowCopied = false
    for (let c = 0; c < formulaRow.length; c++) {
      const formula = formulaRow[c]
      if (formula === undefined) continue
      // Outer array only, on first formula cell found anywhere in the sheet
      // — every row still points at `sheet.rows`' own (readonly-typed, but
      // never actually written to until `rowCopied` below) array until it's
      // this row's turn to actually be written to, just below.
      if (rows === null) rows = sheet.rows.map((row) => row as string[])
      if (!rowCopied) {
        rows[r] = [...rows[r]]
        rowCopied = true
      }

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
  // Only the touched row needs a fresh array — every other row keeps its
  // existing reference (see `recalculateSheet`'s identical note just above:
  // this is called on every keystroke-commit, so cloning every row of a
  // 100k-row sheet to change one cell does not scale). The untouched
  // branch's cast is safe: this function only ever writes into index `row`.
  const rows: string[][] = sheet.rows.map((r, i) => (i === row ? [...r] : (r as string[])))
  const formulas: (string | undefined)[][] = sheet.formulas.map((r, i) =>
    i === row ? [...r] : (r as (string | undefined)[]),
  )

  formulas[row][col] = isFormula ? rawInput.slice(1) : undefined
  // For a formula cell this is immediately overwritten by `recalculateSheet`
  // below with the evaluated (or literal-formula-fallback) text; for a plain
  // value cell this IS the final stored display text.
  rows[row][col] = rawInput

  const editedCells = sheet.editedCells ? new Set(sheet.editedCells).add(cellKey(row, col)) : undefined
  const updated = recalculateSheet({ ...sheet, rows, formulas, ...(editedCells ? { editedCells } : {}) })
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

  // A merge entirely at/after the insertion point shifts down whole; one
  // that STRADDLES it (starts before, ends at/after) instead grows by one
  // row, since the new row lands inside it; one entirely before it is
  // untouched.
  const merges = sheet.merges.map((m) => {
    if (atIndex <= m.r0) return { ...m, r0: m.r0 + 1, r1: m.r1 + 1 }
    if (atIndex <= m.r1) return { ...m, r1: m.r1 + 1 }
    return m
  })

  const tables = shiftedTables(sheet, tablesAfterRowInsert, atIndex)
  const sources = sheet.rowSources ? [...sheet.rowSources] : null
  if (sources) sources.splice(atIndex, 0, null)
  const editedCells = remapEditedCells(sheet.editedCells, (r, c) => [r >= atIndex ? r + 1 : r, c])
  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({
      ...sheet,
      rows,
      formulas,
      rowHeightsPx,
      merges,
      ...tables,
      ...(sources ? { rowSources: sources } : {}),
      ...(editedCells ? { editedCells } : {}),
    }),
  )
}

/** Deletes the row at `atIndex`. A no-op if it would leave the sheet with zero rows. */
export function deleteRowAt(doc: SpreadsheetDocument, sheetIndex: number, atIndex: number): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet || sheet.rows.length <= 1 || atIndex < 0 || atIndex >= sheet.rows.length) return doc

  const rows = sheet.rows.filter((_, i) => i !== atIndex)
  const formulas = sheet.formulas.filter((_, i) => i !== atIndex)
  const rowHeightsPx = sheet.rowHeightsPx.filter((_, i) => i !== atIndex)

  // r0 shifts only when the deleted row was strictly ABOVE it (`> atIndex`):
  // a merge whose top row IS the deleted row keeps r0's numeric value
  // unchanged, since the row that slides up to fill that index was already
  // part of the merge. r1 must shift on `>= atIndex` (not `> atIndex`): a
  // merge whose BOTTOM row is exactly the deleted row needs to shrink by one
  // too, or it silently grows to swallow the next surviving row (which was
  // never part of it) once that row slides up into the vacated index — the
  // asymmetry with r0 is intentional, not a copy-paste of the same rule.
  const merges = sheet.merges
    .filter((m) => !(m.r0 === atIndex && m.r1 === atIndex))
    .map((m) => ({
      r0: m.r0 > atIndex ? m.r0 - 1 : m.r0,
      r1: m.r1 >= atIndex ? m.r1 - 1 : m.r1,
      c0: m.c0,
      c1: m.c1,
    }))

  const tables = shiftedTables(sheet, tablesAfterRowDelete, atIndex)
  const rowSources = sheet.rowSources?.filter((_, i) => i !== atIndex)
  const editedCells = remapEditedCells(sheet.editedCells, (r, c) =>
    r === atIndex ? null : [r > atIndex ? r - 1 : r, c],
  )
  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({
      ...sheet,
      rows,
      formulas,
      rowHeightsPx,
      merges,
      ...tables,
      ...(rowSources ? { rowSources } : {}),
      ...(editedCells ? { editedCells } : {}),
    }),
  )
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

  // See `insertRowAt`'s identical comment: a merge straddling the insertion
  // point grows by one column instead of shifting whole.
  const merges = sheet.merges.map((m) => {
    if (atIndex <= m.c0) return { ...m, c0: m.c0 + 1, c1: m.c1 + 1 }
    if (atIndex <= m.c1) return { ...m, c1: m.c1 + 1 }
    return m
  })

  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({
      ...sheet,
      rows,
      formulas,
      colWidthsPx,
      colCount: sheet.colCount + 1,
      merges,
      ...shiftedTables(sheet, tablesAfterColumnInsert, atIndex),
      ...(sheet.colSources
        ? { colSources: [...sheet.colSources.slice(0, atIndex), null, ...sheet.colSources.slice(atIndex)] }
        : {}),
      ...(() => {
        const editedCells = remapEditedCells(sheet.editedCells, (r, c) => [r, c >= atIndex ? c + 1 : c])
        return editedCells ? { editedCells } : {}
      })(),
    }),
  )
}

/** Deletes the column at `atIndex`. A no-op if it would leave the sheet with zero columns. */
export function deleteColumnAt(doc: SpreadsheetDocument, sheetIndex: number, atIndex: number): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet || sheet.colCount <= 1 || atIndex < 0 || atIndex >= sheet.colCount) return doc

  const rows = sheet.rows.map((row) => row.filter((_, i) => i !== atIndex))
  const formulas = sheet.formulas.map((row) => row.filter((_, i) => i !== atIndex))
  const colWidthsPx = sheet.colWidthsPx.filter((_, i) => i !== atIndex)

  // See `deleteRowAt`'s identical comment: c1 shifts on `>= atIndex` (not
  // `> atIndex`) so a merge whose RIGHT edge is exactly the deleted column
  // shrinks instead of silently growing into the next surviving column.
  const merges = sheet.merges
    .filter((m) => !(m.c0 === atIndex && m.c1 === atIndex))
    .map((m) => ({
      c0: m.c0 > atIndex ? m.c0 - 1 : m.c0,
      c1: m.c1 >= atIndex ? m.c1 - 1 : m.c1,
      r0: m.r0,
      r1: m.r1,
    }))

  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({
      ...sheet,
      rows,
      formulas,
      colWidthsPx,
      colCount: sheet.colCount - 1,
      merges,
      ...shiftedTables(sheet, tablesAfterColumnDelete, atIndex),
      ...(sheet.colSources ? { colSources: sheet.colSources.filter((_, i) => i !== atIndex) } : {}),
      ...(() => {
        const editedCells = remapEditedCells(sheet.editedCells, (r, c) =>
          c === atIndex ? null : [r, c > atIndex ? c - 1 : c],
        )
        return editedCells ? { editedCells } : {}
      })(),
    }),
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
  const colCount = Math.max(sheet.colCount, neededCols)
  // Widening (a paste running past the sheet's right edge) genuinely touches
  // every existing row's array, since each one needs padding out to the new
  // width — but the far more common case (pasting within the sheet's
  // current bounds) doesn't, so only the rows the paste actually WRITES to
  // get cloned then. Same reasoning as `setCellValue`'s identical comment:
  // this must not cost an O(rows) clone of a 100k-row sheet to paste one row.
  const widening = colCount > sheet.colCount
  const isPastedRow = (i: number): boolean => i >= startRow && i < startRow + values.length

  // The untouched-row casts are safe: the write loop below only ever
  // indexes into `[startRow, startRow + values.length)`, exactly the rows
  // this map already clones via `isPastedRow`.
  const rows: string[][] = sheet.rows.map((row, i) => {
    if (!widening && !isPastedRow(i)) return row as string[]
    const next = [...row]
    while (next.length < colCount) next.push('')
    return next
  })
  const formulas: (string | undefined)[][] = sheet.formulas.map((row, i) => {
    if (!widening && !isPastedRow(i)) return row as (string | undefined)[]
    const next = [...row]
    while (next.length < colCount) next.push(undefined)
    return next
  })
  while (rows.length < neededRows) {
    rows.push(emptyRow(colCount))
    formulas.push(emptyFormulaRow(colCount))
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

  // Rows/columns the paste added have no counterpart in the original file.
  const rowSources = sheet.rowSources ? [...sheet.rowSources] : null
  while (rowSources && rowSources.length < rows.length) rowSources.push(null)
  const colSources = sheet.colSources ? [...sheet.colSources] : null
  while (colSources && colSources.length < colCount) colSources.push(null)
  const editedCells = sheet.editedCells ? new Set(sheet.editedCells) : null
  if (editedCells) {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) editedCells.add(cellKey(startRow + r, startCol + c))
    }
  }

  return replaceSheet(
    doc,
    sheetIndex,
    recalculateSheet({
      ...sheet,
      rows,
      formulas,
      colCount,
      colWidthsPx,
      rowHeightsPx,
      ...(rowSources ? { rowSources } : {}),
      ...(colSources ? { colSources } : {}),
      ...(editedCells ? { editedCells } : {}),
    }),
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
