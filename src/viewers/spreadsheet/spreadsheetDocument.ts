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
 * Formula recalculation (SHEET-5): any edit that could change a formula's
 * result (a value edit, a formula edit, a row/column insert or delete, a
 * paste) triggers exactly one recalculation pass over every formula cell in
 * that sheet — but that pass runs in real DEPENDENCY order, not row-major
 * layout order. `recalculateSheet` scans each formula's text for the cells
 * it references (a single ref or a range like `A1:A3`) and topologically
 * sorts the sheet's formula cells so that, by the time a cell is evaluated,
 * every OTHER formula cell it depends on has already been recomputed in
 * THIS SAME pass — however far apart the two cells are in row/column order.
 * The sort is an iterative (not recursive) traversal specifically so a long
 * dependency chain can't blow the call stack. A circular reference (a cell
 * that, directly or through a chain, depends on itself) is caught by that
 * same traversal — an edge back to a node still mid-visit — and every cell
 * in the cycle resolves to `#REF!` (this codebase's existing vocabulary for
 * "this reference doesn't resolve", matching `xlsxPassthrough.ts`'s and
 * `formulaRefs.ts`'s own use of it) instead of hanging or recursing forever.
 *
 * Inserting/deleting a row or column still does NOT shift other cells'
 * formula references (e.g. deleting row 2 leaves a `=A3` elsewhere still
 * reading literal row 3) — implementing Excel's reference-shifting semantics
 * is a substantial, separate undertaking, out of scope here.
 */
import { compareForSort } from './cellOrder'
import { evaluateFormula } from './spreadsheetFormula'
import { applyNumberFormat } from './formatNumber'
import type { CellLookup } from './spreadsheetFormula'
import { parseCellRange } from './cellRef'
import type { FrozenPanes, MergeRange, ParsedSheet } from '../shared/spreadsheetGrid'
import {
  DEFAULT_CELL_FORMAT,
  formatAt,
  NO_BORDER,
  type CellBorder,
  type ResolvedCellFormat,
  type SheetCellStyles,
} from './xlsxCellStyles'
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
  /**
   * SHEETFMT-2 — formatting the user applied in Atlas, by CURRENT `cellKey`
   * (so, unlike `cellStyles` below, it is remapped as rows and columns move,
   * exactly like `editedCells`).
   *
   * A patch, not a whole format: "make this bold" must leave the fill the file
   * gave the cell alone, and the save path needs to know which properties were
   * actually chosen so it can clone the rest of the cell's original `<xf>`.
   */
  readonly formatOverrides?: ReadonlyMap<string, CellFormatPatch>
  /**
   * SHEETFMT-1 — the formatting this sheet's cells carry in the file, indexed
   * by the cell's ORIGINAL row/column (the same coordinates `rowSources`/
   * `colSources` map back to), so a row insert or delete moves formats with
   * the rows they belong to instead of shifting them all down. Resolve a
   * current cell's format through `formatForCell` rather than indexing this
   * directly. Absent for CSV/TSV and for any format with no styles to read.
   */
  readonly cellStyles?: SheetCellStyles
}

/**
 * The format for a cell at its CURRENT row/column.
 *
 * Goes through `rowSources`/`colSources` because the two coordinate systems
 * diverge the moment a row or column is inserted or deleted: `cellStyles` is
 * indexed by where a cell was in the FILE, while `row`/`col` here are where it
 * is now. Inserting a row above a styled block and then reading formats by the
 * current index would slide every format in the sheet one row out of place —
 * a bug that looks like "the formatting is subtly wrong" rather than like a
 * mapping error.
 *
 * A row or column Atlas added itself has no original index (`null`), and
 * correctly has no formatting.
 */
export function formatForCell(sheet: EditableSheet, row: number, col: number): ResolvedCellFormat {
  const patch = sheet.formatOverrides?.get(cellKey(row, col))
  if (!sheet.cellStyles) return mergeFormat(DEFAULT_CELL_FORMAT, patch)
  const sourceRow = sheet.rowSources ? sheet.rowSources[row] : row
  const sourceCol = sheet.colSources ? sheet.colSources[col] : col
  if (sourceRow === null || sourceRow === undefined || sourceCol === null || sourceCol === undefined) {
    // A row or column Atlas added itself has no formatting from the file, but
    // can perfectly well have formatting the user applied to it since.
    return mergeFormat(DEFAULT_CELL_FORMAT, patch)
  }
  return mergeFormat(formatAt(sheet.cellStyles, sourceRow, sourceCol), patch)
}

/**
 * Applies a formatting patch to every cell in a rectangular range.
 *
 * Patches COMPOSE rather than replace: pressing bold and then filling a cell
 * leaves it bold and filled. Clearing a property is `null` in the patch, which
 * survives into the stored override so a save can write "explicitly not bold"
 * over a cell the file made bold.
 *
 * Returns the same document reference when the range is empty or out of bounds,
 * matching every other operation in this module so `mutate` can detect a no-op
 * and keep it off the undo stack.
 */
export function setRangeFormat(
  doc: SpreadsheetDocument,
  sheetIndex: number,
  range: { readonly row0: number; readonly col0: number; readonly row1: number; readonly col1: number },
  patch: CellFormatPatch,
): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet) return doc
  const row0 = Math.max(0, Math.min(range.row0, range.row1))
  const row1 = Math.min(sheet.rows.length - 1, Math.max(range.row0, range.row1))
  const rows: number[] = []
  for (let r = row0; r <= row1; r += 1) rows.push(r)
  return setRowsFormat(doc, sheetIndex, rows, range.col0, range.col1, patch)
}

/**
 * Applies a formatting patch to an explicit LIST of rows, across a column span.
 *
 * The list, rather than a row range, is what the grid actually needs: a row
 * search hides rows, so a contiguous block of selected grid rows can map to a
 * non-contiguous set of sheet rows. Formatting the rows in between — the ones
 * filtered out of view, which the user cannot see and did not select — would be
 * a silent edit to hidden data.
 */
export function setRowsFormat(
  doc: SpreadsheetDocument,
  sheetIndex: number,
  rows: ReadonlyArray<number>,
  colA: number,
  colB: number,
  patch: CellFormatPatch,
): SpreadsheetDocument {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet) return doc
  if (Object.keys(patch).length === 0) return doc

  const col0 = Math.max(0, Math.min(colA, colB))
  const col1 = Math.min(sheet.colCount - 1, Math.max(colA, colB))
  if (col1 < col0) return doc

  const inRange = rows.filter(r => r >= 0 && r < sheet.rows.length)
  if (inRange.length === 0) return doc

  const next = new Map(sheet.formatOverrides ?? [])
  for (const r of inRange) {
    for (let c = col0; c <= col1; c += 1) {
      const key = cellKey(r, c)
      next.set(key, { ...next.get(key), ...patch })
    }
  }
  return replaceSheetAndRecalculate(doc, sheetIndex, { ...sheet, formatOverrides: next })
}

/** Key for `editedCells`. */
export function cellKey(row: number, col: number): string {
  return `${row}:${col}`
}

/**
 * SHEETFMT-2 — a formatting change the user asked for.
 *
 * Every property is optional and "absent" means "leave whatever the cell
 * already had", which is what makes a bold button composable with a fill
 * button. `null` is distinct from absent and means "clear this back to the
 * default" — the only way to express removing a fill the file put there.
 */
export type CellFormatPatch = {
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strike?: boolean
  readonly fontSize?: number | null
  readonly fontName?: string | null
  readonly color?: string | null
  readonly fill?: string | null
  readonly align?: 'left' | 'center' | 'right' | null
  readonly numberFormat?: string | null
  /**
   * SHEETFMT-3 — which edges to put a border on.
   *
   * A whole `CellBorder` rather than per-edge flags, so "box this cell" and
   * "underline this row" are one patch each. `null` clears every edge, which is
   * the only way back from a border the file applied.
   */
  readonly border?: CellBorder | null
}

/** Applies a patch over a resolved format, for rendering. `null` in the patch clears back to the default. */
export function mergeFormat(base: ResolvedCellFormat, patch: CellFormatPatch | undefined): ResolvedCellFormat {
  if (!patch) return base
  const pick = <T,>(patched: T | null | undefined, current: T | undefined): T | undefined =>
    patched === undefined ? current : patched === null ? undefined : patched
  return {
    ...base,
    bold: patch.bold ?? base.bold,
    italic: patch.italic ?? base.italic,
    underline: patch.underline ?? base.underline,
    strike: patch.strike ?? base.strike,
    fontSize: pick(patch.fontSize, base.fontSize),
    fontName: pick(patch.fontName, base.fontName),
    color: pick(patch.color, base.color),
    fill: pick(patch.fill, base.fill),
    align: pick(patch.align, base.align),
    numberFormat: pick(patch.numberFormat, base.numberFormat),
    border: patch.border === undefined ? base.border : (patch.border ?? NO_BORDER),
  }
}

/** Remaps `formatOverrides` after rows or columns move, dropping the ones that were deleted. Mirrors `remapEditedCells`. */
function remapFormatOverrides(
  overrides: ReadonlyMap<string, CellFormatPatch> | undefined,
  remap: (row: number, col: number) => readonly [number, number] | null,
): ReadonlyMap<string, CellFormatPatch> | undefined {
  if (!overrides) return undefined
  const next = new Map<string, CellFormatPatch>()
  for (const [key, patch] of overrides) {
    const [row, col] = key.split(':').map(Number)
    const moved = remap(row, col)
    if (moved) next.set(cellKey(moved[0], moved[1]), patch)
  }
  return next
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
  // SHEETFN-5 — built first, then recalculated ONCE across the whole document,
  // so a formula referencing another sheet resolves on load rather than only
  // after the first edit. A per-sheet pass could not: the sheet it points at
  // may not have been built yet.
  return recalculateDocument({
    sheets: parsedSheets.map((sheet) => ({
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
        ...(sheet.cellStyles ? { cellStyles: sheet.cellStyles } : {}),
        ...(sheet.sourcePath !== undefined
          ? {
              sourcePath: sheet.sourcePath,
              rowSources: sheet.grid.rows.map((_, i) => i),
              colSources: Array.from({ length: sheet.grid.colCount }, (_, i) => i),
              editedCells: new Set<string>(),
            }
          : {}),
      })),
  })
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
 * Puts an edited sheet back and recalculates the WHOLE document.
 *
 * SHEETFN-5 — every operation used to recalculate only the sheet it touched,
 * which is wrong the moment a formula crosses sheets: editing `Sheet2!B1` has
 * to update `Sheet1!A1 = Sheet2!B1`, and a per-sheet pass never looks at
 * Sheet1. It also covers formatting changes, because a cell's number format
 * decides how a formula result is rendered (SHEETFN-3).
 *
 * `recalculateDocument` returns the same document object when nothing changed,
 * so this stays a no-op for an edit that computes to what was already there —
 * which the editor's undo history depends on.
 */
function replaceSheetAndRecalculate(
  doc: SpreadsheetDocument,
  sheetIndex: number,
  sheet: EditableSheet,
): SpreadsheetDocument {
  return recalculateDocument(replaceSheet(doc, sheetIndex, sheet))
}

/** '#REF!' as a circular-reference result — see the module header: the same code this codebase already uses elsewhere for "this reference doesn't resolve" (`xlsxPassthrough.ts`, `formulaRefs.ts`), reused here rather than inventing a distinct "circular" error vocabulary. */
const CIRCULAR_REFERENCE_ERROR = '#REF!'

/**
 * A cell reference/range shape, matching `spreadsheetFormula.ts`'s own
 * tokenizer: an optional `SheetName!`/`'Sheet Name'!` qualifier, then
 * `$?letters$?digits`, optionally `:$?letters$?digits` for a range. Built fresh
 * per call (not a module-level `RegExp`) so concurrent/re-entrant use can't
 * stomp on a shared `lastIndex`.
 *
 * SHEETFN-5 — the qualifier is captured, because a dependency on another
 * sheet's cell has to be ordered against THAT sheet's formulas, not this
 * sheet's cell of the same address.
 */
function referenceOrRangePattern(): RegExp {
  return /(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))?!?(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)/g
}

/**
 * Every `cellKey(row, col)` this formula's text could read during
 * evaluation — one per referenced cell, or every cell inside a range like
 * `A1:A3`. This is a lightweight text scan, not a full parse (see
 * `spreadsheetFormula.ts` for that); a SUPERSET of the true dependencies is
 * safe here, since an extra dependency only adds an ordering constraint, it
 * can never make a value wrong — a MISSED one is exactly the stale-value bug
 * this function exists to fix, so scanning wider than strictly necessary is
 * the safe direction to err in.
 *
 * Bounded to the sheet's own extent (`rowCount`/`colCount`): a formula like
 * `=SUM(A1:A1048576)` typed against a 20-row sheet doesn't walk a million
 * phantom rows that can't contain a formula cell anyway. This costs no more
 * than `evaluateFormula` itself already pays to compute that same formula's
 * result — not a new order of magnitude.
 */
function formulaDependencyKeys(
  formula: string,
  rowCount: number,
  colCount: number,
  /** SHEETFN-5 — turns a reference's sheet qualifier (or its absence) into the key prefix for that sheet. */
  keyFor: (sheet: string | undefined, row: number, col: number) => string | null = (_, row, col) =>
    cellKey(row, col),
): string[] {
  const deps: string[] = []
  const seen = new Set<string>()
  const pattern = referenceOrRangePattern()
  let match: RegExpExecArray | null
  while ((match = pattern.exec(formula)) !== null) {
    const sheet = match[1] ?? match[2]
    const range = parseCellRange(match[3])
    if (!range) continue
    // A qualified reference may point at a sheet with a different extent, so
    // the bound only applies to a reference on this sheet. Scanning wider than
    // necessary is the safe direction (see this function's header).
    const r1 = sheet === undefined ? Math.min(range.end.row, rowCount - 1) : range.end.row
    const c1 = sheet === undefined ? Math.min(range.end.col, colCount - 1) : range.end.col
    for (let r = range.start.row; r <= r1; r++) {
      for (let c = range.start.col; c <= c1; c++) {
        const key = keyFor(sheet, r, c)
        if (key !== null && !seen.has(key)) {
          seen.add(key)
          deps.push(key)
        }
      }
    }
  }
  return deps
}

type DfsFrame = { readonly key: string; readonly deps: ReadonlyArray<string>; depIndex: number }

/** A back-edge from the frame currently being expanded to `target`, still mid-visit on `stack`: every frame from `target` up to the top of the stack is part of one cycle. */
function markCycle(stack: ReadonlyArray<DfsFrame>, target: string, cyclic: Set<string>): void {
  let start = stack.findIndex((frame) => frame.key === target)
  if (start === -1) start = stack.length - 1 // defensive only — an IN-PROGRESS node is always still on this stack
  for (let i = start; i < stack.length; i++) cyclic.add(stack[i].key)
}

/**
 * Topologically sorts `formulaCells` (keyed by `cellKey`) so every cell
 * comes after every OTHER formula cell it depends on — an iterative
 * (explicit-stack) post-order DFS, not a recursive one, so a long
 * dependency chain can't blow the call stack. `cyclic` collects every key
 * that sits on a circular reference (a back-edge to a node still mid-visit);
 * those cells are still included in `order` (still need SOME resolution),
 * just flagged so the caller skips evaluating them and writes the circular-
 * reference error instead. A cell's dependency on its OWN key is never part
 * of this — see `frameFor`'s own comment for why a bare self-mention isn't
 * treated as a one-cell cycle here.
 */
function topologicalFormulaOrder<T>(
  formulaCells: ReadonlyMap<string, T>,
  /**
   * SHEETFN-5 — the dependencies of one entry, as keys into `formulaCells`.
   * Passed in rather than computed here, so the same sort serves a single
   * sheet's `r:c` keys and a whole document's `sheet|r:c` ones.
   */
  dependenciesOf: (key: string, entry: T) => ReadonlyArray<string>,
): { readonly order: ReadonlyArray<T>; readonly cyclic: ReadonlySet<string> } {
  const IN_PROGRESS = 1
  const DONE = 2
  const state = new Map<string, typeof IN_PROGRESS | typeof DONE>()
  const cyclic = new Set<string>()
  const order: T[] = []

  const frameFor = (key: string): DfsFrame => {
    const entry = formulaCells.get(key)
    // A dependency on the cell's OWN key is dropped, not treated as a
    // (trivial, one-cell) cycle: `lookup` is a plain, non-recursive array
    // read (see the header above `recalculateSheet`), so a formula that
    // merely MENTIONS its own cell — e.g. as one argument among several to
    // an unsupported function like `VLOOKUP(A1,...)` stored AT A1, which
    // `createDocument`'s own fixtures exercise — never risks a hang; it just
    // reads whatever text is already sitting there, exactly as it always
    // has. Forcing `#REF!` on every such mention would also misfire on that
    // case: the formula fails for an unrelated reason (the unsupported
    // function), and the existing, tested contract is to keep the cell's
    // cached display text in that situation, not to overwrite it with a
    // circular-reference error it was never actually driven by. A genuinely
    // circular SELF-reference that a supported formula actually depends on
    // (`=A1+1` stored at A1) is a narrower, pre-existing edge case this
    // still doesn't flag — see this function's own header — left as-is
    // rather than risk that same false positive.
    const deps = entry === undefined
      ? []
      : dependenciesOf(key, entry).filter((k) => k !== key && formulaCells.has(k))
    state.set(key, IN_PROGRESS)
    return { key, deps, depIndex: 0 }
  }

  for (const startKey of formulaCells.keys()) {
    if (state.get(startKey) === DONE) continue

    const stack: DfsFrame[] = [frameFor(startKey)]
    while (stack.length > 0) {
      const frame = stack[stack.length - 1]
      if (frame.depIndex < frame.deps.length) {
        const depKey = frame.deps[frame.depIndex]
        frame.depIndex += 1
        const depState = state.get(depKey)
        if (depState === DONE) continue
        if (depState === IN_PROGRESS) {
          markCycle(stack, depKey, cyclic)
          continue
        }
        stack.push(frameFor(depKey))
      } else {
        stack.pop()
        state.set(frame.key, DONE)
        // frame.key always has an entry: every key ever pushed came either
        // from `formulaCells.keys()` directly or from a dependency already
        // filtered to `formulaCells.has(k)` in `frameFor`.
        order.push(formulaCells.get(frame.key)!)
      }
    }
  }

  return { order, cyclic }
}

/**
 * Recomputes every formula cell's display text in true DEPENDENCY order
 * (see module header) rather than row-major layout order — a topological
 * sort over the sheet's formula cells, with circular references caught and
 * resolved to `#REF!` instead of hanging.
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
 * pay for an O(rows) copy of every row it never needed. Collecting
 * `formulaCells` below is still an O(rows × cols) SCAN either way (it always
 * was, even before this rewrite — the original loop walked every cell to
 * check `formula === undefined` too); what must stay lazy is the ARRAY COPY,
 * not the scan. So this returns the original `sheet` unchanged the moment
 * that scan finds zero formula cells, and otherwise only clones the outer
 * `rows` array once (lazily, `ensureRowWritable` below) and each individual
 * row at most once, the first time a value actually changes in it — never a
 * blanket O(rows × cols) copy. This matters because `setCellValue` calls
 * this after EVERY keystroke-commit: a sheet with even a single formula cell
 * anywhere would otherwise pay a full-sheet copy on every unrelated edit
 * elsewhere in a 100k-row sheet, not just the O(rows) the edit itself needs.
 */
export function recalculateSheet(sheet: EditableSheet): EditableSheet {
  // A single sheet has no way to resolve another one's cells. A qualified
  // reference to its OWN name still works (Excel allows `=Sheet1!A1` on
  // Sheet1); anything else is `#REF!`, which is honest — the alternative,
  // silently reading this sheet's cell of the same address, would give a
  // confidently wrong number.
  const recalculated = recalculateDocument({ sheets: [sheet] })
  return recalculated.sheets[0]
}

/** Everything one formula cell needs, across a whole document. */
type DocumentFormulaEntry = {
  readonly sheetIndex: number
  readonly row: number
  readonly col: number
  readonly formula: string
}

/** The key a formula cell is ordered by, document-wide. */
function documentCellKey(sheetIndex: number, row: number, col: number): string {
  return `${sheetIndex}|${row}:${col}`
}

/**
 * SHEETFN-5 — recalculates every formula in the document, across sheets.
 *
 * `recalculateSheet` could only ever see one sheet, so `=Sheet2!A1` was
 * unresolvable and displayed as its own formula text. Resolving it needs three
 * things at once, which is why this is document-level rather than a lookup
 * passed into the per-sheet version:
 *
 *   - **a sheet-aware lookup**, so a reference reads the right sheet;
 *   - **one dependency graph spanning all sheets**, so `Sheet1!A1 = Sheet2!B1`
 *     is evaluated after `Sheet2!B1` even though they live in different sheets
 *     — ordering per sheet would leave one of them reading a stale value, and
 *     which one would depend on sheet order;
 *   - **cycle detection across sheets**, so `Sheet1!A1 = Sheet2!A1` and
 *     `Sheet2!A1 = Sheet1!A1` report a circular reference instead of looping or
 *     silently settling on whatever was there before.
 *
 * Sheet names resolve case-insensitively, as Excel's do.
 *
 * Returns the SAME document object when nothing changed. The editor's undo
 * history detects a no-op edit by reference (`if (next !== current)`), so
 * always returning a fresh object would put a duplicate entry on the undo stack
 * for every blocked action.
 */
export function recalculateDocument(doc: SpreadsheetDocument): SpreadsheetDocument {
  const indexByName = new Map<string, number>()
  doc.sheets.forEach((sheet, index) => {
    // First wins, matching how a workbook with duplicate names would be read.
    const key = sheet.name.toLowerCase()
    if (!indexByName.has(key)) indexByName.set(key, index)
  })

  const formulaCells = new Map<string, DocumentFormulaEntry>()
  doc.sheets.forEach((sheet, sheetIndex) => {
    for (let r = 0; r < sheet.formulas.length; r++) {
      const formulaRow = sheet.formulas[r]
      for (let c = 0; c < formulaRow.length; c++) {
        const formula = formulaRow[c]
        if (formula === undefined) continue
        formulaCells.set(documentCellKey(sheetIndex, r, c), { sheetIndex, row: r, col: c, formula })
      }
    }
  })

  if (formulaCells.size === 0) return doc

  /** Resolves a reference's sheet qualifier to a sheet index, or `null` for one that does not exist. */
  const sheetIndexFor = (name: string | undefined, fallback: number): number | null => {
    if (name === undefined) return fallback
    return indexByName.get(name.toLowerCase()) ?? null
  }

  const { order, cyclic } = topologicalFormulaOrder(formulaCells, (_key, entry) => {
    const sheet = doc.sheets[entry.sheetIndex]
    return formulaDependencyKeys(
      entry.formula,
      sheet.rows.length,
      sheet.colCount,
      (refSheet, row, col) => {
        const index = sheetIndexFor(refSheet, entry.sheetIndex)
        return index === null ? null : documentCellKey(index, row, col)
      },
    )
  })

  // Rows are copied lazily, per sheet, so a document whose formulas all
  // re-compute to what they already showed returns unchanged.
  const working = new Map<number, string[][]>()
  const clonedRows = new Map<number, Set<number>>()

  const rowsOf = (sheetIndex: number): ReadonlyArray<ReadonlyArray<string>> =>
    working.get(sheetIndex) ?? doc.sheets[sheetIndex].rows

  const ensureRowWritable = (sheetIndex: number, r: number): string[] => {
    let rows = working.get(sheetIndex)
    if (rows === undefined) {
      rows = doc.sheets[sheetIndex].rows.map((row) => row as string[])
      working.set(sheetIndex, rows)
      clonedRows.set(sheetIndex, new Set())
    }
    const cloned = clonedRows.get(sheetIndex)!
    if (!cloned.has(r)) {
      rows[r] = [...rows[r]]
      cloned.add(r)
    }
    return rows[r]
  }

  /**
   * Writes a computed value, but only when it DIFFERS from what is already
   * there.
   *
   * Without the comparison every recalculation allocates new row arrays for
   * every formula cell, so the document is a new object even when nothing
   * changed — which makes the identity contract in this function's own doc
   * comment false, and costs a React re-render of the grid on every pass.
   */
  const write = (sheetIndex: number, r: number, c: number, text: string): void => {
    if (rowsOf(sheetIndex)[r]?.[c] === text) return
    ensureRowWritable(sheetIndex, r)[c] = text
  }

  for (const entry of order) {
    const { sheetIndex, row: r, col: c, formula } = entry
    if (cyclic.has(documentCellKey(sheetIndex, r, c))) {
      write(sheetIndex, r, c, CIRCULAR_REFERENCE_ERROR)
      continue
    }

    const lookup: CellLookup = (row, col, refSheet) => {
      const index = sheetIndexFor(refSheet, sheetIndex)
      if (index === null) return null
      return rowsOf(index)[row]?.[col] ?? ''
    }

    const result = evaluateFormula(formula, lookup)
    if (result.ok) {
      // SHEETFN-3 — the CELL's own number format wins over the evaluator's
      // rendering, so `=A1*1.2` in a currency column reads as currency and a
      // date formula follows the workbook's own date convention rather than
      // ISO. Falls back to the evaluator's text when the cell has no format or
      // the code cannot be applied.
      const formatted =
        result.value === undefined
          ? null
          : applyNumberFormat(result.value, formatForCell(doc.sheets[sheetIndex], r, c).numberFormat)
      write(sheetIndex, r, c, formatted ?? result.text)
    } else if (rowsOf(sheetIndex)[r][c] === '') {
      write(sheetIndex, r, c, `=${formula}`)
    }
  }

  if (working.size === 0) return doc
  return {
    sheets: doc.sheets.map((sheet, index) => {
      const rows = working.get(index)
      return rows === undefined ? sheet : { ...sheet, rows }
    }),
  }
}

/** SHEET-SORT-1 — why a sort did nothing, so the UI can say something true. */
export type SortRefusal = 'no-rows' | 'has-formulas'

/**
 * SHEET-SORT-2 — a sort that went ahead but changed what a formula ELSEWHERE
 * reads.
 *
 * Excel does not adjust a formula outside the sorted block, and neither does
 * this: `=B3` keeps pointing at B3, which now holds a different row's value.
 * That is correct behaviour and a well-known way to break a spreadsheet
 * quietly, so Atlas says it happened.
 */
export type SortWarning = 'outside-formulas-read-block'

export type SortRowsResult = {
  readonly document: SpreadsheetDocument
  readonly sorted: boolean
  readonly refusal?: SortRefusal
  readonly warning?: SortWarning
  /** The rows actually sorted, for the message that reports it. */
  readonly block?: { readonly first: number; readonly last: number }
}

/** A row with nothing in it — the boundary of a data block, as Excel's current region treats it. */
function isBlankRow(row: ReadonlyArray<string> | undefined): boolean {
  return row === undefined || row.every((cell) => cell.trim() === '')
}

/**
 * SHEET-SORT-2 — the contiguous run of non-blank rows containing `atRow`.
 *
 * This is what Excel calls the current region, and adopting it is what lets a
 * sheet with a totals row be sorted at all: `=SUM(B1:B9)` sitting under a blank
 * row is OUTSIDE the block, so it neither blocks the sort nor moves with it.
 * Sorting the whole sheet — which is all the first version could do — had to
 * refuse such a sheet entirely, because the totals row held a formula.
 *
 * Returns `null` when `atRow` is itself blank: there is no data block there to
 * sort, and picking a neighbouring one would be guessing which.
 */
export function dataBlockForRow(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  atRow: number,
  headerRows: number,
): { readonly first: number; readonly last: number } | null {
  if (atRow < headerRows || atRow >= rows.length || isBlankRow(rows[atRow])) return null
  let first = atRow
  while (first > headerRows && !isBlankRow(rows[first - 1])) first -= 1
  let last = atRow
  while (last + 1 < rows.length && !isBlankRow(rows[last + 1])) last += 1
  return { first, last }
}

/**
 * Does any formula OUTSIDE the block read rows INSIDE it, in a way a
 * permutation would change?
 *
 * A formula that reads the WHOLE block is not reported. That is the totals row
 * — `=SUM(B1:B9)` over exactly the rows being sorted — and its value does not
 * depend on their order, so warning about it would fire on almost every sort
 * and teach the user to ignore the message. A formula reading SOME of the
 * block's rows (`=B3`, `=AVERAGE(B2:B4)`) is reported: after the sort those
 * cells hold different rows' data.
 *
 * The scan deliberately over-reports in one direction — `formulaDependencyKeys`
 * is a text scan returning a superset of the true references, and a lookup like
 * `VLOOKUP` over the whole block IS order-sensitive but is not reported by the
 * whole-block rule above. The second of those is a judged trade, written down
 * in `docs/KNOWN_LIMITATIONS.md` rather than left implicit.
 */
function outsideFormulasReadBlock(
  doc: SpreadsheetDocument,
  sheetIndex: number,
  first: number,
  last: number,
): boolean {
  const indexByName = new Map<string, number>()
  doc.sheets.forEach((sheet, index) => {
    const key = sheet.name.toLowerCase()
    if (!indexByName.has(key)) indexByName.set(key, index)
  })
  const blockRowCount = last - first + 1

  for (let si = 0; si < doc.sheets.length; si++) {
    const sheet = doc.sheets[si]
    for (let r = 0; r < sheet.formulas.length; r++) {
      // A formula inside the block is handled by the refusal, not here.
      if (si === sheetIndex && r >= first && r <= last) continue
      const formulaRow = sheet.formulas[r]
      for (let c = 0; c < formulaRow.length; c++) {
        const formula = formulaRow[c]
        if (formula === undefined) continue
        const rowsRead = new Set<number>()
        formulaDependencyKeys(formula, sheet.rows.length, sheet.colCount, (refSheet, row, col) => {
          const target = refSheet === undefined ? si : (indexByName.get(refSheet.toLowerCase()) ?? -1)
          if (target === sheetIndex && row >= first && row <= last) rowsRead.add(row)
          return cellKey(row, col)
        })
        if (rowsRead.size > 0 && rowsRead.size < blockRowCount) return true
      }
    }
  }
  return false
}

/**
 * SHEET-SORT-1 — sorts a block of rows by one column.
 *
 * Moves each row WHOLE: its values, its formatting overrides, its height, and
 * its `rowSources` entry, so the save path still writes every untouched cell
 * through the original package and a cell's colour follows its data.
 *
 * **Refuses when any cell in the block holds a formula**, and says so. This is
 * a deliberate limit rather than an oversight. Excel adjusts a moved formula's
 * relative references so it keeps pointing at its own row, and the exact rules
 * for a reference that leaves the sorted block are subtle enough that
 * implementing them from memory would risk silently producing a workbook with
 * wrong numbers in it — the worst outcome this app can have. Sorting a block of
 * plain data, which is what the overwhelming majority of sorts are, is safe and
 * works. See `docs/KNOWN_LIMITATIONS.md`.
 *
 * `headerRows` are left in place at the top of the block.
 */
export function sortRows(
  doc: SpreadsheetDocument,
  sheetIndex: number,
  options: {
    readonly col: number
    readonly direction: 'asc' | 'desc'
    /** Rows above this many, from the top of the sheet, are not moved. */
    readonly headerRows?: number
    /**
     * SHEET-SORT-2 — sort only the data block containing this row, as Excel's
     * current region does. Omitted, the whole sheet below the header is one
     * block, which is what the first version always did.
     */
    readonly atRow?: number
  },
): SortRowsResult {
  const sheet = doc.sheets[sheetIndex]
  if (!sheet) return { document: doc, sorted: false, refusal: 'no-rows' }

  const headerRows = Math.max(0, options.headerRows ?? 0)
  const bounds =
    options.atRow === undefined
      ? { first: headerRows, last: sheet.rows.length - 1 }
      : dataBlockForRow(sheet.rows, options.atRow, headerRows)
  if (!bounds) return { document: doc, sorted: false, refusal: 'no-rows' }
  const { first, last } = bounds
  if (last <= first) return { document: doc, sorted: false, refusal: 'no-rows' }

  for (let r = first; r <= last; r += 1) {
    const formulaRow = sheet.formulas[r]
    if (formulaRow?.some(formula => formula !== undefined)) {
      return { document: doc, sorted: false, refusal: 'has-formulas', block: { first, last } }
    }
  }

  const warning: SortWarning | undefined = outsideFormulasReadBlock(doc, sheetIndex, first, last)
    ? 'outside-formulas-read-block'
    : undefined

  // Sorted as a list of row INDEXES, so every parallel array (heights,
  // sources, the override map) can be permuted the same way rather than each
  // being re-derived and risking going out of step.
  const order = Array.from({ length: last - first + 1 }, (_, i) => first + i)
  order.sort((ra, rb) => {
    const cmp = compareForSort(
      sheet.rows[ra]?.[options.col] ?? '',
      sheet.rows[rb]?.[options.col] ?? '',
      options.direction,
    )
    // A stable tie-break on the original position, so rows that compare equal
    // keep their relative order — and so a descending sort does not reverse
    // them, which `sort`'s own instability would otherwise do.
    return cmp === 0 ? ra - rb : cmp
  })

  if (order.every((from, i) => from === first + i)) {
    // Already in order: the same document object, so this costs nothing and
    // does not add an undo entry. No warning either — nothing moved, so
    // nothing an outside formula reads has changed.
    return { document: doc, sorted: false, refusal: undefined, block: { first, last } }
  }

  const permute = <T,>(source: ReadonlyArray<T>): T[] => {
    const next = [...source]
    order.forEach((from, i) => {
      next[first + i] = source[from]
    })
    return next
  }

  const rows = permute(sheet.rows)
  const formulas = permute(sheet.formulas)
  const rowHeightsPx = sheet.rowHeightsPx.length > 0 ? permute(sheet.rowHeightsPx) : sheet.rowHeightsPx
  const rowSources = sheet.rowSources ? permute(sheet.rowSources) : undefined

  // The override map is keyed by current position, so it is rebuilt through the
  // same permutation rather than permuted in place.
  const newIndexOfOldRow = new Map<number, number>()
  order.forEach((from, i) => newIndexOfOldRow.set(from, first + i))
  const formatOverrides = remapFormatOverrides(sheet.formatOverrides, (r, c) => {
    const moved = newIndexOfOldRow.get(r)
    return [moved ?? r, c]
  })
  const editedCells = remapEditedCells(sheet.editedCells, (r, c) => {
    const moved = newIndexOfOldRow.get(r)
    return [moved ?? r, c]
  })

  return {
    document: replaceSheetAndRecalculate(doc, sheetIndex, {
      ...sheet,
      rows,
      formulas,
      rowHeightsPx,
      ...(rowSources ? { rowSources } : {}),
      ...(formatOverrides ? { formatOverrides } : {}),
      ...(editedCells ? { editedCells } : {}),
    }),
    sorted: true,
    ...(warning ? { warning } : {}),
    block: { first, last },
  }
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
  const updated = { ...sheet, rows, formulas, ...(editedCells ? { editedCells } : {}) }
  return replaceSheetAndRecalculate(doc, sheetIndex, updated)
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
  const formatOverrides = remapFormatOverrides(sheet.formatOverrides, (r, c) => [r >= atIndex ? r + 1 : r, c])
  return replaceSheetAndRecalculate(
    doc,
    sheetIndex,
    {
      ...sheet,
      rows,
      formulas,
      rowHeightsPx,
      merges,
      ...tables,
      ...(sources ? { rowSources: sources } : {}),
      ...(editedCells ? { editedCells } : {}),
      ...(formatOverrides ? { formatOverrides } : {}),
    },
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
  const formatOverrides = remapFormatOverrides(sheet.formatOverrides, (r, c) =>
    r === atIndex ? null : [r > atIndex ? r - 1 : r, c],
  )
  return replaceSheetAndRecalculate(
    doc,
    sheetIndex,
    {
      ...sheet,
      rows,
      formulas,
      rowHeightsPx,
      merges,
      ...tables,
      ...(rowSources ? { rowSources } : {}),
      ...(editedCells ? { editedCells } : {}),
      ...(formatOverrides ? { formatOverrides } : {}),
    },
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

  return replaceSheetAndRecalculate(
    doc,
    sheetIndex,
    {
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
        const formatOverrides = remapFormatOverrides(sheet.formatOverrides, (r, c) => [r, c >= atIndex ? c + 1 : c])
        return { ...(editedCells ? { editedCells } : {}), ...(formatOverrides ? { formatOverrides } : {}) }
      })(),
    },
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

  return replaceSheetAndRecalculate(
    doc,
    sheetIndex,
    {
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
        const formatOverrides = remapFormatOverrides(sheet.formatOverrides, (r, c) =>
          c === atIndex ? null : [r, c > atIndex ? c - 1 : c],
        )
        return { ...(editedCells ? { editedCells } : {}), ...(formatOverrides ? { formatOverrides } : {}) }
      })(),
    },
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

  return replaceSheetAndRecalculate(
    doc,
    sheetIndex,
    {
      ...sheet,
      rows,
      formulas,
      colCount,
      colWidthsPx,
      rowHeightsPx,
      ...(rowSources ? { rowSources } : {}),
      ...(colSources ? { colSources } : {}),
      ...(editedCells ? { editedCells } : {}),
    },
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

  return replaceSheetAndRecalculate(doc, sheetIndex, { ...sheet, name: trimmed })
}

/** Deletes a sheet. A no-op if it's the only sheet left (a workbook must keep at least one). */
export function deleteSheet(doc: SpreadsheetDocument, sheetIndex: number): SpreadsheetDocument {
  if (doc.sheets.length <= 1 || !doc.sheets[sheetIndex]) return doc
  return { sheets: doc.sheets.filter((_, i) => i !== sheetIndex) }
}
