/**
 * Cell-to-display-text formatting for the SheetJS `xlsx` cell object (T1/DAT-04).
 *
 * `XLSX.utils.sheet_to_json` with the default `raw:true` returns each cell's
 * raw value (a serial number for dates, the numeric error code for a formula
 * error), which is why SpreadsheetViewer previously showed dates as bare
 * numbers and error cells as blank. `formatCellText` renders the same text a
 * spreadsheet application would show, cell by cell, using SheetJS's own
 * formatted-text cache (`cell.w`) whenever it exists.
 *
 * `cell.w` is populated by SheetJS itself at parse time for the overwhelming
 * majority of cells (formatted using the file's own number-format code), so
 * trusting it here also means Atlas renders whatever formatting the
 * spreadsheet's author actually applied. The Intl fallbacks below only run
 * for the rare case where SheeJS handed back no display text at all (see the
 * locale note).
 */
import type { CellObject } from 'xlsx'

// Mirrors XLSX.js's internal `BErr` formula-error code table (not exported
// by the package) — last-resort fallback for an error cell that somehow has
// neither `.w` nor a recognized code.
const FORMULA_ERROR_CODES: Readonly<Record<number, string>> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
  0x2b: '#GETTING_DATA',
  0xff: '#WTF?',
}

function formatFormulaError(code: unknown): string {
  if (typeof code === 'number' && code in FORMULA_ERROR_CODES) {
    return FORMULA_ERROR_CODES[code]
  }
  return typeof code === 'string' ? code : '#ERROR!'
}

/**
 * Renders one worksheet cell's display text.
 *
 * Order of preference:
 *  1. `cell.w` — SheetJS's own formatted text, reflecting the file's actual
 *     number format (dates, currency, percentages, ...). Always preferred
 *     when present so Atlas never second-guesses the author's formatting.
 *  2. A formula-error code (`cell.t === 'e'`) — rendered as its `#NAME?` /
 *     `#DIV/0!` / ... text instead of a blank cell (DAT-04).
 *  3. Locale note (critic-review DEFER-6): if SheetJS produced no `.w` at
 *     all for a date or number value, Atlas must synthesize display text
 *     itself — done via `Intl.DateTimeFormat`/`Intl.NumberFormat` with no
 *     explicit locale argument (defers to the OS locale) rather than
 *     hardcoding an assumed locale such as `en-US`.
 */
export function formatCellText(cell: CellObject | undefined): string {
  if (!cell || cell.t === 'z' || cell.v === undefined) {
    return ''
  }

  if (cell.t === 'e') {
    return cell.w ?? formatFormulaError(cell.v)
  }

  if (cell.w !== undefined) {
    return cell.w
  }

  if (cell.t === 'd' && cell.v instanceof Date) {
    return new Intl.DateTimeFormat().format(cell.v)
  }

  if (cell.t === 'n' && typeof cell.v === 'number') {
    return new Intl.NumberFormat().format(cell.v)
  }

  return String(cell.v)
}
