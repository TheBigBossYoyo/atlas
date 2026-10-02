/**
 * SHEETFN-3 — applying a cell's own number format to a value.
 *
 * Atlas displays a formula's result as the evaluator rendered it, ignoring the
 * cell's number format entirely: `=A1*1.2` in a currency column showed `15.6`
 * where Excel shows `$15.60`, and `=TODAY()` in a `dd/mm/yyyy` column showed an
 * ISO date instead of the workbook's own convention.
 *
 * The formatting itself is SheetJS's `SSF`, which is already bundled (it is what
 * produces the display text for every value READ from a workbook, via
 * `formatCellText`). Using it here is what makes a formula result and a stored
 * value with the same format render identically — a second implementation of
 * OOXML format codes would differ from the one already in use, which is worse
 * than having none.
 */
import { SSF } from 'xlsx'

/**
 * `value` rendered with the OOXML `formatCode`, or `null` when the format
 * cannot be applied.
 *
 * `null` rather than a fallback string, so the caller keeps the evaluator's own
 * rendering: a format code SSF rejects should leave the number readable, not
 * replace it with an error or an empty cell.
 */
export function applyNumberFormat(value: number, formatCode: string | undefined): string | null {
  if (!formatCode || formatCode.toLowerCase() === 'general') return null
  try {
    const formatted = SSF.format(formatCode, value)
    return typeof formatted === 'string' ? formatted : null
  } catch {
    // A malformed or unsupported format code must not break recalculation.
    return null
  }
}
