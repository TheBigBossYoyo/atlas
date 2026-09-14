/**
 * A1-style cell reference <-> zero-indexed (row, col) conversion, shared by
 * the formula evaluator and the editable-document model (spreadsheet
 * editing + save, wave 3).
 */

/** Converts a zero-indexed column number to its spreadsheet letter(s): 0->A, 25->Z, 26->AA. */
export function columnIndexToLetters(index: number): string {
  let letters = ''
  let n = index
  while (n >= 0) {
    letters = String.fromCharCode(65 + (n % 26)) + letters
    n = Math.floor(n / 26) - 1
  }
  return letters
}

/** Converts spreadsheet column letters (case-insensitive) to a zero-indexed column number: A->0, Z->25, AA->26. */
export function columnLettersToIndex(letters: string): number {
  let index = 0
  const upper = letters.toUpperCase()
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64)
  }
  return index - 1
}

export type CellCoord = { readonly row: number; readonly col: number }

const CELL_REF_PATTERN = /^\$?([A-Za-z]+)\$?(\d+)$/

/** Parses an (optionally `$`-anchored) A1-style reference, e.g. `B3` or `$B$3`. Returns `null` when malformed. */
export function parseCellRef(ref: string): CellCoord | null {
  const match = CELL_REF_PATTERN.exec(ref.trim())
  if (!match) return null
  const row = Number.parseInt(match[2], 10) - 1
  if (row < 0) return null
  return { row, col: columnLettersToIndex(match[1]) }
}

/** Formats a zero-indexed (row, col) coordinate as an A1-style reference, e.g. `{row:2,col:1}` -> `B3`. */
export function formatCellRef({ row, col }: CellCoord): string {
  return `${columnIndexToLetters(col)}${row + 1}`
}

export type CellRange = { readonly start: CellCoord; readonly end: CellCoord }

/** Parses an A1:B3-style range. Returns `null` when malformed; normalizes so `start` is top-left. */
export function parseCellRange(ref: string): CellRange | null {
  const [rawStart, rawEnd] = ref.split(':')
  if (!rawStart) return null
  const start = parseCellRef(rawStart)
  if (!start) return null
  if (rawEnd === undefined) {
    return { start, end: start }
  }
  const end = parseCellRef(rawEnd)
  if (!end) return null
  return {
    start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
    end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
  }
}
