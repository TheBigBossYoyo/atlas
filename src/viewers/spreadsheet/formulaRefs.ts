/**
 * Re-anchoring OOXML formula-reference syntax through a renamed, reordered or
 * deleted sheet, and through a row/column insert or delete (USR-17 follow-up).
 *
 * `spreadsheetRangeShift.ts` already re-anchors plain `sqref`/`ref`-style
 * ranges (conditional formatting, data validation, hyperlinks, autoFilter);
 * `xlsxPassthrough.ts` already fixes up a single-area `<definedName>`. What
 * neither touches is the actual FORMULA TEXT stored in a cell's `<f>` (or a
 * `<definedName>`'s own, richer formula body) — `=Sheet2!A1`, `='My
 * Sheet'!A1:B2`, `=SUM(Sheet1:Sheet3!A1)`, `=Sheet1!$A$1,Sheet1!$C$3`,
 * `=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)`. Left alone, these keep
 * pointing at the OLD sheet name or the OLD cell position after a save that
 * renamed/reordered/deleted a sheet or shifted rows/columns — a silent wrong
 * number, not a repair prompt, which is worse.
 *
 * This module is a small, self-contained reference-syntax scanner (NOT a
 * full formula parser/evaluator — see `spreadsheetFormula.ts` for that, a
 * different piece of code with a different job). It walks a formula (or a
 * `<definedName>` body, which uses the identical reference syntax) once,
 * left to right, and rewrites every reference it recognizes while copying
 * everything else through byte-for-byte:
 *
 *  - a string literal (`"..."`, with `""` as an escaped quote) is copied
 *    through untouched — a reference-shaped substring inside one (`"A1"`) is
 *    text, not a reference;
 *  - a structured table reference (`Table1[Column]`, `[@Column]`,
 *    `Table1[[#Headers],[Column1]]`) is copied through untouched — its `[`
 *    contents are never scanned for references at all, since a column named
 *    e.g. "A1" would otherwise be misread as a cell reference;
 *  - a bare defined name (`MyRange`, not immediately followed by `!` or
 *    shaped like a cell reference) is copied through untouched — Excel
 *    itself forbids naming a defined name anything that WOULD parse as a
 *    valid cell reference, so there is no real ambiguity here: anything that
 *    parses as a reference below genuinely is one;
 *  - a sheet-qualified reference (`Sheet1!A1`, `'My Sheet'!A1:B2`, or a 3-D
 *    span `Sheet1:Sheet3!A1`) has its sheet name(s) rewritten for a rename,
 *    or the WHOLE reference replaced with `#REF!` for a delete — matching
 *    what Excel itself writes when the sheet a formula pointed at is gone;
 *  - a cell reference, a rectangular range, a whole-column range (`A:C`) or
 *    a whole-row range (`1:5`) — qualified or not — has its row/column
 *    numbers re-anchored through that sheet's insert/delete history exactly
 *    the way `spreadsheetRangeShift.ts` re-anchors an `sqref`; a range fully
 *    consumed by a delete becomes `#REF!` too.
 *
 * Two different call sites need two different amounts of rewriting, which is
 * why `remapCoordinates` is the caller's choice rather than inferred here:
 *
 *  - a formula CLONED byte-for-byte from the original file (an untouched
 *    cell `xlsxPassthrough.ts` copies as-is) is still expressed in the
 *    ORIGINAL file's row/column numbering, so both its sheet names AND its
 *    coordinates need re-anchoring (`remapCoordinates: true`);
 *  - a formula the user actually TYPED in Atlas (or a formula cell that was
 *    otherwise rebuilt from the live model) is already expressed against the
 *    CURRENT, saved layout — the editor only ever shows the live grid, never
 *    the original file's numbering — so re-anchoring its coordinates again
 *    would double-shift a cell that never moved. Its sheet name(s) can still
 *    go stale, though: `spreadsheetDocument.ts` never rewrites one sheet's
 *    already-stored formula text when ANOTHER sheet is renamed or deleted
 *    later in the same editing session (an explicit, documented scope cut —
 *    see that module's header), so the sheet-name half of this rewrite still
 *    applies (`remapCoordinates: false`, sheet renames/deletes still fixed).
 *
 * Deliberate limits, matching the module's own scope:
 *  - an UNQUALIFIED reference always means "this formula's own sheet" — its
 *    row/column re-anchoring (when `remapCoordinates` is true) uses the
 *    `ownRowSources`/`ownColSources` the caller passes for THAT sheet;
 *  - a 3-D reference's row/column re-anchoring (rare in practice, and rarer
 *    still combined with a row/column edit) uses the FIRST sheet's own
 *    insert/delete history for both endpoints, on the assumption implicit in
 *    a 3-D reference itself — that every sheet it spans is laid out the same
 *    way. If either endpoint sheet was deleted, the whole reference becomes
 *    `#REF!` rather than attempting to shrink the span to the sheets that
 *    remain (a real spreadsheet application sometimes does the latter; this
 *    is the same conservative, documented choice as "sheet deleted ->
 *    #REF!" elsewhere in this module, and never leaves a formula pointing at
 *    a sheet that no longer exists);
 *  - an external-workbook reference (`[1]Sheet1!A1`, or quoted
 *    `'[1]Sheet1'!A1`) is left alone: its sheet name never matches anything
 *    in `changesByOriginalName` (a map of THIS workbook's own original sheet
 *    names), so the lookup misses and the reference passes through
 *    unchanged — a safe default, not a special case;
 *  - R1C1-style references do not appear here at all: OOXML always stores
 *    formula text in A1 style regardless of the user's locale/display
 *    settings, so there is nothing else to parse.
 */
import { columnIndexToLetters, columnLettersToIndex } from './cellRef'
import { remapRange, type IndexSources } from './spreadsheetRangeShift'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** One sheet's structural fate between the original file and the document being saved. */
export type SheetChange = {
  /** `undefined` when the sheet still exists but kept its original name. */
  readonly newName: string | undefined
  readonly deleted: boolean
  readonly rowSources: IndexSources
  readonly colSources: IndexSources
}

export type FormulaRewriteOptions = {
  /** Every ORIGINAL sheet's fate, keyed by its ORIGINAL name (before any rename). */
  readonly changesByOriginalName: ReadonlyMap<string, SheetChange>
  /**
   * This formula's own sheet's insert/delete history, used ONLY to
   * re-anchor an UNQUALIFIED reference (`A1`, not `Sheet1!A1`). Pass
   * `undefined` for both when the formula has no single "own sheet" to fall
   * back to (a workbook-scoped defined name with no `localSheetId`) or when
   * `remapCoordinates` is `false` (they are never consulted in that case).
   */
  readonly ownRowSources?: IndexSources
  readonly ownColSources?: IndexSources
  /** See the module header: `true` for a formula cloned from the original file, `false` for one typed/rebuilt in Atlas. */
  readonly remapCoordinates: boolean
}

// ---------------------------------------------------------------------------
// Character classification
// ---------------------------------------------------------------------------

/** A character that continues an identifier (sheet name, defined name, table/column name) once started. Non-ASCII letters are included so an accented sheet name is treated as one token, not scanned char-by-char. */
function isIdentContinue(ch: string | undefined): boolean {
  if (ch === undefined) return false
  return /[A-Za-z0-9_.]/.test(ch) || ch.charCodeAt(0) > 127
}

function isIdentStart(ch: string | undefined): boolean {
  if (ch === undefined) return false
  return /[A-Za-z_]/.test(ch) || ch.charCodeAt(0) > 127
}

// ---------------------------------------------------------------------------
// Sticky-regex token matchers (each tried only at an exact position — no
// scanning/backtracking across the rest of the formula).
// ---------------------------------------------------------------------------

const RE_BARE_IDENT = /[A-Za-z_][A-Za-z0-9_.]*/y
const RE_QUOTED_SHEET = /'(?:[^']|'')*'/y
const RE_EXTERNAL_MARKER = /\[[0-9]+\]/y
const RE_CELL_RANGE = /\$?[A-Za-z]{1,3}\$?[0-9]+:\$?[A-Za-z]{1,3}\$?[0-9]+/y
const RE_COL_RANGE = /\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/y
const RE_ROW_RANGE = /\$?[0-9]+:\$?[0-9]+/y
const RE_SINGLE_CELL = /\$?[A-Za-z]{1,3}\$?[0-9]+/y

function execAt(re: RegExp, text: string, at: number): string | null {
  re.lastIndex = at
  const match = re.exec(text)
  return match && match.index === at ? match[0] : null
}

// ---------------------------------------------------------------------------
// $A$1-style coordinate decode/encode, generalized to a plain cell, a
// column-only token (`A`, `$A` — only meaningful as one side of `A:C`) and a
// row-only token (`1`, `$1` — only meaningful as one side of `1:5`).
// ---------------------------------------------------------------------------

type DollarCell = { readonly row: number; readonly col: number; readonly colAbs: boolean; readonly rowAbs: boolean }
type DollarCol = { readonly col: number; readonly abs: boolean }
type DollarRow = { readonly row: number; readonly abs: boolean }

function decodeDollarCell(text: string): DollarCell | null {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/.exec(text)
  if (!match) return null
  return {
    colAbs: match[1] === '$',
    col: columnLettersToIndex(match[2]),
    rowAbs: match[3] === '$',
    row: Number(match[4]) - 1,
  }
}

function encodeDollarCell(c: DollarCell): string {
  return `${c.colAbs ? '$' : ''}${columnIndexToLetters(c.col)}${c.rowAbs ? '$' : ''}${c.row + 1}`
}

function decodeDollarCol(text: string): DollarCol | null {
  const match = /^(\$?)([A-Za-z]{1,3})$/.exec(text)
  if (!match) return null
  return { abs: match[1] === '$', col: columnLettersToIndex(match[2]) }
}

function encodeDollarCol(c: DollarCol): string {
  return `${c.abs ? '$' : ''}${columnIndexToLetters(c.col)}`
}

function decodeDollarRow(text: string): DollarRow | null {
  const match = /^(\$?)([0-9]+)$/.exec(text)
  if (!match) return null
  return { abs: match[1] === '$', row: Number(match[2]) - 1 }
}

function encodeDollarRow(r: DollarRow): string {
  return `${r.abs ? '$' : ''}${r.row + 1}`
}

/** `#REF!` — what Excel itself writes in place of a reference it can no longer resolve. */
const REF_ERROR = '#REF!'

/**
 * Re-anchors one coordinate span's text through `rowSources`/`colSources`.
 * `shape` picks which sides of the span carry row information, column
 * information, or both, and how to re-encode the result. Returns `#REF!`
 * when the span was entirely consumed by a delete.
 */
function remapCoordinateText(rawText: string, rowSources: IndexSources, colSources: IndexSources): string {
  const [startText, endText] = rawText.split(':')

  if (endText === undefined) {
    // Single cell.
    const cell = decodeDollarCell(startText)
    if (!cell) return rawText // shouldn't happen (caller only reaches here after a shape match) — pass through rather than corrupt
    const mapped = remapRange({ r0: cell.row, c0: cell.col, r1: cell.row, c1: cell.col }, rowSources, colSources)
    if (!mapped) return REF_ERROR
    return encodeDollarCell({ row: mapped.r0, col: mapped.c0, rowAbs: cell.rowAbs, colAbs: cell.colAbs })
  }

  const startCell = decodeDollarCell(startText)
  const endCell = decodeDollarCell(endText)
  if (startCell && endCell) {
    const range = {
      r0: Math.min(startCell.row, endCell.row),
      c0: Math.min(startCell.col, endCell.col),
      r1: Math.max(startCell.row, endCell.row),
      c1: Math.max(startCell.col, endCell.col),
    }
    const mapped = remapRange(range, rowSources, colSources)
    if (!mapped) return REF_ERROR
    const newStart = encodeDollarCell({ row: mapped.r0, col: mapped.c0, rowAbs: startCell.rowAbs, colAbs: startCell.colAbs })
    const newEnd = encodeDollarCell({ row: mapped.r1, col: mapped.c1, rowAbs: endCell.rowAbs, colAbs: endCell.colAbs })
    return `${newStart}:${newEnd}`
  }

  const startCol = decodeDollarCol(startText)
  const endCol = decodeDollarCol(endText)
  if (startCol && endCol) {
    const c0 = Math.min(startCol.col, endCol.col)
    const c1 = Math.max(startCol.col, endCol.col)
    // Rows are unaffected by a whole-column range — pass `undefined` (identity) for that axis.
    const mapped = remapRange({ r0: 0, c0, r1: 0, c1 }, undefined, colSources)
    if (!mapped) return REF_ERROR
    const newStart = encodeDollarCol({ col: mapped.c0, abs: startCol.abs })
    const newEnd = encodeDollarCol({ col: mapped.c1, abs: endCol.abs })
    return `${newStart}:${newEnd}`
  }

  const startRow = decodeDollarRow(startText)
  const endRow = decodeDollarRow(endText)
  if (startRow && endRow) {
    const r0 = Math.min(startRow.row, endRow.row)
    const r1 = Math.max(startRow.row, endRow.row)
    const mapped = remapRange({ r0, c0: 0, r1, c1: 0 }, rowSources, undefined)
    if (!mapped) return REF_ERROR
    const newStart = encodeDollarRow({ row: mapped.r0, abs: startRow.abs })
    const newEnd = encodeDollarRow({ row: mapped.r1, abs: endRow.abs })
    return `${newStart}:${newEnd}`
  }

  return rawText // unreachable given the shape matchers below, but never corrupt on a surprise input
}

// ---------------------------------------------------------------------------
// Structured table references (`Table1[Column]`, `[@Column]`) — recognized
// and skipped WITHOUT looking at their contents at all.
// ---------------------------------------------------------------------------

/** Scans a bracket-matched `[...]` span starting exactly at `openAt` (must be `[`). Returns the index right after the matching `]`, or `null` if unterminated. */
function scanBracketSpan(formula: string, openAt: number): number | null {
  let depth = 0
  for (let i = openAt; i < formula.length; i++) {
    if (formula[i] === '[') depth += 1
    else if (formula[i] === ']') {
      depth -= 1
      if (depth === 0) return i + 1
    }
  }
  return null // unterminated — malformed formula; caller falls back to copying char-by-char
}

/** Matches `Identifier[...]` or a bare `[...]` (an in-table calculated-column reference, e.g. `[@Column]`) starting at `at`. Returns the end index, or `null` if this isn't one. */
function matchStructuredRefEnd(formula: string, at: number): number | null {
  let pos = at
  if (isIdentStart(formula[pos])) {
    const ident = execAt(RE_BARE_IDENT, formula, pos)
    if (!ident) return null
    pos += ident.length
  }
  if (formula[pos] !== '[') return null
  return scanBracketSpan(formula, pos)
}

/**
 * Matches an external-workbook reference marker (`[1]`, as in `[1]Sheet1!A1`
 * — the numeric index Excel assigns another open/linked workbook) plus the
 * sheet-qualified reference that follows it, starting at `at`. The whole
 * span is treated as OPAQUE (copied through untouched, never rewritten):
 * the sheet name that follows belongs to a DIFFERENT workbook this module
 * has no information about, and could coincidentally match one of THIS
 * workbook's own original sheet names, which would be a genuine (if rare)
 * corruption risk if left unguarded. Returns just the marker's own end when
 * a full reference does not follow (an unusual/malformed formula — safe
 * either way, since the marker alone is never mistaken for a local sheet
 * reference by `matchReferenceAt`).
 */
function matchExternalRefEnd(formula: string, at: number): number | null {
  const marker = execAt(RE_EXTERNAL_MARKER, formula, at)
  if (!marker) return null
  const markerEnd = at + marker.length

  const first = matchSheetNameAt(formula, markerEnd)
  if (!first) return markerEnd
  if (formula[first.end] === '!') {
    const shape = matchShapeAt(formula, first.end + 1)
    if (shape) return shape.end
  } else if (formula[first.end] === ':') {
    const second = matchSheetNameAt(formula, first.end + 1)
    if (second && formula[second.end] === '!') {
      const shape = matchShapeAt(formula, second.end + 1)
      if (shape) return shape.end
    }
  }
  return markerEnd
}

// ---------------------------------------------------------------------------
// Reference matching
// ---------------------------------------------------------------------------

type ShapeMatch = { readonly text: string; readonly end: number }

function matchShapeAt(formula: string, at: number): ShapeMatch | null {
  for (const re of [RE_CELL_RANGE, RE_COL_RANGE, RE_ROW_RANGE, RE_SINGLE_CELL]) {
    const text = execAt(re, formula, at)
    if (text && !isIdentContinue(formula[at + text.length])) return { text, end: at + text.length }
  }
  return null
}

type SheetToken = { readonly name: string; readonly end: number }

/** Matches one (bare or quoted) sheet name at `at`, without requiring what follows it. */
function matchSheetNameAt(formula: string, at: number): SheetToken | null {
  if (formula[at] === "'") {
    const quoted = execAt(RE_QUOTED_SHEET, formula, at)
    if (!quoted) return null
    return { name: quoted.slice(1, -1).replace(/''/g, "'"), end: at + quoted.length }
  }
  if (!isIdentStart(formula[at])) return null
  const bare = execAt(RE_BARE_IDENT, formula, at)
  if (!bare) return null
  return { name: bare, end: at + bare.length }
}

type ReferenceMatch = {
  readonly end: number
  /** One sheet name for a plain reference, two for a 3-D span (`Sheet1:Sheet3!...`); absent for an unqualified reference. */
  readonly sheets?: readonly [string] | readonly [string, string]
  readonly rangeText: string
}

/**
 * Matches one full reference — sheet-qualified (single or 3-D) or bare —
 * starting exactly at `at`. Returns `null` when nothing reference-shaped
 * starts there (the caller then falls back to copying a single character,
 * the normal path for a defined name, a function name, an operator, ...).
 */
function matchReferenceAt(formula: string, at: number): ReferenceMatch | null {
  // Sheet-qualified: try a first sheet name, then either "!range" or ":secondSheet!range".
  const first = matchSheetNameAt(formula, at)
  if (first) {
    if (formula[first.end] === '!') {
      const shape = matchShapeAt(formula, first.end + 1)
      if (shape) return { end: shape.end, sheets: [first.name], rangeText: shape.text }
    } else if (formula[first.end] === ':') {
      const second = matchSheetNameAt(formula, first.end + 1)
      if (second && formula[second.end] === '!') {
        const shape = matchShapeAt(formula, second.end + 1)
        if (shape) return { end: shape.end, sheets: [first.name, second.name], rangeText: shape.text }
      }
    }
    // A bare identifier that turned out not to be a sheet-qualifying prefix
    // after all (e.g. a defined name, or a function name before "(") falls
    // through to the unqualified attempt below, which will also fail for it
    // (an identifier never matches the digit-anchored cell/range shapes) —
    // so it is correctly left for the generic char-by-char copy.
  }

  const shape = matchShapeAt(formula, at)
  return shape ? { end: shape.end, rangeText: shape.text } : null
}

// ---------------------------------------------------------------------------
// Rewriting one matched reference
// ---------------------------------------------------------------------------

function needsQuoting(sheetName: string): boolean {
  return !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")}'`
}

function rewriteReference(match: ReferenceMatch, options: FormulaRewriteOptions): string {
  if (!match.sheets) {
    // Unqualified — always this formula's own sheet.
    if (!options.remapCoordinates) return match.rangeText
    if (!options.ownRowSources && !options.ownColSources) return match.rangeText
    return remapCoordinateText(match.rangeText, options.ownRowSources, options.ownColSources)
  }

  const changes = match.sheets.map((name) => options.changesByOriginalName.get(name))
  // A name the map has never heard of is either already-current (the map is
  // keyed by ORIGINAL names, so an up-to-date reference simply never
  // matches) or genuinely external — either way, safe to leave untouched.
  if (changes.some((c) => c === undefined)) {
    const prefix = match.sheets.length === 1 ? formatSheetPrefix(match.sheets[0]) : formatSpanPrefix(match.sheets[0], match.sheets[1])
    return `${prefix}!${match.rangeText}`
  }
  const resolved = changes as SheetChange[]
  if (resolved.some((c) => c.deleted)) return REF_ERROR

  const newNames = match.sheets.map((name, i) => resolved[i].newName ?? name)
  const prefix = newNames.length === 1 ? formatSheetPrefix(newNames[0]) : formatSpanPrefix(newNames[0], newNames[1])

  if (!options.remapCoordinates) return `${prefix}!${match.rangeText}`

  // A 3-D span's shift is taken from its FIRST sheet (see module header).
  const shiftSource = resolved[0]
  const rangeText = remapCoordinateText(match.rangeText, shiftSource.rowSources, shiftSource.colSources)
  return rangeText === REF_ERROR ? REF_ERROR : `${prefix}!${rangeText}`
}

function formatSheetPrefix(name: string): string {
  return needsQuoting(name) ? quoteSheetName(name) : name
}

function formatSpanPrefix(name1: string, name2: string): string {
  if (!needsQuoting(name1) && !needsQuoting(name2)) return `${name1}:${name2}`
  // A 3-D span quotes the WHOLE "Name1:Name2" as one string when either side needs it — not each name separately.
  return `'${name1.replace(/'/g, "''")}:${name2.replace(/'/g, "''")}'`
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Rewrites every sheet-qualified and unqualified cell/range reference in
 * `formula` per `options` (see module header). Safe to call on a formula
 * with nothing to rewrite — it is then returned unchanged (same string
 * reference, no allocation beyond the scan itself is observable).
 */
export function rewriteFormulaReferences(formula: string, options: FormulaRewriteOptions): string {
  let out = ''
  let changed = false
  let i = 0
  const n = formula.length

  while (i < n) {
    const ch = formula[i]

    if (ch === '"') {
      const start = i
      i += 1
      while (i < n) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') {
            i += 2
            continue
          }
          i += 1
          break
        }
        i += 1
      }
      out += formula.slice(start, i)
      continue
    }

    const precededByIdent = isIdentContinue(formula[i - 1])
    if (!precededByIdent) {
      const externalEnd = matchExternalRefEnd(formula, i)
      if (externalEnd !== null) {
        out += formula.slice(i, externalEnd)
        i = externalEnd
        continue
      }

      const structuredEnd = matchStructuredRefEnd(formula, i)
      if (structuredEnd !== null) {
        out += formula.slice(i, structuredEnd)
        i = structuredEnd
        continue
      }

      const ref = matchReferenceAt(formula, i)
      if (ref) {
        out += rewriteReference(ref, options)
        changed = true
        i = ref.end
        continue
      }
    }

    out += ch
    i += 1
  }

  return changed ? out : formula
}
