/**
 * A deliberately small formula evaluator for spreadsheet editing (Wave 3).
 *
 * The bundled `xlsx` (SheetJS) package has no formula-calculation engine at
 * all in its community build — it only ever reads/writes whatever cached
 * `<v>` a real spreadsheet application already computed. So when the user
 * types `=...` into a cell in Atlas's own grid, nothing downstream will ever
 * "recompute" that value for us unless we do it ourselves.
 *
 * This module supports exactly the common subset spreadsheet users reach
 * for first: `+ - * / ^` arithmetic (with parens and unary minus), A1/`$A$1`
 * cell references, `A1:B3` ranges, string literals (`"like this"`, with a
 * doubled `""` for a literal quote), `&` text concatenation, and the handful
 * of aggregate/text functions listed in `FUNCTIONS` below. Anything outside
 * that — an unknown function name, a malformed range, a bare syntax error —
 * resolves to `null`, and callers (see `spreadsheetDocument.ts`) fall back to
 * displaying the literal formula text (`=<formula>`) instead of a fabricated
 * number. A *runtime* error within an otherwise-valid formula (division by
 * zero, a text operand in an arithmetic expression) still resolves to a real
 * spreadsheet error string (`#DIV/0!`, `#VALUE!`) — that IS a value a real
 * spreadsheet would show, unlike a fallback.
 *
 * SHEET-6 follow-up: `&` and text functions (`CONCATENATE`) used to be
 * unreachable — the tokenizer had no string-literal rule at all (a bare `"`
 * was an "unrecognized character" syntax error) and no `&` operator, so
 * `=A1&"!"` fell all the way through to the unsupported-formula case above
 * and displayed as literal formula text instead of an evaluated value. `&`
 * sits at the lowest precedence (below `+ - * /  ^`, same as real Excel),
 * parsed by `parseConcat` above `parseExpression` — `="1"&1+1` concatenates
 * "1" with the ALREADY-COMPUTED `1+1`, not `("1"&1)+1`.
 */
import { parseCellRef, type CellCoord } from './cellRef'
import { applyNumberFormat } from './formatNumber'
import {
  isoDateFromSerial,
  weekdayFromSerial,
  isoDateTimeFromSerial,
  nowSerial,
  serialFromDate,
  serialFromIsoDate,
  todaySerial,
} from './excelDate'

export type CellLookup = (row: number, col: number) => string

export type FormulaResult =
  | {
      readonly ok: true
      readonly text: string
      /**
       * SHEETFN-3 — the numeric result, when there is one, so the caller can
       * apply the CELL's own number format to it (`=A1*1.2` in a currency
       * column should read as currency). `text` is the evaluator's own
       * rendering, used when the cell has no format of its own.
       */
      readonly value?: number
      /** Set when `value` is a date serial — see `Value`'s own note. */
      readonly dateKind?: 'date' | 'datetime'
    }
  | { readonly ok: false }

type Token =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'string'; readonly text: string }
  | { readonly kind: 'ref'; readonly text: string }
  | { readonly kind: 'ident'; readonly text: string }
  | { readonly kind: 'op'; readonly text: string }
  | { readonly kind: 'lparen' }
  | { readonly kind: 'rparen' }
  | { readonly kind: 'comma' }
  | { readonly kind: 'colon' }

// String literals are matched FIRST so a quoted `&`, digit or letter inside
// one (`"A1 & B1"`) is never mistaken for a ref/op/ident token of its own —
// `"(?:[^"]|"")*"` accepts a doubled `""` as an escaped literal quote inside
// the string, the same convention Excel itself uses. `&` joins the operator
// character class alongside the existing arithmetic/comparison operators.
const TOKEN_PATTERN =
  /\s*(?:("(?:[^"]|"")*")|(\d+(?:\.\d+)?)|(\$?[A-Za-z]+\$?\d+)|([A-Za-z_][A-Za-z0-9_]*)|(<=|>=|<>|[-+*/^()=<>,:&])|(\S))/g

class FormulaSyntaxError extends Error {}

function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  TOKEN_PATTERN.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = TOKEN_PATTERN.exec(formula)) !== null) {
    const [, string, number, ref, ident, op, junk] = match
    if (string !== undefined) {
      // Strip the surrounding quotes, then unescape a doubled `""` to one `"`.
      tokens.push({ kind: 'string', text: string.slice(1, -1).replace(/""/g, '"') })
    } else if (number !== undefined) {
      tokens.push({ kind: 'number', value: Number.parseFloat(number) })
    } else if (ref !== undefined) {
      tokens.push({ kind: 'ref', text: ref })
    } else if (ident !== undefined) {
      tokens.push({ kind: 'ident', text: ident.toUpperCase() })
    } else if (op === '(') {
      tokens.push({ kind: 'lparen' })
    } else if (op === ')') {
      tokens.push({ kind: 'rparen' })
    } else if (op === ',') {
      tokens.push({ kind: 'comma' })
    } else if (op === ':') {
      tokens.push({ kind: 'colon' })
    } else if (op !== undefined) {
      tokens.push({ kind: 'op', text: op })
    } else if (junk !== undefined) {
      throw new FormulaSyntaxError(`Unrecognized character: ${junk}`)
    }
  }
  return tokens
}

/** A resolved value flowing through evaluation: either numeric, text, or a spreadsheet error code. */
/**
 * SHEETFN-2 — `dateKind` marks a number that MEANS a date.
 *
 * A spreadsheet date is a serial number, so `TODAY()` is really 46266. In Excel
 * the cell's number FORMAT is what makes that display as a date; Atlas does not
 * yet apply a cell's number format to a formula result (see
 * `docs/KNOWN_LIMITATIONS.md`), so without this a date formula would show its
 * serial and be useless.
 *
 * The flag rides on the value and is propagated by `+`/`-` the way Excel's own
 * type inference does — date + number is a date, date - date is a plain count
 * of days — so `=TODAY()+7` displays as a date while `=B1-A1` displays as a
 * number of days.
 */
type Value =
  | { readonly kind: 'number'; readonly value: number; readonly dateKind?: 'date' | 'datetime' }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'error'; readonly code: string }

function numberValue(n: number): Value {
  if (Number.isNaN(n)) return { kind: 'error', code: '#VALUE!' }
  if (!Number.isFinite(n)) return { kind: 'error', code: '#DIV/0!' }
  return { kind: 'number', value: n }
}

function toNumeric(value: Value): number | { readonly code: string } {
  if (value.kind === 'error') return { code: value.code }
  if (value.kind === 'number') return value.value
  // SHEETFN-1 — a comparison resolves to the text "TRUE"/"FALSE" (see
  // `boolValue`), and Excel treats those as 1/0 in arithmetic, so `=(A1>2)*5`
  // gives 5 rather than #VALUE!.
  const upper = value.value.trim().toUpperCase()
  if (upper === 'TRUE') return 1
  if (upper === 'FALSE') return 0
  // A blank cell is 0 in arithmetic, which is what lets `=A1+1` work on an
  // empty cell while `ISBLANK(A1)` still reports it as blank.
  if (upper === '') return 0
  const parsed = Number.parseFloat(value.value)
  return Number.isNaN(parsed) ? { code: '#VALUE!' } : parsed
}

/**
 * Renders a resolved value as the text `&`/`CONCATENATE` splice in: a number
 * formatted the same way a numeric formula RESULT is displayed elsewhere
 * (`formatNumericResult`, defined below — hoisted, so the forward reference
 * here is fine), text passed through as-is, and an error code rendered as
 * its own literal text (`#DIV/0!`) for the rare direct call — every actual
 * call site here (`applyConcat`, `CONCATENATE`) already propagates an error
 * operand instead of ever stringifying it, so this branch is defensive only.
 */
function valueToText(value: Value): string {
  if (value.kind === 'number') return formatNumericResult(value.value)
  if (value.kind === 'text') return value.value
  return value.code
}

/** Tags a numeric value as a date, leaving anything else alone. */
function dated(value: Value, dateKind: 'date' | 'datetime' | undefined): Value {
  if (dateKind === undefined || value.kind !== 'number') return value
  return { ...value, dateKind }
}

/**
 * The serial behind a value, for a date function's argument.
 *
 * Accepts a serial (what a date formula produces) AND an `YYYY-MM-DD` string
 * (what a cell holding a typed date shows), because both are what a user will
 * point `YEAR(...)` at and neither is obviously the "real" one from where they
 * are sitting.
 */
function toSerial(value: Value): number | { readonly code: string } {
  if (value.kind === 'error') return { code: value.code }
  if (value.kind === 'number') return value.value
  const fromIso = serialFromIsoDate(value.value)
  if (fromIso !== null) return fromIso
  return { code: '#VALUE!' }
}

const COMPARISON_OPS: ReadonlySet<string> = new Set(['=', '<>', '<', '<=', '>', '>='])

/**
 * Compares two values the way a spreadsheet does.
 *
 * Two rules worth stating, because both differ from JavaScript's own `==`:
 *   - a number is never equal to text, and in an ORDERING comparison every
 *     number sorts before every text value (Excel's documented type ordering),
 *     rather than the two being coerced into each other;
 *   - text compares case-INSENSITIVELY (`="a"="A"` is TRUE in Excel), which is
 *     the single most surprising difference for anyone expecting JS semantics.
 */
function compareValues(op: string, left: Value, right: Value): Value {
  if (left.kind === 'error') return left
  if (right.kind === 'error') return right

  let ordering: number
  if (left.kind === 'number' && right.kind === 'number') {
    ordering = left.value === right.value ? 0 : left.value < right.value ? -1 : 1
  } else if (left.kind === 'text' && right.kind === 'text') {
    const a = left.value.toUpperCase()
    const b = right.value.toUpperCase()
    ordering = a === b ? 0 : a < b ? -1 : 1
  } else {
    // Mixed types: numbers sort before text, and are never equal to it.
    ordering = left.kind === 'number' ? -1 : 1
  }

  switch (op) {
    case '=': return boolValue(ordering === 0)
    case '<>': return boolValue(ordering !== 0)
    case '<': return boolValue(ordering < 0)
    case '<=': return boolValue(ordering <= 0)
    case '>': return boolValue(ordering > 0)
    case '>=': return boolValue(ordering >= 0)
    default: throw new FormulaSyntaxError(`Unknown comparison: ${op}`)
  }
}

/**
 * A boolean result, as the TEXT "TRUE"/"FALSE".
 *
 * Excel has a real boolean type; this evaluator has number/text/error. Text is
 * the honest choice of the three: it displays as a user expects (`TRUE`), and
 * `toNumeric` below maps it back to 1/0 so `=(A1>2)*5` still arithmetically
 * works. Representing it as the number 1/0 instead would display `1` where
 * every spreadsheet shows `TRUE`.
 */
function boolValue(on: boolean): Value {
  return { kind: 'text', value: on ? 'TRUE' : 'FALSE' }
}

/** Reads a value's truthiness for `IF`/`AND`/`OR`: a number is false only at 0, and `"TRUE"`/`"FALSE"` round-trip. */
function truthy(value: Value): boolean | { readonly code: string } {
  if (value.kind === 'error') return { code: value.code }
  if (value.kind === 'number') return value.value !== 0
  const upper = value.value.trim().toUpperCase()
  if (upper === 'TRUE') return true
  if (upper === 'FALSE' || upper === '') return false
  const parsed = Number.parseFloat(value.value)
  if (!Number.isNaN(parsed)) return parsed !== 0
  return { code: '#VALUE!' }
}

/** A minimal recursive-descent parser/evaluator over the token stream produced by `tokenize`. */
class FormulaParser {
  private position = 0
  private readonly tokens: ReadonlyArray<Token>
  private readonly lookup: CellLookup

  // Explicit fields + assignment (not TS constructor-parameter-property
  // shorthand): this project's `erasableSyntaxOnly` tsconfig setting forbids
  // that shorthand since it requires emitting real field-assignment code
  // rather than being erasable at the type layer alone.
  constructor(tokens: ReadonlyArray<Token>, lookup: CellLookup) {
    this.tokens = tokens
    this.lookup = lookup
  }

  private peek(): Token | undefined {
    return this.tokens[this.position]
  }

  private consume(): Token {
    const token = this.tokens[this.position]
    if (!token) throw new FormulaSyntaxError('Unexpected end of formula')
    this.position += 1
    return token
  }

  parseTopLevel(): Value {
    const result = this.parseComparison()
    if (this.position !== this.tokens.length) {
      throw new FormulaSyntaxError('Unexpected trailing tokens')
    }
    return result
  }

  /**
   * SHEETFN-1 — `= <> < <= > >=`, the lowest-precedence operators, below `&`
   * exactly as in Excel (`="a"&"b"="ab"` compares the concatenation).
   *
   * The tokenizer has recognised these since the `&` work, but nothing ever
   * parsed them: `=1>0` fell through to a syntax error, and the whole
   * conditional half of a spreadsheet — IF, COUNTIF, any comparison at all —
   * was unreachable.
   */
  private parseComparison(): Value {
    let left = this.parseConcat()
    for (;;) {
      const token = this.peek()
      if (token?.kind !== 'op' || !COMPARISON_OPS.has(token.text)) break
      this.consume()
      const right = this.parseConcat()
      left = compareValues(token.text, left, right)
    }
    return left
  }

  /** `&` text concatenation — the lowest-precedence operator this evaluator supports, matching real Excel (`="1"&1+1` concatenates "1" with the already-computed `1+1`, not `("1"&1)+1`). */
  private parseConcat(): Value {
    let left = this.parseExpression()
    for (;;) {
      const token = this.peek()
      if (token?.kind !== 'op' || token.text !== '&') break
      this.consume()
      const right = this.parseExpression()
      left = this.applyConcat(left, right)
    }
    return left
  }

  private applyConcat(left: Value, right: Value): Value {
    if (left.kind === 'error') return left
    if (right.kind === 'error') return right
    return { kind: 'text', value: valueToText(left) + valueToText(right) }
  }

  private parseExpression(): Value {
    let left = this.parseTerm()
    for (;;) {
      const token = this.peek()
      if (token?.kind !== 'op' || (token.text !== '+' && token.text !== '-')) break
      this.consume()
      const right = this.parseTerm()
      left = this.applyArithmetic(token.text, left, right)
    }
    return left
  }

  private parseTerm(): Value {
    let left = this.parseUnary()
    for (;;) {
      const token = this.peek()
      if (token?.kind !== 'op' || (token.text !== '*' && token.text !== '/')) break
      this.consume()
      const right = this.parseUnary()
      left = this.applyArithmetic(token.text, left, right)
    }
    return left
  }

  private parseUnary(): Value {
    const token = this.peek()
    if (token?.kind === 'op' && (token.text === '-' || token.text === '+')) {
      this.consume()
      const operand = this.parseUnary()
      if (token.text === '-') {
        const n = toNumeric(operand)
        return typeof n === 'number' ? numberValue(-n) : { kind: 'error', code: n.code }
      }
      return operand
    }
    return this.parsePower()
  }

  private parsePower(): Value {
    const base = this.parsePrimary()
    const token = this.peek()
    if (token?.kind === 'op' && token.text === '^') {
      this.consume()
      const exponent = this.parseUnary()
      return this.applyArithmetic('^', base, exponent)
    }
    return base
  }

  private applyArithmetic(op: string, left: Value, right: Value): Value {
    const a = toNumeric(left)
    const b = toNumeric(right)
    if (typeof a !== 'number') return { kind: 'error', code: a.code }
    if (typeof b !== 'number') return { kind: 'error', code: b.code }

    // SHEETFN-2 — Excel's own date arithmetic: adding or subtracting a plain
    // number keeps a date a date, and subtracting one date from another gives
    // a plain count of days. Multiplying a date is not a date.
    const leftDate = left.kind === 'number' ? left.dateKind : undefined
    const rightDate = right.kind === 'number' ? right.dateKind : undefined
    let resultDate: 'date' | 'datetime' | undefined
    if (op === '+') {
      resultDate = leftDate && rightDate ? undefined : (leftDate ?? rightDate)
    } else if (op === '-') {
      resultDate = leftDate && rightDate ? undefined : leftDate
    }

    switch (op) {
      case '+': return dated(numberValue(a + b), resultDate)
      case '-': return dated(numberValue(a - b), resultDate)
      case '*': return numberValue(a * b)
      case '/': return b === 0 ? { kind: 'error', code: '#DIV/0!' } : numberValue(a / b)
      case '^': return numberValue(Math.pow(a, b))
      default: throw new FormulaSyntaxError(`Unknown operator: ${op}`)
    }
  }

  private parsePrimary(): Value {
    const token = this.consume()
    if (token.kind === 'number') return { kind: 'number', value: token.value }
    if (token.kind === 'string') return { kind: 'text', value: token.text }
    if (token.kind === 'lparen') {
      // `parseComparison`, not `parseExpression`: `=(A1>2)*1` has to compare
      // inside the parentheses before multiplying.
      const inner = this.parseComparison()
      const close = this.consume()
      if (close.kind !== 'rparen') throw new FormulaSyntaxError('Expected closing parenthesis')
      return inner
    }
    if (token.kind === 'ref') {
      return this.resolveSingleRef(token.text)
    }
    if (token.kind === 'ident') {
      return this.parseFunctionCall(token.text)
    }
    throw new FormulaSyntaxError(`Unexpected token near ${JSON.stringify(token)}`)
  }

  private resolveSingleRef(refText: string): Value {
    const coord = parseRefToken(refText)
    // SHEETFN-1 — the same `cellValue` a RANGE member goes through, so a blank
    // cell reads as blank either way. It used to resolve to the number 0 here
    // and to blank inside a range, which made `ISBLANK(A1)` impossible to
    // answer correctly and meant the two paths disagreed about the same cell.
    // Arithmetic is unaffected: `toNumeric` maps blank to 0, as Excel does.
    return cellValue(this.lookup(coord.row, coord.col))
  }

  /**
   * Parses `SUM(...)`-style calls.
   *
   * Each argument is collected as an `Arg` — a scalar value or a 2-D block for
   * a range — and handed to `applyFunction` whole. The previous version
   * flattened everything into a list of numbers here, before any function saw
   * it, which made every non-aggregate function impossible to express.
   */
  private parseFunctionCall(name: string): Value {
    const open = this.consume()
    if (open.kind !== 'lparen') throw new FormulaSyntaxError(`Expected "(" after function name ${name}`)

    const args: Arg[] = []

    const collectArg = (): void => {
      const first = this.peek()
      // A range only ever appears as a bare `A1:B3` function argument.
      if (first?.kind === 'ref') {
        const next = this.tokens[this.position + 1]
        if (next?.kind === 'colon') {
          const startRef = this.consume() as Token & { kind: 'ref' }
          this.consume() // colon
          const endToken = this.consume()
          if (endToken.kind !== 'ref') throw new FormulaSyntaxError('Expected cell reference after ":"')
          const start = parseRefToken(startRef.text)
          const end = parseRefToken(endToken.text)
          const rows: Value[][] = []
          for (let r = Math.min(start.row, end.row); r <= Math.max(start.row, end.row); r++) {
            const row: Value[] = []
            for (let c = Math.min(start.col, end.col); c <= Math.max(start.col, end.col); c++) {
              row.push(cellValue(this.lookup(r, c)))
            }
            rows.push(row)
          }
          args.push({ kind: 'range', cells: rows })
          return
        }
      }
      // `parseComparison` (not the narrower `parseExpression`) so an argument
      // can itself be a comparison or use `&` — `IF(A1>2,"big","small")`
      // depends on it.
      args.push({ kind: 'value', value: this.parseComparison() })
    }

    if (this.peek()?.kind !== 'rparen') {
      collectArg()
      while (this.peek()?.kind === 'comma') {
        this.consume()
        collectArg()
      }
    }

    const close = this.consume()
    if (close.kind !== 'rparen') throw new FormulaSyntaxError(`Expected ")" to close ${name}(`)

    return applyFunction(name, args)
  }
}

/**
 * One argument to a function: a scalar, or a rectangular block of cells.
 *
 * A range keeps its SHAPE (rows of values) rather than being flattened,
 * because `VLOOKUP` and `INDEX` address it by row and column.
 */
type Arg =
  | { readonly kind: 'value'; readonly value: Value }
  | { readonly kind: 'range'; readonly cells: ReadonlyArray<ReadonlyArray<Value>> }

/**
 * A cell's display text as a value.
 *
 * Blank stays blank, not 0 — `COUNT` and `AVERAGE` have to tell them apart.
 *
 * SHEETFN-2 — parsing is STRICT: the whole text must be a number. It used to
 * use `Number.parseFloat`, which reads a leading number and ignores the rest,
 * so a cell showing `2026-10-02` was the number 2026 and one showing
 * `15/03/2023` was 15. A column of dates therefore summed to a total of its
 * day-of-month numbers, and `YEAR(A1)` on a date cell answered for 1905.
 *
 * An ISO date is recognised as the date it is, which is what makes
 * `YEAR(A1)`/`A1+7` work on a column of typed dates — the common real case.
 */
function cellValue(text: string): Value {
  const trimmed = text.trim()
  if (trimmed === '') return { kind: 'text', value: '' }

  const serial = serialFromIsoDate(trimmed)
  if (serial !== null) return { kind: 'number', value: serial, dateKind: 'date' }

  // `Number(...)` rejects trailing junk where `parseFloat` would accept it.
  const parsed = Number(trimmed)
  return Number.isNaN(parsed) ? { kind: 'text', value: trimmed } : { kind: 'number', value: parsed }
}

// Delegates to `cellRef.ts`'s own `parseCellRef` (single source of truth for
// A1-style parsing, also used — and tested — independently of the formula
// evaluator) rather than re-implementing the same `$?letters$?digits` regex
// here a second time.
function parseRefToken(refText: string): CellCoord {
  const coord = parseCellRef(refText)
  if (!coord) throw new FormulaSyntaxError(`Malformed reference: ${refText}`)
  return coord
}

// ---------------------------------------------------------------------------
// SHEETFN-1 — the function library.
//
// Before this there were seven functions (SUM, AVERAGE, MIN, MAX, COUNT,
// COUNTA, CONCATENATE) and no comparison operators, which meant no IF: the
// entire conditional half of a spreadsheet was unreachable, and any formula
// using it displayed as literal text.
//
// Each function reads the `Arg`s it actually needs. The aggregates flatten
// ranges and scalars together; the rest address their arguments directly.
// ---------------------------------------------------------------------------

/** Every number in the arguments, from scalars and ranges alike — what the aggregates operate on. */
function flattenNumbers(args: ReadonlyArray<Arg>): number[] | { readonly code: string } {
  const out: number[] = []
  for (const arg of args) {
    if (arg.kind === 'value') {
      if (arg.value.kind === 'error') return { code: arg.value.code }
      if (arg.value.kind === 'number') out.push(arg.value.value)
      else {
        // A scalar text argument that looks numeric counts (`SUM("1",2)` is 3
        // in Excel); one that does not is an error, not a silent zero.
        const trimmed = arg.value.value.trim()
        if (trimmed === '') continue
        const parsed = Number.parseFloat(trimmed)
        if (Number.isNaN(parsed)) return { code: '#VALUE!' }
        out.push(parsed)
      }
      continue
    }
    for (const row of arg.cells) {
      for (const cell of row) {
        // Inside a RANGE, text and blanks are skipped rather than erroring —
        // that is what makes `SUM(A1:A9)` work on a column with a header.
        if (cell.kind === 'error') return { code: cell.code }
        if (cell.kind === 'number') out.push(cell.value)
      }
    }
  }
  return out
}

/** Every cell/scalar in the arguments, flat, for COUNTA and the lookup helpers. */
function flattenValues(args: ReadonlyArray<Arg>): Value[] {
  const out: Value[] = []
  for (const arg of args) {
    if (arg.kind === 'value') out.push(arg.value)
    else for (const row of arg.cells) for (const cell of row) out.push(cell)
  }
  return out
}

function isBlank(value: Value): boolean {
  return value.kind === 'text' && value.value.trim() === ''
}

/** The first argument as a scalar; a 1x1 range counts, as Excel allows. */
function scalarOf(arg: Arg | undefined): Value | null {
  if (!arg) return null
  if (arg.kind === 'value') return arg.value
  const first = arg.cells[0]?.[0]
  return first ?? null
}

function numberOf(arg: Arg | undefined): number | { readonly code: string } {
  const value = scalarOf(arg)
  if (!value) return { code: '#VALUE!' }
  return toNumeric(value)
}

function textOf(arg: Arg | undefined): string | { readonly code: string } {
  const value = scalarOf(arg)
  if (!value) return { code: '#VALUE!' }
  if (value.kind === 'error') return { code: value.code }
  return valueToText(value)
}

/**
 * A COUNTIF/SUMIF criterion: `">10"`, `"<=3"`, `"<>x"`, or a bare value meaning
 * equality.
 *
 * Returns a predicate over a cell value. The comparison itself goes through
 * `compareValues`, so a criterion inherits exactly the same type and
 * case-insensitivity rules as a `>` written out in a formula — rather than a
 * second, subtly different comparison implementation.
 */
function criterionMatcher(criterion: Value): (cell: Value) => boolean {
  let op = '='
  let operand: Value = criterion

  if (criterion.kind === 'text') {
    const match = /^(<=|>=|<>|<|>|=)?\s*(.*)$/.exec(criterion.value.trim())
    if (match) {
      op = match[1] ?? '='
      const rest = match[2]
      const parsed = Number.parseFloat(rest)
      operand =
        rest !== '' && !Number.isNaN(parsed) && String(parsed) === rest.trim()
          ? { kind: 'number', value: parsed }
          : { kind: 'text', value: rest }
    }
  }

  return (cell: Value): boolean => {
    // A blank cell matches no criterion except an explicit "" one, matching
    // Excel — otherwise `COUNTIF(A1:A99,"<10")` would count every empty row.
    if (isBlank(cell) && !(operand.kind === 'text' && operand.value === '')) return false
    const result = compareValues(op, cell, operand)
    return result.kind === 'text' && result.value === 'TRUE'
  }
}

/** The cells of a range argument, or `null` when the argument is not a range. */
function rangeOf(arg: Arg | undefined): ReadonlyArray<ReadonlyArray<Value>> | null {
  if (!arg) return null
  if (arg.kind === 'range') return arg.cells
  return [[arg.value]]
}

const AGGREGATES: ReadonlySet<string> = new Set(['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'PRODUCT', 'MEDIAN'])

function applyAggregate(name: string, numbers: ReadonlyArray<number>): Value {
  switch (name) {
    case 'SUM':
      return numberValue(numbers.reduce((sum, n) => sum + n, 0))
    case 'AVERAGE':
      return numbers.length === 0
        ? { kind: 'error', code: '#DIV/0!' }
        : numberValue(numbers.reduce((s, n) => s + n, 0) / numbers.length)
    case 'MIN':
      // 0 for an empty set, not an error — Excel's own behaviour, and the
      // behaviour this function already had before the library grew.
      return numberValue(numbers.length === 0 ? 0 : Math.min(...numbers))
    case 'MAX':
      return numberValue(numbers.length === 0 ? 0 : Math.max(...numbers))
    case 'COUNT':
      return numberValue(numbers.length)
    case 'PRODUCT':
      return numberValue(numbers.length === 0 ? 0 : numbers.reduce((p, n) => p * n, 1))
    case 'MEDIAN': {
      if (numbers.length === 0) return { kind: 'error', code: '#NUM!' }
      const sorted = [...numbers].sort((a, b) => a - b)
      const mid = Math.floor(sorted.length / 2)
      return numberValue(sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid])
    }
    default:
      throw new FormulaSyntaxError(`Unsupported aggregate: ${name}`)
  }
}

const FUNCTIONS: ReadonlySet<string> = new Set([
  // aggregates
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'COUNTBLANK', 'PRODUCT', 'MEDIAN',
  // logic
  'IF', 'IFERROR', 'AND', 'OR', 'NOT', 'TRUE', 'FALSE',
  // conditional aggregates
  'SUMIF', 'COUNTIF', 'AVERAGEIF', 'SUMIFS', 'COUNTIFS', 'AVERAGEIFS',
  // lookup
  'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP',
  // array-ish
  'SUMPRODUCT',
  // maths
  'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS', 'INT', 'SQRT', 'MOD', 'POWER', 'SIGN',
  // text
  'CONCATENATE', 'CONCAT', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM', 'UPPER', 'LOWER', 'TEXT',
  // type tests
  'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR',
  // dates
  'TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'DAYS', 'EDATE', 'EOMONTH',
])

function applyFunction(name: string, args: ReadonlyArray<Arg>): Value {
  if (!FUNCTIONS.has(name)) {
    throw new FormulaSyntaxError(`Unsupported function: ${name}`)
  }

  // --- logic, first: these are the only functions that must see an ERROR
  // argument rather than propagate it, which is the whole point of IFERROR.
  if (name === 'IFERROR') {
    const value = scalarOf(args[0])
    if (!value) return { kind: 'error', code: '#VALUE!' }
    if (value.kind === 'error') return scalarOf(args[1]) ?? { kind: 'text', value: '' }
    return value
  }
  if (name === 'ISERROR') {
    const value = scalarOf(args[0])
    return boolValue(value?.kind === 'error')
  }

  switch (name) {
    case 'IF': {
      const condition = scalarOf(args[0])
      if (!condition) return { kind: 'error', code: '#VALUE!' }
      const test = truthy(condition)
      if (typeof test !== 'boolean') return { kind: 'error', code: test.code }
      // A missing third argument is FALSE, as in Excel, not an error.
      const branch = test ? scalarOf(args[1]) : scalarOf(args[2])
      return branch ?? boolValue(false)
    }
    case 'AND':
    case 'OR': {
      const values = flattenValues(args).filter(v => !isBlank(v))
      if (values.length === 0) return { kind: 'error', code: '#VALUE!' }
      let result = name === 'AND'
      for (const value of values) {
        const test = truthy(value)
        if (typeof test !== 'boolean') return { kind: 'error', code: test.code }
        result = name === 'AND' ? result && test : result || test
      }
      return boolValue(result)
    }
    case 'NOT': {
      const value = scalarOf(args[0])
      if (!value) return { kind: 'error', code: '#VALUE!' }
      const test = truthy(value)
      return typeof test === 'boolean' ? boolValue(!test) : { kind: 'error', code: test.code }
    }
    case 'TRUE':
      return boolValue(true)
    case 'FALSE':
      return boolValue(false)

    case 'COUNTA':
      return numberValue(flattenValues(args).filter(v => !isBlank(v) && v.kind !== 'error').length)
    case 'COUNTBLANK':
      return numberValue(flattenValues(args).filter(isBlank).length)

    case 'CONCATENATE':
    case 'CONCAT': {
      const parts: string[] = []
      for (const value of flattenValues(args)) {
        if (value.kind === 'error') return value
        parts.push(valueToText(value))
      }
      return { kind: 'text', value: parts.join('') }
    }

    case 'SUMIF':
    case 'COUNTIF':
    case 'AVERAGEIF': {
      const range = rangeOf(args[0])
      const criterion = scalarOf(args[1])
      if (!range || !criterion) return { kind: 'error', code: '#VALUE!' }
      const matches = criterionMatcher(criterion)
      // SUMIF's optional third argument sums a PARALLEL range instead of the
      // tested one — the common "sum column B where column A says x" shape.
      const sumRange = args[2] ? rangeOf(args[2]) : range
      let count = 0
      let total = 0
      let numeric = 0
      for (let r = 0; r < range.length; r += 1) {
        for (let c = 0; c < range[r].length; c += 1) {
          if (!matches(range[r][c])) continue
          count += 1
          const target = sumRange?.[r]?.[c]
          if (target?.kind === 'number') {
            total += target.value
            numeric += 1
          }
        }
      }
      if (name === 'COUNTIF') return numberValue(count)
      if (name === 'SUMIF') return numberValue(total)
      return numeric === 0 ? { kind: 'error', code: '#DIV/0!' } : numberValue(total / numeric)
    }

    case 'SUMIFS':
    case 'COUNTIFS':
    case 'AVERAGEIFS': {
      // The multi-criteria forms, whose argument ORDER differs from the
      // single-criterion ones in a way that is easy to get backwards:
      //   SUMIFS(sumRange, criteriaRange1, criteria1, ...)   -- sum range FIRST
      //   SUMIF(criteriaRange, criteria, [sumRange])         -- sum range LAST
      // COUNTIFS has no sum range at all.
      const counting = name === 'COUNTIFS'
      const sumRange = counting ? null : rangeOf(args[0])
      const pairsFrom = counting ? 0 : 1
      if (!counting && !sumRange) return { kind: 'error', code: '#VALUE!' }

      const conditions: Array<{
        readonly cells: ReadonlyArray<ReadonlyArray<Value>>
        readonly matches: (cell: Value) => boolean
      }> = []
      for (let i = pairsFrom; i + 1 < args.length + 1 && args[i] !== undefined; i += 2) {
        const cells = rangeOf(args[i])
        const criterion = scalarOf(args[i + 1])
        if (!cells || !criterion) return { kind: 'error', code: '#VALUE!' }
        conditions.push({ cells, matches: criterionMatcher(criterion) })
      }
      if (conditions.length === 0) return { kind: 'error', code: '#VALUE!' }

      const rowCount = conditions[0].cells.length
      let count = 0
      let total = 0
      let numeric = 0
      for (let r = 0; r < rowCount; r += 1) {
        const colCount = conditions[0].cells[r]?.length ?? 0
        for (let c = 0; c < colCount; c += 1) {
          // EVERY condition has to match the same position — that is what
          // makes these the "and" forms.
          const all = conditions.every(condition => {
            const cell = condition.cells[r]?.[c]
            return cell !== undefined && condition.matches(cell)
          })
          if (!all) continue
          count += 1
          const target = counting ? undefined : sumRange?.[r]?.[c]
          if (target?.kind === 'number') {
            total += target.value
            numeric += 1
          }
        }
      }
      if (name === 'COUNTIFS') return numberValue(count)
      if (name === 'SUMIFS') return numberValue(total)
      return numeric === 0 ? { kind: 'error', code: '#DIV/0!' } : numberValue(total / numeric)
    }

    case 'XLOOKUP': {
      // XLOOKUP(key, lookupRange, resultRange, [ifNotFound]) — the modern
      // replacement for VLOOKUP, and the one people reach for now. Exact match
      // only, same reasoning as VLOOKUP: its own default IS exact, so unlike
      // VLOOKUP nothing has to be refused here.
      const key = scalarOf(args[0])
      const lookupRange = rangeOf(args[1])
      const resultRange = rangeOf(args[2])
      if (!key || !lookupRange || !resultRange) return { kind: 'error', code: '#VALUE!' }

      const flatten = (cells: ReadonlyArray<ReadonlyArray<Value>>): ReadonlyArray<Value> =>
        cells.length === 1 ? cells[0] : cells.map(row => row[0])
      const keys = flatten(lookupRange)
      const results = flatten(resultRange)

      for (let i = 0; i < keys.length; i += 1) {
        const equal = compareValues('=', keys[i], key)
        if (equal.kind === 'text' && equal.value === 'TRUE') {
          return results[i] ?? { kind: 'error', code: '#REF!' }
        }
      }
      // The fourth argument is what XLOOKUP is liked for: a fallback instead
      // of wrapping the whole call in IFERROR.
      return args[3] ? (scalarOf(args[3]) ?? { kind: 'error', code: '#N/A' }) : { kind: 'error', code: '#N/A' }
    }

    case 'SUMPRODUCT': {
      const ranges = args.map(rangeOf)
      if (ranges.some(r => r === null) || ranges.length === 0) return { kind: 'error', code: '#VALUE!' }
      const first = ranges[0]!
      // Every range must be the same shape; Excel reports #VALUE! otherwise
      // rather than quietly summing the overlap.
      for (const range of ranges) {
        if (range!.length !== first.length) return { kind: 'error', code: '#VALUE!' }
        for (let r = 0; r < first.length; r += 1) {
          if ((range![r]?.length ?? 0) !== (first[r]?.length ?? 0)) return { kind: 'error', code: '#VALUE!' }
        }
      }
      let total = 0
      for (let r = 0; r < first.length; r += 1) {
        for (let c = 0; c < (first[r]?.length ?? 0); c += 1) {
          let product = 1
          for (const range of ranges) {
            const cell = range![r][c]
            if (cell.kind === 'error') return cell
            // Text and blanks count as zero, which is Excel's rule and the
            // reason SUMPRODUCT works on a column with a header.
            product *= cell.kind === 'number' ? cell.value : 0
          }
          total += product
        }
      }
      return numberValue(total)
    }

    case 'TEXT': {
      // Now implementable because SHEETFN-3 brought in a real OOXML number
      // formatter; before that there was nothing to format WITH.
      const value = scalarOf(args[0])
      const formatCode = textOf(args[1])
      if (!value) return { kind: 'error', code: '#VALUE!' }
      if (typeof formatCode !== 'string') return { kind: 'error', code: formatCode.code }
      if (value.kind === 'error') return value
      const numeric = toNumeric(value)
      if (typeof numeric !== 'number') {
        // TEXT on something non-numeric returns it unchanged, as Excel does.
        return { kind: 'text', value: valueToText(value) }
      }
      const formatted = applyNumberFormat(numeric, formatCode)
      return { kind: 'text', value: formatted ?? formatNumericResult(numeric) }
    }

    case 'VLOOKUP':
    case 'HLOOKUP': {
      const key = scalarOf(args[0])
      const table = rangeOf(args[1])
      const indexArg = numberOf(args[2])
      if (!key || !table) return { kind: 'error', code: '#VALUE!' }
      if (typeof indexArg !== 'number') return { kind: 'error', code: indexArg.code }
      const index = Math.trunc(indexArg)
      if (index < 1) return { kind: 'error', code: '#VALUE!' }

      // Only EXACT match is supported. Excel's default is the approximate
      // (sorted-range) match, which silently returns a wrong row on unsorted
      // data — implementing the default wrongly would be worse than being
      // explicit, so a request for approximate matching is refused rather
      // than approximated. See docs/KNOWN_LIMITATIONS.md.
      if (args[3]) {
        const approx = truthy(scalarOf(args[3]) ?? boolValue(false))
        if (approx === true) throw new FormulaSyntaxError('VLOOKUP approximate match is not supported')
      }

      const vertical = name === 'VLOOKUP'
      const lineCount = vertical ? table.length : (table[0]?.length ?? 0)
      for (let i = 0; i < lineCount; i += 1) {
        const probe = vertical ? table[i]?.[0] : table[0]?.[i]
        if (!probe) continue
        const equal = compareValues('=', probe, key)
        if (equal.kind === 'text' && equal.value === 'TRUE') {
          const found = vertical ? table[i]?.[index - 1] : table[index - 1]?.[i]
          return found ?? { kind: 'error', code: '#REF!' }
        }
      }
      return { kind: 'error', code: '#N/A' }
    }

    case 'INDEX': {
      const table = rangeOf(args[0])
      const rowArg = numberOf(args[1])
      if (!table) return { kind: 'error', code: '#VALUE!' }
      if (typeof rowArg !== 'number') return { kind: 'error', code: rowArg.code }
      const rowIndex = Math.trunc(rowArg)
      // A single-row or single-column range can be addressed with one index,
      // which is how INDEX/MATCH is almost always written.
      if (args[2] === undefined) {
        if (table.length === 1) {
          return table[0][rowIndex - 1] ?? { kind: 'error', code: '#REF!' }
        }
        return table[rowIndex - 1]?.[0] ?? { kind: 'error', code: '#REF!' }
      }
      const colArg = numberOf(args[2])
      if (typeof colArg !== 'number') return { kind: 'error', code: colArg.code }
      return table[rowIndex - 1]?.[Math.trunc(colArg) - 1] ?? { kind: 'error', code: '#REF!' }
    }

    case 'MATCH': {
      const key = scalarOf(args[0])
      const table = rangeOf(args[1])
      if (!key || !table) return { kind: 'error', code: '#VALUE!' }
      if (args[2]) {
        const type = numberOf(args[2])
        // Same reasoning as VLOOKUP: only exact (type 0) is implemented.
        if (typeof type === 'number' && type !== 0) {
          throw new FormulaSyntaxError('MATCH only supports exact match (match_type 0)')
        }
      }
      const flat = table.length === 1 ? table[0] : table.map(row => row[0])
      for (let i = 0; i < flat.length; i += 1) {
        const equal = compareValues('=', flat[i], key)
        if (equal.kind === 'text' && equal.value === 'TRUE') return numberValue(i + 1)
      }
      return { kind: 'error', code: '#N/A' }
    }

    case 'LEN': {
      const text = textOf(args[0])
      return typeof text === 'string' ? numberValue(text.length) : { kind: 'error', code: text.code }
    }
    case 'TRIM':
    case 'UPPER':
    case 'LOWER': {
      const text = textOf(args[0])
      if (typeof text !== 'string') return { kind: 'error', code: text.code }
      const out = name === 'TRIM' ? text.trim().replace(/\s+/g, ' ') : name === 'UPPER' ? text.toUpperCase() : text.toLowerCase()
      return { kind: 'text', value: out }
    }
    case 'LEFT':
    case 'RIGHT': {
      const text = textOf(args[0])
      if (typeof text !== 'string') return { kind: 'error', code: text.code }
      const countArg = args[1] === undefined ? 1 : numberOf(args[1])
      if (typeof countArg !== 'number') return { kind: 'error', code: countArg.code }
      const count = Math.trunc(countArg)
      if (count < 0) return { kind: 'error', code: '#VALUE!' }
      return { kind: 'text', value: name === 'LEFT' ? text.slice(0, count) : count === 0 ? '' : text.slice(-count) }
    }
    case 'MID': {
      const text = textOf(args[0])
      if (typeof text !== 'string') return { kind: 'error', code: text.code }
      const startArg = numberOf(args[1])
      const countArg = numberOf(args[2])
      if (typeof startArg !== 'number') return { kind: 'error', code: startArg.code }
      if (typeof countArg !== 'number') return { kind: 'error', code: countArg.code }
      const start = Math.trunc(startArg)
      if (start < 1) return { kind: 'error', code: '#VALUE!' }
      return { kind: 'text', value: text.slice(start - 1, start - 1 + Math.max(0, Math.trunc(countArg))) }
    }

    case 'TODAY':
      return dated(numberValue(todaySerial()), 'date')
    case 'NOW':
      return dated(numberValue(nowSerial()), 'datetime')

    case 'DATE': {
      const parts = [numberOf(args[0]), numberOf(args[1]), numberOf(args[2])]
      for (const part of parts) {
        if (typeof part !== 'number') return { kind: 'error', code: part.code }
      }
      const [year, month, day] = parts as number[]
      const serial = serialFromDate({
        year: Math.trunc(year),
        month: Math.trunc(month),
        day: Math.trunc(day),
      })
      // `#NUM!` is what Excel reports for a date it cannot build, and
      // `serialFromDate` rejects an impossible day rather than rolling it over
      // into a different one.
      return serial === null ? { kind: 'error', code: '#NUM!' } : dated(numberValue(serial), 'date')
    }

    case 'YEAR':
    case 'MONTH':
    case 'DAY':
    case 'WEEKDAY': {
      const value = scalarOf(args[0])
      if (!value) return { kind: 'error', code: '#VALUE!' }
      const serial = toSerial(value)
      if (typeof serial !== 'number') return { kind: 'error', code: serial.code }
      const date = dateFromSerialChecked(serial)
      if (!date) return { kind: 'error', code: '#NUM!' }
      if (name === 'YEAR') return numberValue(date.year)
      if (name === 'MONTH') return numberValue(date.month)
      if (name === 'DAY') return numberValue(date.day)
      // WEEKDAY's default numbering is 1 = Sunday. Derived from the calendar,
      // not from `serial % 7`: Excel's two epochs (see `excelDate.ts`) make the
      // modulo answer off by one for part of the range. 1900-01-01 was a
      // Monday, not a Sunday — which is exactly the kind of remembered fact
      // that made the first version of this wrong.
      const weekday = weekdayFromSerial(serial)
      return weekday === null ? { kind: 'error', code: '#NUM!' } : numberValue(weekday + 1)
    }

    case 'DAYS': {
      // DAYS(end, start) -- end first, matching Excel, which is the opposite
      // order from how it reads aloud.
      const end = scalarOf(args[0])
      const start = scalarOf(args[1])
      if (!end || !start) return { kind: 'error', code: '#VALUE!' }
      const endSerial = toSerial(end)
      const startSerial = toSerial(start)
      if (typeof endSerial !== 'number') return { kind: 'error', code: endSerial.code }
      if (typeof startSerial !== 'number') return { kind: 'error', code: startSerial.code }
      return numberValue(Math.floor(endSerial) - Math.floor(startSerial))
    }

    case 'EDATE':
    case 'EOMONTH': {
      const value = scalarOf(args[0])
      const monthsArg = numberOf(args[1])
      if (!value) return { kind: 'error', code: '#VALUE!' }
      if (typeof monthsArg !== 'number') return { kind: 'error', code: monthsArg.code }
      const serial = toSerial(value)
      if (typeof serial !== 'number') return { kind: 'error', code: serial.code }
      const date = dateFromSerialChecked(serial)
      if (!date) return { kind: 'error', code: '#NUM!' }

      const months = Math.trunc(monthsArg)
      const targetMonthIndex = date.month - 1 + months
      const targetYear = date.year + Math.floor(targetMonthIndex / 12)
      const targetMonth = ((targetMonthIndex % 12) + 12) % 12 + 1
      const lastDay = daysInMonth(targetYear, targetMonth)
      // EDATE clamps the day to the target month's length -- Excel's own
      // behaviour, and the reason EDATE("2026-01-31",1) is the 28th of
      // February rather than an error or the 3rd of March.
      const day = name === 'EOMONTH' ? lastDay : Math.min(date.day, lastDay)
      const result = serialFromDate({ year: targetYear, month: targetMonth, day })
      return result === null ? { kind: 'error', code: '#NUM!' } : dated(numberValue(result), 'date')
    }

    case 'ISBLANK': {
      const value = scalarOf(args[0])
      return boolValue(value ? isBlank(value) : false)
    }
    case 'ISNUMBER': {
      const value = scalarOf(args[0])
      return boolValue(value?.kind === 'number')
    }
    case 'ISTEXT': {
      const value = scalarOf(args[0])
      return boolValue(value?.kind === 'text' && value.value.trim() !== '')
    }

    default:
      break
  }

  // --- maths and the aggregates, all of which propagate an error operand.
  if (AGGREGATES.has(name)) {
    const numbers = flattenNumbers(args)
    if (!Array.isArray(numbers)) return { kind: 'error', code: numbers.code }
    return applyAggregate(name, numbers)
  }

  const first = numberOf(args[0])
  if (typeof first !== 'number') return { kind: 'error', code: first.code }
  switch (name) {
    case 'ABS':
      return numberValue(Math.abs(first))
    case 'INT':
      return numberValue(Math.floor(first))
    case 'SIGN':
      return numberValue(Math.sign(first))
    case 'SQRT':
      return first < 0 ? { kind: 'error', code: '#NUM!' } : numberValue(Math.sqrt(first))
    case 'ROUND':
    case 'ROUNDUP':
    case 'ROUNDDOWN': {
      const digitsArg = args[1] === undefined ? 0 : numberOf(args[1])
      if (typeof digitsArg !== 'number') return { kind: 'error', code: digitsArg.code }
      const factor = Math.pow(10, Math.trunc(digitsArg))
      const scaled = first * factor
      // ROUNDUP/ROUNDDOWN are away-from-zero and toward-zero, NOT ceil/floor —
      // they differ for negative numbers, which is where a naive
      // implementation is wrong.
      const rounded =
        name === 'ROUND'
          ? Math.sign(scaled) * Math.round(Math.abs(scaled))
          : name === 'ROUNDUP'
            ? Math.sign(scaled) * Math.ceil(Math.abs(scaled))
            : Math.sign(scaled) * Math.floor(Math.abs(scaled))
      return numberValue(rounded / factor)
    }
    case 'MOD': {
      const divisor = numberOf(args[1])
      if (typeof divisor !== 'number') return { kind: 'error', code: divisor.code }
      if (divisor === 0) return { kind: 'error', code: '#DIV/0!' }
      // Excel's MOD takes the sign of the DIVISOR; JavaScript's % takes the
      // sign of the dividend, so `MOD(-3,2)` is 1 in Excel and -1 in JS.
      return numberValue(first - divisor * Math.floor(first / divisor))
    }
    case 'POWER': {
      const exponent = numberOf(args[1])
      if (typeof exponent !== 'number') return { kind: 'error', code: exponent.code }
      return numberValue(Math.pow(first, exponent))
    }
    default:
      throw new FormulaSyntaxError(`Unsupported function: ${name}`)
  }
}

/** A serial's calendar date, or `null` — re-exported through a name that reads at the call sites above. */
function dateFromSerialChecked(serial: number): { year: number; month: number; day: number } | null {
  const iso = isoDateFromSerial(serial)
  if (!iso) return null
  const [year, month, day] = iso.split('-').map(Number)
  return { year, month, day }
}

/** Days in a month, with the real leap-year rule (not Excel's 1900 quirk, which only affects February 1900). */
function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate()
}

/** Strips a fixed-point rounding artifact (e.g. `0.30000000000000004`) before stringifying a computed number. */
function formatNumericResult(value: number): string {
  const rounded = Math.round(value * 1e10) / 1e10
  return String(rounded)
}

/**
 * Evaluates a formula (without its leading `=`) against `lookup`, which
 * resolves a zero-indexed (row, col) to that cell's current display text.
 * Returns `{ok:false}` when the formula can't be evaluated at all (syntax
 * error, unsupported function) — callers should fall back to showing the
 * literal `=<formula>` text. A runtime error within a valid formula (e.g.
 * division by zero) resolves `{ok:true}` with the spreadsheet error string
 * as its text, matching real spreadsheet behavior.
 */
export function evaluateFormula(formula: string, lookup: CellLookup): FormulaResult {
  try {
    const tokens = tokenize(formula)
    if (tokens.length === 0) return { ok: false }
    const value = new FormulaParser(tokens, lookup).parseTopLevel()
    if (value.kind === 'error') return { ok: true, text: value.code }
    if (value.kind === 'number') {
      // SHEETFN-2 — a value the evaluator knows is a date is rendered as one,
      // rather than as the serial number it is underneath. See `Value`'s own
      // note on why the flag exists rather than relying on the cell's format.
      if (value.dateKind !== undefined) {
        const rendered =
          value.dateKind === 'datetime' ? isoDateTimeFromSerial(value.value) : isoDateFromSerial(value.value)
        if (rendered !== null) {
          return { ok: true, text: rendered, value: value.value, dateKind: value.dateKind }
        }
      }
      return { ok: true, text: formatNumericResult(value.value), value: value.value }
    }
    return { ok: true, text: value.value }
  } catch {
    return { ok: false }
  }
}
