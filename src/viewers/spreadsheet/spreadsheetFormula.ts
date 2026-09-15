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
 * cell references, `A1:B3` ranges, and the handful of aggregate functions
 * listed in `FUNCTIONS` below. Anything outside that — an unknown function
 * name, a malformed range, a bare syntax error — resolves to `null`, and
 * callers (see `spreadsheetDocument.ts`) fall back to displaying the literal
 * formula text (`=<formula>`) instead of a fabricated number. A *runtime*
 * error within an otherwise-valid formula (division by zero, a text operand
 * in an arithmetic expression) still resolves to a real spreadsheet error
 * string (`#DIV/0!`, `#VALUE!`) — that IS a value a real spreadsheet would
 * show, unlike a fallback.
 */
import { parseCellRef, type CellCoord } from './cellRef'

export type CellLookup = (row: number, col: number) => string

export type FormulaResult =
  | { readonly ok: true; readonly text: string }
  | { readonly ok: false }

type Token =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'ref'; readonly text: string }
  | { readonly kind: 'ident'; readonly text: string }
  | { readonly kind: 'op'; readonly text: string }
  | { readonly kind: 'lparen' }
  | { readonly kind: 'rparen' }
  | { readonly kind: 'comma' }
  | { readonly kind: 'colon' }

const TOKEN_PATTERN = /\s*(?:(\d+(?:\.\d+)?)|(\$?[A-Za-z]+\$?\d+)|([A-Za-z_][A-Za-z0-9_]*)|(<=|>=|<>|[-+*/^()=<>,:])|(\S))/g

class FormulaSyntaxError extends Error {}

function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  TOKEN_PATTERN.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = TOKEN_PATTERN.exec(formula)) !== null) {
    const [, number, ref, ident, op, junk] = match
    if (number !== undefined) {
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
type Value = { readonly kind: 'number'; readonly value: number } | { readonly kind: 'text'; readonly value: string } | { readonly kind: 'error'; readonly code: string }

function numberValue(n: number): Value {
  if (Number.isNaN(n)) return { kind: 'error', code: '#VALUE!' }
  if (!Number.isFinite(n)) return { kind: 'error', code: '#DIV/0!' }
  return { kind: 'number', value: n }
}

function toNumeric(value: Value): number | { readonly code: string } {
  if (value.kind === 'error') return { code: value.code }
  if (value.kind === 'number') return value.value
  const parsed = Number.parseFloat(value.value)
  return Number.isNaN(parsed) ? { code: '#VALUE!' } : parsed
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
    const result = this.parseExpression()
    if (this.position !== this.tokens.length) {
      throw new FormulaSyntaxError('Unexpected trailing tokens')
    }
    return result
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
    switch (op) {
      case '+': return numberValue(a + b)
      case '-': return numberValue(a - b)
      case '*': return numberValue(a * b)
      case '/': return b === 0 ? { kind: 'error', code: '#DIV/0!' } : numberValue(a / b)
      case '^': return numberValue(Math.pow(a, b))
      default: throw new FormulaSyntaxError(`Unknown operator: ${op}`)
    }
  }

  private parsePrimary(): Value {
    const token = this.consume()
    if (token.kind === 'number') return { kind: 'number', value: token.value }
    if (token.kind === 'lparen') {
      const inner = this.parseExpression()
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
    const text = this.lookup(coord.row, coord.col)
    if (text.trim() === '') return numberValue(0)
    const parsed = Number.parseFloat(text)
    return Number.isNaN(parsed) ? { kind: 'text', value: text } : { kind: 'number', value: parsed }
  }

  /** Parses `SUM(...)`-style calls, including range arguments (only meaningful inside a function call). */
  private parseFunctionCall(name: string): Value {
    const open = this.consume()
    if (open.kind !== 'lparen') throw new FormulaSyntaxError(`Expected "(" after function name ${name}`)

    const numbers: number[] = []
    let nonEmptyCount = 0

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
          for (let r = Math.min(start.row, end.row); r <= Math.max(start.row, end.row); r++) {
            for (let c = Math.min(start.col, end.col); c <= Math.max(start.col, end.col); c++) {
              const text = this.lookup(r, c)
              if (text.trim() === '') continue
              nonEmptyCount += 1
              const parsed = Number.parseFloat(text)
              if (!Number.isNaN(parsed)) numbers.push(parsed)
            }
          }
          return
        }
      }
      const value = this.parseExpression()
      if (value.kind === 'error') throw new FormulaSyntaxError(value.code)
      if (value.kind === 'number') {
        numbers.push(value.value)
        nonEmptyCount += 1
      } else if (value.value.trim() !== '') {
        nonEmptyCount += 1
      }
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

    return applyFunction(name, numbers, nonEmptyCount)
  }
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

const FUNCTIONS = new Set(['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA'])

function applyFunction(name: string, numbers: ReadonlyArray<number>, nonEmptyCount: number): Value {
  if (!FUNCTIONS.has(name)) {
    throw new FormulaSyntaxError(`Unsupported function: ${name}`)
  }
  switch (name) {
    case 'SUM':
      return numberValue(numbers.reduce((sum, n) => sum + n, 0))
    case 'AVERAGE':
      return numbers.length === 0 ? { kind: 'error', code: '#DIV/0!' } : numberValue(numbers.reduce((s, n) => s + n, 0) / numbers.length)
    case 'MIN':
      return numberValue(numbers.length === 0 ? 0 : Math.min(...numbers))
    case 'MAX':
      return numberValue(numbers.length === 0 ? 0 : Math.max(...numbers))
    case 'COUNT':
      return numberValue(numbers.length)
    case 'COUNTA':
      return numberValue(nonEmptyCount)
    default:
      throw new FormulaSyntaxError(`Unsupported function: ${name}`)
  }
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
    if (value.kind === 'number') return { ok: true, text: formatNumericResult(value.value) }
    return { ok: true, text: value.value }
  } catch {
    return { ok: false }
  }
}
