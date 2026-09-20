import { describe, expect, it } from 'vitest'

import { evaluateFormula, type CellLookup } from '../spreadsheetFormula'

function gridLookup(rows: ReadonlyArray<ReadonlyArray<string>>): CellLookup {
  return (row, col) => rows[row]?.[col] ?? ''
}

describe('evaluateFormula — arithmetic', () => {
  it('evaluates a literal arithmetic expression', () => {
    expect(evaluateFormula('1+2*3', gridLookup([]))).toEqual({ ok: true, text: '7' })
  })

  it('respects parentheses', () => {
    expect(evaluateFormula('(1+2)*3', gridLookup([]))).toEqual({ ok: true, text: '9' })
  })

  it('handles unary minus', () => {
    expect(evaluateFormula('-5+2', gridLookup([]))).toEqual({ ok: true, text: '-3' })
  })

  it('handles exponentiation', () => {
    expect(evaluateFormula('2^10', gridLookup([]))).toEqual({ ok: true, text: '1024' })
  })

  it('strips floating-point rounding noise', () => {
    expect(evaluateFormula('0.1+0.2', gridLookup([]))).toEqual({ ok: true, text: '0.3' })
  })

  it('returns #DIV/0! for division by zero', () => {
    expect(evaluateFormula('5/0', gridLookup([]))).toEqual({ ok: true, text: '#DIV/0!' })
  })
})

describe('evaluateFormula — cell references', () => {
  const lookup = gridLookup([
    ['10', '20', 'hello'],
    ['30', '', '5'],
  ])

  it('resolves a single cell reference', () => {
    expect(evaluateFormula('A1+B1', lookup)).toEqual({ ok: true, text: '30' })
  })

  it('resolves an absolute-anchored reference the same as a relative one', () => {
    expect(evaluateFormula('$A$1+$B$1', lookup)).toEqual({ ok: true, text: '30' })
  })

  it('treats a blank referenced cell as zero', () => {
    expect(evaluateFormula('A2+B2', lookup)).toEqual({ ok: true, text: '30' })
  })

  it('returns #VALUE! when a referenced cell holds non-numeric text used arithmetically', () => {
    expect(evaluateFormula('C1+1', lookup)).toEqual({ ok: true, text: '#VALUE!' })
  })
})

describe('evaluateFormula — functions', () => {
  const lookup = gridLookup([
    ['1', '2', '3'],
    ['4', '5', ''],
    ['x', '', '10'],
  ])

  it('SUM over a range', () => {
    expect(evaluateFormula('SUM(A1:C2)', lookup)).toEqual({ ok: true, text: '15' })
  })

  it('AVERAGE over a range ignores blanks', () => {
    expect(evaluateFormula('AVERAGE(A1:A2)', lookup)).toEqual({ ok: true, text: '2.5' })
  })

  it('MIN and MAX over a range', () => {
    expect(evaluateFormula('MIN(A1:C1)', lookup)).toEqual({ ok: true, text: '1' })
    expect(evaluateFormula('MAX(A1:C1)', lookup)).toEqual({ ok: true, text: '3' })
  })

  it('COUNT counts only numeric cells, COUNTA counts every non-empty cell', () => {
    expect(evaluateFormula('COUNT(A1:C3)', lookup)).toEqual({ ok: true, text: '6' })
    expect(evaluateFormula('COUNTA(A1:C3)', lookup)).toEqual({ ok: true, text: '7' })
  })

  it('is case-insensitive on function names', () => {
    expect(evaluateFormula('sum(A1:A2)', lookup)).toEqual({ ok: true, text: '5' })
  })

  it('combines a function with arithmetic', () => {
    expect(evaluateFormula('SUM(A1:A2)+10', lookup)).toEqual({ ok: true, text: '15' })
  })
})

describe('evaluateFormula — unsupported / malformed input falls back to null', () => {
  const lookup = gridLookup([['1']])

  it('rejects an unknown function', () => {
    expect(evaluateFormula('VLOOKUP(A1, A1:A1, 1)', lookup)).toEqual({ ok: false })
  })

  it('rejects unbalanced parentheses', () => {
    expect(evaluateFormula('(1+2', lookup)).toEqual({ ok: false })
  })

  it('rejects trailing garbage', () => {
    expect(evaluateFormula('1+2 3', lookup)).toEqual({ ok: false })
  })

  it('rejects an empty formula', () => {
    expect(evaluateFormula('', lookup)).toEqual({ ok: false })
  })

  it('rejects an unrecognized character', () => {
    // `&` used to land here too (see the SHEET-6 describe block below for
    // why it no longer does) — `@` still isn't part of this evaluator's
    // grammar at all, so it's the replacement regression case for "a bare
    // unrecognized character is still a syntax error".
    expect(evaluateFormula('1 @ 2', lookup)).toEqual({ ok: false })
  })

  it('rejects an unterminated string literal', () => {
    expect(evaluateFormula('"unterminated', lookup)).toEqual({ ok: false })
  })
})

// SHEET-6 — text-returning formulas were never evaluated: the tokenizer had
// no string-literal rule at all (a bare `"` was "unrecognized character")
// and no `&` operator, so both `=A1&"!"` and `=CONCATENATE(...)` fell all
// the way through to the "unsupported" case above and displayed as literal
// formula text instead of their evaluated value. These fixed that gap.
describe('evaluateFormula — string literals and text concatenation (SHEET-6)', () => {
  it('evaluates a bare string literal', () => {
    expect(evaluateFormula('"hello"', gridLookup([]))).toEqual({ ok: true, text: 'hello' })
  })

  it('unescapes a doubled quote inside a string literal', () => {
    expect(evaluateFormula('"say ""hi"""', gridLookup([]))).toEqual({ ok: true, text: 'say "hi"' })
  })

  it('concatenates a cell reference with a string literal (`=A1&"!"`)', () => {
    const lookup = gridLookup([['hello']])
    expect(evaluateFormula('A1&"!"', lookup)).toEqual({ ok: true, text: 'hello!' })
  })

  it('formats a numeric operand the same way a numeric result displays when concatenating', () => {
    const lookup = gridLookup([['5']])
    expect(evaluateFormula('A1&"!"', lookup)).toEqual({ ok: true, text: '5!' })
  })

  it('& binds looser than arithmetic, matching real spreadsheet precedence', () => {
    expect(evaluateFormula('"1"&1+1', gridLookup([]))).toEqual({ ok: true, text: '12' })
  })

  it('chains multiple & operators left to right', () => {
    const lookup = gridLookup([['a', 'b', 'c']])
    expect(evaluateFormula('A1&B1&C1', lookup)).toEqual({ ok: true, text: 'abc' })
  })

  it('CONCATENATE joins string-literal and cell-reference arguments', () => {
    const lookup = gridLookup([['World']])
    expect(evaluateFormula('CONCATENATE("Hello, ",A1,"!")', lookup)).toEqual({ ok: true, text: 'Hello, World!' })
  })

  it('CONCATENATE is case-insensitive, like the other function names', () => {
    const lookup = gridLookup([['a', 'b']])
    expect(evaluateFormula('concatenate(A1,B1)', lookup)).toEqual({ ok: true, text: 'ab' })
  })

  it('a text formula can depend on another text formula (dependency chaining)', () => {
    // `recalculateSheet` resolves the DEPENDENCY ORDER; this evaluator only
    // needs to prove it can consume another formula CELL'S already-evaluated
    // display text through `lookup`, exactly like it always has for numbers.
    const lookup = gridLookup([['hello!', '']]) // B1 depends on A1's own evaluated text
    expect(evaluateFormula('A1&" world"', lookup)).toEqual({ ok: true, text: 'hello! world' })
  })

  it('a text formula can reference a numeric cell', () => {
    const lookup = gridLookup([['5', '']])
    expect(evaluateFormula('"Total: "&A1', lookup)).toEqual({ ok: true, text: 'Total: 5' })
  })

  it('propagates an error operand through & instead of stringifying it', () => {
    expect(evaluateFormula('1/0&"!"', gridLookup([]))).toEqual({ ok: true, text: '#DIV/0!' })
  })
})
