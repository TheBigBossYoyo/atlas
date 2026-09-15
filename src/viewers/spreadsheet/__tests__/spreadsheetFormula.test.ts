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
    expect(evaluateFormula('1 & 2', lookup)).toEqual({ ok: false })
  })
})
