/**
 * SHEETFN-4 — the multi-criteria, XLOOKUP, SUMPRODUCT and TEXT functions.
 *
 * These are the ones a real workbook reaches for after `IF` and `VLOOKUP`, and
 * three of them have a trap worth pinning:
 *   - the `*IFS` forms take their arguments in the OPPOSITE order to the
 *     single-criterion ones (sum range first, not last);
 *   - `SUMPRODUCT` must reject mismatched shapes rather than summing the
 *     overlap, and treats text as zero so it works on a column with a header;
 *   - `XLOOKUP`'s fourth argument is a not-found fallback, which is most of why
 *     people prefer it to `VLOOKUP`.
 */
import { describe, expect, it } from 'vitest'

import { evaluateFormula } from '../spreadsheetFormula'

function gridLookup(rows: ReadonlyArray<ReadonlyArray<string>>) {
  return (row: number, col: number): string => rows[row]?.[col] ?? ''
}

function evaluate(formula: string, rows: ReadonlyArray<ReadonlyArray<string>> = []): string | null {
  const result = evaluateFormula(formula, gridLookup(rows))
  return result.ok ? result.text : null
}

/** Region, product, units — the shape a multi-criteria formula is written for. */
const SALES = [
  ['North', 'Paper', '10'],
  ['South', 'Paper', '25'],
  ['North', 'Ink', '5'],
  ['North', 'Paper', '7'],
]

describe('multi-criteria aggregates (SHEETFN-4)', () => {
  it('COUNTIFS counts rows matching every condition', () => {
    expect(evaluate('COUNTIFS(A1:A4,"North",B1:B4,"Paper")', SALES)).toBe('2')
    expect(evaluate('COUNTIFS(A1:A4,"North",B1:B4,"Ink")', SALES)).toBe('1')
    expect(evaluate('COUNTIFS(A1:A4,"East",B1:B4,"Paper")', SALES)).toBe('0')
  })

  it('SUMIFS takes the sum range FIRST, unlike SUMIF', () => {
    // The order trap: SUMIF(criteriaRange, criteria, sumRange) puts the sum
    // range last, SUMIFS(sumRange, criteriaRange, criteria) puts it first.
    expect(evaluate('SUMIFS(C1:C4,A1:A4,"North",B1:B4,"Paper")', SALES)).toBe('17')
    expect(evaluate('SUMIFS(C1:C4,A1:A4,"South")', SALES)).toBe('25')
  })

  it('AVERAGEIFS averages the matching rows, and reports #DIV/0! for none', () => {
    expect(evaluate('AVERAGEIFS(C1:C4,A1:A4,"North",B1:B4,"Paper")', SALES)).toBe('8.5')
    expect(evaluate('AVERAGEIFS(C1:C4,A1:A4,"East")', SALES)).toBe('#DIV/0!')
  })

  it('combines a comparison criterion with a text one', () => {
    expect(evaluate('COUNTIFS(A1:A4,"North",C1:C4,">6")', SALES)).toBe('2')
    expect(evaluate('SUMIFS(C1:C4,A1:A4,"North",C1:C4,">6")', SALES)).toBe('17')
  })
})

describe('XLOOKUP (SHEETFN-4)', () => {
  const table = [
    ['Paper', '10'],
    ['Ink', '25'],
    ['Toner', '80'],
  ]

  it('finds a value by key', () => {
    expect(evaluate('XLOOKUP("Ink",A1:A3,B1:B3)', table)).toBe('25')
  })

  it('returns the fourth argument when the key is missing', () => {
    // The reason people prefer it to VLOOKUP: no IFERROR wrapper needed.
    expect(evaluate('XLOOKUP("Pens",A1:A3,B1:B3,"none")', table)).toBe('none')
  })

  it('reports #N/A with no fallback given', () => {
    expect(evaluate('XLOOKUP("Pens",A1:A3,B1:B3)', table)).toBe('#N/A')
  })

  it('looks left as easily as right, which VLOOKUP cannot', () => {
    // The lookup and result ranges are independent, so the key column can sit
    // to the RIGHT of the answer.
    //
    // The key is written unquoted, because the cell holds the NUMBER 25 and a
    // number is never equal to text — the type rule the comparison operators
    // deliberately implement. `XLOOKUP("25",...)` correctly finds nothing.
    expect(evaluate('XLOOKUP(25,B1:B3,A1:A3)', table)).toBe('Ink')
    expect(evaluate('XLOOKUP("25",B1:B3,A1:A3,"no text match")', table)).toBe('no text match')
  })

  it('works across a row as well as down a column', () => {
    const row = [['Paper', 'Ink'], ['10', '25']]
    expect(evaluate('XLOOKUP("Ink",A1:B1,A2:B2)', row)).toBe('25')
  })
})

describe('SUMPRODUCT (SHEETFN-4)', () => {
  const grid = [
    ['2', '3'],
    ['4', '5'],
  ]

  it('sums the pairwise products', () => {
    // 2*3 + 4*5 = 26
    expect(evaluate('SUMPRODUCT(A1:A2,B1:B2)', grid)).toBe('26')
  })

  it('sums a single range, like SUM', () => {
    expect(evaluate('SUMPRODUCT(A1:A2)', grid)).toBe('6')
  })

  it('rejects mismatched shapes rather than summing the overlap', () => {
    const ragged = [['1', '2'], ['3', '4'], ['5', '6']]
    expect(evaluate('SUMPRODUCT(A1:A3,B1:B2)', ragged)).toBe('#VALUE!')
  })

  it('treats text as zero, so it works on a column with a header', () => {
    const withHeader = [['Qty', 'Price'], ['2', '3'], ['4', '5']]
    expect(evaluate('SUMPRODUCT(A1:A3,B1:B3)', withHeader)).toBe('26')
  })

  it('handles a weighted total, the thing it exists for', () => {
    const order = [['2', '1.50'], ['3', '2.00'], ['1', '10.00']]
    expect(evaluate('SUMPRODUCT(A1:A3,B1:B3)', order)).toBe('19')
  })
})

describe('TEXT (SHEETFN-4)', () => {
  it('formats a number with an OOXML format code', () => {
    expect(evaluate('TEXT(1234.5,"#,##0.00")')).toBe('1,234.50')
    expect(evaluate('TEXT(0.125,"0%")')).toBe('13%')
  })

  it('formats a date', () => {
    expect(evaluate('TEXT(DATE(2026,10,2),"dd/mm/yyyy")')).toBe('02/10/2026')
  })

  it('returns a non-numeric value unchanged', () => {
    expect(evaluate('TEXT("already text","0.00")')).toBe('already text')
  })

  it('produces TEXT, so it concatenates rather than adding', () => {
    expect(evaluate('TEXT(5,"0.00")&" units"')).toBe('5.00 units')
  })

  it('falls back to a plain number for a format code it cannot apply', () => {
    expect(evaluate('TEXT(5,"[[[nonsense")')).toBe('5')
  })
})
