/**
 * SHEETFN-1 — comparisons, IF, and the rest of the function library.
 *
 * Before this, the evaluator had seven aggregate/text functions and — although
 * the tokenizer recognised `< > <= >= <>` — nothing that PARSED a comparison.
 * `=1>0` was a syntax error, so `IF` was unreachable and with it every
 * conditional, lookup and criterion-based formula. A spreadsheet without `IF`
 * is a table.
 *
 * The cases below lean on the places a spreadsheet deliberately disagrees with
 * JavaScript, because those are what a from-scratch evaluator gets wrong:
 * MOD's sign, ROUNDUP on negatives, case-insensitive text equality, numbers
 * sorting before text, and blanks not matching a `<` criterion.
 */
import { describe, expect, it } from 'vitest'

import { evaluateFormula } from '../spreadsheetFormula'

/** A lookup over a literal grid of display text, as the viewer supplies. */
function gridLookup(rows: ReadonlyArray<ReadonlyArray<string>>) {
  return (row: number, col: number): string => rows[row]?.[col] ?? ''
}

const EMPTY = gridLookup([])

/** Evaluates against `rows` and returns the text, or `null` when the formula could not be evaluated. */
function evaluate(formula: string, rows: ReadonlyArray<ReadonlyArray<string>> = []): string | null {
  const result = evaluateFormula(formula, gridLookup(rows))
  return result.ok ? result.text : null
}

describe('comparison operators (SHEETFN-1)', () => {
  it('compares numbers', () => {
    expect(evaluate('1>0')).toBe('TRUE')
    expect(evaluate('1<0')).toBe('FALSE')
    expect(evaluate('2>=2')).toBe('TRUE')
    expect(evaluate('2<=1')).toBe('FALSE')
    expect(evaluate('3=3')).toBe('TRUE')
    expect(evaluate('3<>3')).toBe('FALSE')
  })

  it('compares text case-insensitively, as Excel does', () => {
    // The surprising one for anyone expecting JS `===`.
    expect(evaluate('"a"="A"')).toBe('TRUE')
    expect(evaluate('"abc"<"abd"')).toBe('TRUE')
  })

  it('never equates a number with text, and sorts numbers first', () => {
    expect(evaluate('1="1"')).toBe('FALSE')
    expect(evaluate('1<"a"')).toBe('TRUE')
  })

  it('sits below & in precedence, so a concatenation is compared whole', () => {
    expect(evaluate('"a"&"b"="ab"')).toBe('TRUE')
  })

  it('feeds arithmetic as 1 and 0', () => {
    // What makes `=(A1>2)*5` work instead of erroring.
    expect(evaluate('(1>0)*5')).toBe('5')
    expect(evaluate('(1<0)*5')).toBe('0')
  })

  it('compares cell references', () => {
    expect(evaluate('A1>B1', [['5', '3']])).toBe('TRUE')
    expect(evaluate('A1=A2', [['5'], ['5']])).toBe('TRUE')
  })
})

describe('IF and the logic functions (SHEETFN-1)', () => {
  it('picks a branch', () => {
    expect(evaluate('IF(1>0,"big","small")')).toBe('big')
    expect(evaluate('IF(1<0,"big","small")')).toBe('small')
  })

  it('treats a missing third argument as FALSE rather than erroring', () => {
    expect(evaluate('IF(1<0,"yes")')).toBe('FALSE')
  })

  it('nests', () => {
    expect(evaluate('IF(A1>10,"high",IF(A1>5,"mid","low"))', [['7']])).toBe('mid')
    expect(evaluate('IF(A1>10,"high",IF(A1>5,"mid","low"))', [['2']])).toBe('low')
  })

  it('works on a cell condition and returns a computed branch', () => {
    expect(evaluate('IF(A1="yes",B1*2,0)', [['yes', '21']])).toBe('42')
  })

  it('IFERROR catches an error and passes a good value through', () => {
    expect(evaluate('IFERROR(1/0,"oops")')).toBe('oops')
    expect(evaluate('IFERROR(6/2,"oops")')).toBe('3')
  })

  it('AND, OR and NOT', () => {
    expect(evaluate('AND(1>0,2>1)')).toBe('TRUE')
    expect(evaluate('AND(1>0,2<1)')).toBe('FALSE')
    expect(evaluate('OR(1<0,2>1)')).toBe('TRUE')
    expect(evaluate('OR(1<0,2<1)')).toBe('FALSE')
    expect(evaluate('NOT(1>0)')).toBe('FALSE')
  })

  it('combines into a real conditional', () => {
    expect(evaluate('IF(AND(A1>0,B1>0),"both","not both")', [['1', '2']])).toBe('both')
    expect(evaluate('IF(AND(A1>0,B1>0),"both","not both")', [['1', '-2']])).toBe('not both')
  })
})

describe('conditional aggregates (SHEETFN-1)', () => {
  const grid = [
    ['Paper', '10'],
    ['Ink', '25'],
    ['Paper', '5'],
    ['', ''],
  ]

  it('COUNTIF with a comparison criterion', () => {
    expect(evaluate('COUNTIF(B1:B4,">8")', grid)).toBe('2')
    expect(evaluate('COUNTIF(B1:B4,"<=10")', grid)).toBe('2')
  })

  it('COUNTIF with an equality criterion', () => {
    expect(evaluate('COUNTIF(A1:A4,"Paper")', grid)).toBe('2')
    // Case-insensitive, like every other text comparison here.
    expect(evaluate('COUNTIF(A1:A4,"paper")', grid)).toBe('2')
  })

  it('does not count blank cells against a comparison criterion', () => {
    // The trap: a blank cell reads as 0, so `<8` would match every empty row
    // of a column and silently inflate the count.
    expect(evaluate('COUNTIF(B1:B4,"<8")', grid)).toBe('1')
  })

  it('SUMIF over the tested range', () => {
    expect(evaluate('SUMIF(B1:B4,">8")', grid)).toBe('35')
  })

  it('SUMIF over a parallel range — the "sum B where A says x" shape', () => {
    expect(evaluate('SUMIF(A1:A4,"Paper",B1:B4)', grid)).toBe('15')
  })

  it('AVERAGEIF, and #DIV/0! when nothing matches', () => {
    expect(evaluate('AVERAGEIF(A1:A4,"Paper",B1:B4)', grid)).toBe('7.5')
    expect(evaluate('AVERAGEIF(A1:A4,"Toner",B1:B4)', grid)).toBe('#DIV/0!')
  })
})

describe('lookup functions (SHEETFN-1)', () => {
  const table = [
    ['Paper', '10', 'box'],
    ['Ink', '25', 'bottle'],
    ['Toner', '80', 'cartridge'],
  ]

  it('VLOOKUP finds a row and returns the requested column', () => {
    expect(evaluate('VLOOKUP("Ink",A1:C3,2)', table)).toBe('25')
    expect(evaluate('VLOOKUP("Ink",A1:C3,3)', table)).toBe('bottle')
  })

  it('VLOOKUP reports #N/A for a key that is not there', () => {
    expect(evaluate('VLOOKUP("Pens",A1:C3,2)', table)).toBe('#N/A')
  })

  it('VLOOKUP refuses an approximate match rather than guessing', () => {
    // Excel's DEFAULT is approximate, which returns a confidently wrong row on
    // unsorted data. Refusing is honest; implementing it half-correctly would
    // be worse than not having it. The formula falls back to literal text.
    expect(evaluate('VLOOKUP("Ink",A1:C3,2,TRUE())', table)).toBeNull()
  })

  it('MATCH returns a position, INDEX reads one', () => {
    expect(evaluate('MATCH("Toner",A1:A3,0)', table)).toBe('3')
    expect(evaluate('INDEX(B1:B3,2)', table)).toBe('25')
  })

  it('INDEX/MATCH composes, which is the point of both', () => {
    expect(evaluate('INDEX(B1:B3,MATCH("Ink",A1:A3,0))', table)).toBe('25')
  })

  it('INDEX reports #REF! past the end of the range', () => {
    expect(evaluate('INDEX(B1:B3,9)', table)).toBe('#REF!')
  })
})

describe('maths functions (SHEETFN-1)', () => {
  it('ROUND to a digit count', () => {
    expect(evaluate('ROUND(2.567,2)')).toBe('2.57')
    expect(evaluate('ROUND(2.5,0)')).toBe('3')
    expect(evaluate('ROUND(1234.5678,-2)')).toBe('1200')
  })

  it('ROUNDUP and ROUNDDOWN go away from and toward zero, not up and down', () => {
    // The negative-number trap: ceil(-1.2) is -1, but ROUNDUP(-1.2) is -2.
    expect(evaluate('ROUNDUP(1.1,0)')).toBe('2')
    expect(evaluate('ROUNDUP(-1.1,0)')).toBe('-2')
    expect(evaluate('ROUNDDOWN(1.9,0)')).toBe('1')
    expect(evaluate('ROUNDDOWN(-1.9,0)')).toBe('-1')
  })

  it('MOD takes the sign of the divisor, as Excel does', () => {
    // JavaScript's % takes the sign of the dividend: -3 % 2 is -1, while
    // Excel's MOD(-3,2) is 1.
    expect(evaluate('MOD(-3,2)')).toBe('1')
    expect(evaluate('MOD(3,-2)')).toBe('-1')
    expect(evaluate('MOD(7,3)')).toBe('1')
    expect(evaluate('MOD(1,0)')).toBe('#DIV/0!')
  })

  it('ABS, INT, SIGN, SQRT and POWER', () => {
    expect(evaluate('ABS(-4)')).toBe('4')
    expect(evaluate('INT(-1.5)')).toBe('-2') // floor, matching Excel
    expect(evaluate('SIGN(-9)')).toBe('-1')
    expect(evaluate('SQRT(9)')).toBe('3')
    expect(evaluate('SQRT(-1)')).toBe('#NUM!')
    expect(evaluate('POWER(2,10)')).toBe('1024')
  })
})

describe('text functions (SHEETFN-1)', () => {
  it('LEFT, RIGHT, MID and LEN', () => {
    expect(evaluate('LEFT("spreadsheet",6)')).toBe('spread')
    expect(evaluate('RIGHT("spreadsheet",5)')).toBe('sheet')
    expect(evaluate('MID("spreadsheet",7,5)')).toBe('sheet')
    expect(evaluate('LEN("spreadsheet")')).toBe('11')
  })

  it('LEFT and RIGHT default to one character', () => {
    expect(evaluate('LEFT("abc")')).toBe('a')
    expect(evaluate('RIGHT("abc")')).toBe('c')
  })

  it('RIGHT with a zero count is empty, not the whole string', () => {
    // `slice(-0)` is `slice(0)` — the whole string — which is the bug this
    // guards against.
    expect(evaluate('RIGHT("abc",0)')).toBe('')
  })

  it('TRIM collapses internal runs as well as trimming the ends', () => {
    expect(evaluate('TRIM("  a   b  ")')).toBe('a b')
  })

  it('UPPER and LOWER', () => {
    expect(evaluate('UPPER("abC")')).toBe('ABC')
    expect(evaluate('LOWER("AbC")')).toBe('abc')
  })

  it('reads text out of cells', () => {
    expect(evaluate('UPPER(A1)', [['hello']])).toBe('HELLO')
    expect(evaluate('LEN(A1)', [['hello']])).toBe('5')
  })
})

describe('type tests and counting (SHEETFN-1)', () => {
  const grid = [['5'], [''], ['text']]

  it('ISBLANK, ISNUMBER, ISTEXT', () => {
    expect(evaluate('ISNUMBER(A1)', grid)).toBe('TRUE')
    expect(evaluate('ISBLANK(A2)', grid)).toBe('TRUE')
    expect(evaluate('ISBLANK(A1)', grid)).toBe('FALSE')
    expect(evaluate('ISTEXT(A3)', grid)).toBe('TRUE')
    expect(evaluate('ISTEXT(A1)', grid)).toBe('FALSE')
  })

  it('ISERROR', () => {
    expect(evaluate('ISERROR(1/0)')).toBe('TRUE')
    expect(evaluate('ISERROR(1/1)')).toBe('FALSE')
  })

  it('COUNTBLANK, and COUNT still ignoring text', () => {
    expect(evaluate('COUNTBLANK(A1:A3)', grid)).toBe('1')
    expect(evaluate('COUNT(A1:A3)', grid)).toBe('1')
    expect(evaluate('COUNTA(A1:A3)', grid)).toBe('2')
  })

  it('PRODUCT and MEDIAN', () => {
    expect(evaluate('PRODUCT(A1:A3)', [['2'], ['3'], ['4']])).toBe('24')
    expect(evaluate('MEDIAN(A1:A3)', [['1'], ['9'], ['5']])).toBe('5')
    expect(evaluate('MEDIAN(A1:A4)', [['1'], ['2'], ['3'], ['4']])).toBe('2.5')
  })
})

describe('errors propagate instead of silently vanishing (SHEETFN-1)', () => {
  it('an aggregate over an erroring argument reports the error', () => {
    // Previously this fell back to showing the literal formula text; Excel
    // reports the error, which is both more useful and more honest.
    expect(evaluate('SUM(1/0,1)')).toBe('#DIV/0!')
  })

  it('a non-numeric scalar in an aggregate is #VALUE!, but text INSIDE a range is skipped', () => {
    // The distinction that makes `SUM(A1:A9)` work on a column with a header
    // while `SUM("abc")` still reports a problem.
    expect(evaluate('SUM("abc")')).toBe('#VALUE!')
    expect(evaluate('SUM(A1:A3)', [['Header'], ['2'], ['3']])).toBe('5')
  })

  it('still falls back to literal text for a formula it cannot evaluate at all', () => {
    expect(evaluateFormula('NOSUCHFN(1)', EMPTY)).toEqual({ ok: false })
    expect(evaluateFormula('IF(', EMPTY)).toEqual({ ok: false })
  })
})
