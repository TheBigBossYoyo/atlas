/**
 * SHEETFN-2 — date functions.
 *
 * A spreadsheet date is a serial NUMBER, so `TODAY()` is really 46266. In Excel
 * the cell's number format is what turns that into a date on screen; Atlas does
 * not yet apply a cell's number format to a formula result, so the evaluator
 * carries a `dateKind` flag on the value and renders a date as ISO. The flag
 * propagating through `+`/`-` the way Excel's own type inference does is what
 * makes `=TODAY()+7` show a date rather than a five-digit number.
 *
 * The cases that matter here are the ones a from-scratch implementation gets
 * wrong: `DATE(2024,2,30)` must not silently become March 1st (`Date.UTC`
 * normalizes it), `EDATE` must clamp to the end of a shorter month,
 * `WEEKDAY`'s numbering starts at Sunday = 1, and `DAYS` takes its arguments
 * end-first.
 */
import { describe, expect, it } from 'vitest'

import { evaluateFormula } from '../spreadsheetFormula'
import {
  dateFromSerial,
  isoDateFromSerial,
  isoDateTimeFromSerial,
  serialFromDate,
  serialFromIsoDate,
  todaySerial,
} from '../excelDate'

function gridLookup(rows: ReadonlyArray<ReadonlyArray<string>>) {
  return (row: number, col: number): string => rows[row]?.[col] ?? ''
}

function evaluate(formula: string, rows: ReadonlyArray<ReadonlyArray<string>> = []): string | null {
  const result = evaluateFormula(formula, gridLookup(rows))
  return result.ok ? result.text : null
}

describe('excelDate (SHEETFN-2)', () => {
  it('uses the 1899-12-30 epoch, so 1900-01-01 is serial 1', () => {
    expect(serialFromIsoDate('1900-01-01')).toBe(1)
  })

  it('agrees with Excel for a modern date', () => {
    // 2026-10-02 is serial 46297 in Excel. Spot-checked against the epoch
    // rather than derived from the same code being tested.
    const serial = serialFromIsoDate('2026-10-02')!
    expect(isoDateFromSerial(serial)).toBe('2026-10-02')
    expect(serial).toBe(Math.round((Date.UTC(2026, 9, 2) - Date.UTC(1899, 11, 30)) / 86_400_000))
  })

  it('rejects an impossible date rather than rolling it over', () => {
    // `Date.UTC(2024, 1, 30)` is 2024-03-01, so a naive conversion answers for
    // a different day than the one asked about.
    expect(serialFromDate({ year: 2024, month: 2, day: 30 })).toBeNull()
    expect(serialFromDate({ year: 2024, month: 13, day: 1 })).toBeNull()
    expect(serialFromIsoDate('2024-02-30')).toBeNull()
  })

  it('accepts a real leap day', () => {
    expect(serialFromIsoDate('2024-02-29')).not.toBeNull()
    expect(serialFromIsoDate('2023-02-29')).toBeNull()
  })

  it('discards the time of day by flooring, not rounding', () => {
    // 18:00 on the 2nd must still be the 2nd.
    const serial = serialFromIsoDate('2026-10-02')! + 0.75
    expect(isoDateFromSerial(serial)).toBe('2026-10-02')
  })

  it('refuses serial 0, which Excel shows as the fictitious 1900-01-00', () => {
    expect(dateFromSerial(0)).toBeNull()
    expect(isoDateFromSerial(-5)).toBeNull()
  })

  it('renders a time of day, carrying a rounded-up fraction into the next day', () => {
    const base = serialFromIsoDate('2026-10-02')!
    expect(isoDateTimeFromSerial(base + 0.5)).toBe('2026-10-02 12:00')
    // A fraction a hair under a full day must not render as "24:00".
    expect(isoDateTimeFromSerial(base + 0.99999)).toBe('2026-10-03 00:00')
  })
})

describe('date functions (SHEETFN-2)', () => {
  it('TODAY returns today, rendered as a date rather than a serial', () => {
    const expected = isoDateFromSerial(todaySerial())
    expect(evaluate('TODAY()')).toBe(expected)
    // The point of the dateKind flag: a bare serial would be unreadable.
    expect(evaluate('TODAY()')).toMatch(/^\d{4}-\d{2}-\d{2}$/)
  })

  it('NOW includes a time of day', () => {
    expect(evaluate('NOW()')).toMatch(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/)
  })

  it('DATE builds a date from its parts', () => {
    expect(evaluate('DATE(2026,10,2)')).toBe('2026-10-02')
  })

  it('DATE reports #NUM! for a day that does not exist', () => {
    expect(evaluate('DATE(2024,2,30)')).toBe('#NUM!')
    expect(evaluate('DATE(2024,13,1)')).toBe('#NUM!')
  })

  it('YEAR, MONTH and DAY read a date apart', () => {
    expect(evaluate('YEAR(DATE(2026,10,2))')).toBe('2026')
    expect(evaluate('MONTH(DATE(2026,10,2))')).toBe('10')
    expect(evaluate('DAY(DATE(2026,10,2))')).toBe('2')
  })

  it('reads a date out of a cell holding ISO text', () => {
    // What a user actually has: a column of typed dates, which the grid shows
    // as text. Requiring a serial here would make the functions unusable.
    const grid = [['2026-10-02']]
    expect(evaluate('YEAR(A1)', grid)).toBe('2026')
    expect(evaluate('MONTH(A1)', grid)).toBe('10')
  })

  it('WEEKDAY numbers from Sunday = 1', () => {
    // 2026-10-02 is a Friday, so 6 with Excel's default numbering.
    expect(evaluate('WEEKDAY(DATE(2026,10,2))')).toBe('6')
    // 1900-01-01 was a MONDAY, so 2 — checked against the calendar, not from
    // memory, which is what made the first implementation off by one.
    expect(evaluate('WEEKDAY(1)')).toBe('2')
  })

  it('DAYS counts between two dates, end first', () => {
    // The argument order reads backwards from its name, and is Excel's.
    expect(evaluate('DAYS(DATE(2026,10,10),DATE(2026,10,2))')).toBe('8')
    expect(evaluate('DAYS(DATE(2026,10,2),DATE(2026,10,10))')).toBe('-8')
  })

  it('EDATE clamps to the end of a shorter month', () => {
    // The 31st of January plus one month is the 28th of February, not the 3rd
    // of March — the case a naive month increment gets wrong.
    expect(evaluate('EDATE(DATE(2026,1,31),1)')).toBe('2026-02-28')
    expect(evaluate('EDATE(DATE(2024,1,31),1)')).toBe('2024-02-29') // leap year
    expect(evaluate('EDATE(DATE(2026,3,15),-1)')).toBe('2026-02-15')
  })

  it('EDATE crosses a year boundary in both directions', () => {
    expect(evaluate('EDATE(DATE(2026,12,15),1)')).toBe('2027-01-15')
    expect(evaluate('EDATE(DATE(2026,1,15),-1)')).toBe('2025-12-15')
  })

  it('EOMONTH returns the last day of the target month', () => {
    expect(evaluate('EOMONTH(DATE(2026,10,2),0)')).toBe('2026-10-31')
    expect(evaluate('EOMONTH(DATE(2026,1,15),1)')).toBe('2026-02-28')
  })
})

describe('date arithmetic (SHEETFN-2)', () => {
  it('a date plus a number is still a date', () => {
    // The whole reason dateKind propagates: this is the single most common
    // thing anyone does with a date, and without it the answer is a serial.
    expect(evaluate('DATE(2026,10,2)+7')).toBe('2026-10-09')
    expect(evaluate('7+DATE(2026,10,2)')).toBe('2026-10-09')
  })

  it('a date minus a number is still a date', () => {
    expect(evaluate('DATE(2026,10,2)-1')).toBe('2026-10-01')
  })

  it('a date minus a date is a count of days, not a date', () => {
    expect(evaluate('DATE(2026,10,10)-DATE(2026,10,2)')).toBe('8')
  })

  it('a date added to a date is not a date', () => {
    // Nonsense arithmetically, but it must not render as a date in the year
    // 4000-something either.
    expect(evaluate('DATE(2026,10,2)+DATE(2026,10,2)')).toMatch(/^\d+$/)
  })

  it('multiplying a date gives a plain number', () => {
    expect(evaluate('DATE(2026,10,2)*2')).toMatch(/^\d+$/)
  })

  it('composes with the rest of the library', () => {
    expect(evaluate('IF(DATE(2026,10,2)>DATE(2026,1,1),"later","earlier")')).toBe('later')
    expect(evaluate('YEAR(EDATE(DATE(2026,12,1),1))')).toBe('2027')
    expect(evaluate('IFERROR(DATE(2024,2,30),"bad date")')).toBe('bad date')
  })

  it('reports #VALUE! for a non-date argument', () => {
    expect(evaluate('YEAR("not a date")')).toBe('#VALUE!')
  })
})
