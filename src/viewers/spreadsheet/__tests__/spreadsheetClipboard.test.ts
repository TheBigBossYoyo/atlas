import { describe, expect, it } from 'vitest'

import { rangeToTsv, tsvToRows } from '../spreadsheetClipboard'

describe('rangeToTsv', () => {
  it('joins cells with tabs and rows with newlines', () => {
    expect(
      rangeToTsv([
        ['a', 'b'],
        ['c', 'd'],
      ]),
    ).toBe('a\tb\nc\td')
  })

  it('serializes a single cell with no separators', () => {
    expect(rangeToTsv([['x']])).toBe('x')
  })
})

describe('tsvToRows', () => {
  it('parses tab/newline-separated text back into rows', () => {
    expect(tsvToRows('a\tb\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
  })

  it('round-trips with rangeToTsv', () => {
    const rows = [
      ['1', '2', '3'],
      ['4', '5', '6'],
    ]
    expect(tsvToRows(rangeToTsv(rows))).toEqual(rows)
  })

  it('strips Windows-style \\r\\n line endings', () => {
    expect(tsvToRows('a\tb\r\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
  })

  it('drops a single trailing newline instead of producing an empty extra row', () => {
    expect(tsvToRows('a\tb\n')).toEqual([['a', 'b']])
  })

  it('parses a single plain value with no separators as a 1x1 block', () => {
    expect(tsvToRows('hello')).toEqual([['hello']])
  })

  it('treats empty clipboard text as one empty cell', () => {
    expect(tsvToRows('')).toEqual([['']])
  })
})
