import { describe, expect, it } from 'vitest'

import { columnIndexToLetters, columnLettersToIndex, parseCellRef, formatCellRef, parseCellRange } from '../cellRef'

describe('columnIndexToLetters / columnLettersToIndex', () => {
  it.each([
    [0, 'A'],
    [25, 'Z'],
    [26, 'AA'],
    [27, 'AB'],
    [51, 'AZ'],
    [701, 'ZZ'],
    [702, 'AAA'],
  ])('round-trips index %i <-> %s', (index, letters) => {
    expect(columnIndexToLetters(index)).toBe(letters)
    expect(columnLettersToIndex(letters)).toBe(index)
    expect(columnLettersToIndex(letters.toLowerCase())).toBe(index)
  })
})

describe('parseCellRef / formatCellRef', () => {
  it('parses a plain reference', () => {
    expect(parseCellRef('B3')).toEqual({ row: 2, col: 1 })
  })

  it('parses an absolute-anchored reference', () => {
    expect(parseCellRef('$B$3')).toEqual({ row: 2, col: 1 })
  })

  it('is case-insensitive', () => {
    expect(parseCellRef('b3')).toEqual({ row: 2, col: 1 })
  })

  it('returns null for malformed input', () => {
    expect(parseCellRef('not-a-ref')).toBeNull()
    expect(parseCellRef('3B')).toBeNull()
    expect(parseCellRef('')).toBeNull()
  })

  it('formats back to the canonical A1 form', () => {
    expect(formatCellRef({ row: 2, col: 1 })).toBe('B3')
    expect(formatCellRef({ row: 0, col: 0 })).toBe('A1')
  })
})

describe('parseCellRange', () => {
  it('parses a two-corner range', () => {
    expect(parseCellRange('A1:B3')).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } })
  })

  it('normalizes a reversed range so start is top-left', () => {
    expect(parseCellRange('B3:A1')).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } })
  })

  it('treats a single cell as a 1x1 range', () => {
    expect(parseCellRange('C4')).toEqual({ start: { row: 3, col: 2 }, end: { row: 3, col: 2 } })
  })

  it('returns null for a malformed range', () => {
    expect(parseCellRange('A1:???')).toBeNull()
    expect(parseCellRange('')).toBeNull()
  })
})
