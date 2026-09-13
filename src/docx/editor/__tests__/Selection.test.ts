import { describe, expect, it } from 'vitest'

import {
  comparePositions,
  isRangeCollapsed,
  normalizeRange,
  positionEquals,
  type Position,
  type Range,
} from '../Selection'

function makePosition(
  paragraphPath: ReadonlyArray<number>,
  runIndex: number,
  charOffset: number,
): Position {
  return {
    paragraphPath,
    runIndex,
    charOffset,
  }
}

function makeRange(anchor: Position, focus: Position): Range {
  return {
    anchor,
    focus,
  }
}

describe('Selection helpers', () => {
  it('comparePositions returns 0 for identical positions', () => {
    const position = makePosition([0], 1, 3)

    expect(comparePositions(position, position)).toBe(0)
  })

  it('comparePositions orders earlier paragraph paths first', () => {
    expect(comparePositions(makePosition([0], 0, 0), makePosition([1], 0, 0))).toBe(-1)
  })

  it('comparePositions treats a shorter matching path as earlier', () => {
    expect(comparePositions(makePosition([0], 0, 0), makePosition([0, 1], 0, 0))).toBe(-1)
  })

  it('comparePositions orders by run index inside the same paragraph', () => {
    expect(comparePositions(makePosition([2], 0, 5), makePosition([2], 1, 0))).toBe(-1)
  })

  it('comparePositions orders by character offset inside the same run', () => {
    expect(comparePositions(makePosition([2], 1, 4), makePosition([2], 1, 7))).toBe(-1)
  })

  it('normalizeRange keeps already ordered ranges unchanged', () => {
    const anchor = makePosition([0], 0, 1)
    const focus = makePosition([0], 0, 4)
    const normalized = normalizeRange(makeRange(anchor, focus))

    expect(normalized).toEqual({ start: anchor, end: focus })
  })

  it('normalizeRange swaps anchor and focus when the range is reversed', () => {
    const anchor = makePosition([3], 2, 8)
    const focus = makePosition([1], 0, 0)
    const normalized = normalizeRange(makeRange(anchor, focus))

    expect(normalized).toEqual({ start: focus, end: anchor })
  })

  it('isRangeCollapsed returns true when anchor and focus match', () => {
    const position = makePosition([1, 2], 3, 4)

    expect(isRangeCollapsed(makeRange(position, position))).toBe(true)
  })

  it('isRangeCollapsed returns false when anchor and focus differ', () => {
    expect(isRangeCollapsed(makeRange(makePosition([1], 0, 0), makePosition([1], 0, 1)))).toBe(false)
  })

  it('positionEquals returns true for equal values across separate objects', () => {
    expect(positionEquals(makePosition([4, 1], 2, 9), makePosition([4, 1], 2, 9))).toBe(true)
  })

  it('positionEquals returns false when paragraph paths differ', () => {
    expect(positionEquals(makePosition([4], 2, 9), makePosition([4, 1], 2, 9))).toBe(false)
  })
})
