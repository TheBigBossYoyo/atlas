/**
 * USR-17 — pure range re-anchoring through row/column inserts and deletes.
 */
import { describe, expect, it } from 'vitest'

import { encodeRangeRef, remapRange, remapSqref, type RangeBounds } from '../spreadsheetRangeShift'

/** `sources[newIndex] = originalIndex | null`, built from a "delete these original indexes, insert `nulls` new ones at these new indexes" description — a thin convenience so tests read like the row/column op they model. */
function sourcesAfter(originalCount: number, deletedOriginal: ReadonlyArray<number>, insertedAt: ReadonlyArray<number>): (number | null)[] {
  const survivors = Array.from({ length: originalCount }, (_, i) => i).filter((i) => !deletedOriginal.includes(i))
  const result: (number | null)[] = [...survivors]
  // Insert `null`s from the highest index down so earlier insertions don't shift later ones.
  for (const at of [...insertedAt].sort((a, b) => b - a)) result.splice(at, 0, null)
  return result
}

describe('remapRange', () => {
  const range: RangeBounds = { r0: 1, c0: 2, r1: 2, c1: 2 } // C2:C3

  it('leaves a range untouched with identity sources', () => {
    expect(remapRange(range, [0, 1, 2], undefined)).toEqual(range)
    expect(remapRange(range, undefined, undefined)).toEqual(range)
  })

  it('shifts a range whole when the insertion is at or before its top edge', () => {
    // insertRowAt(atIndex=1): a new row lands exactly at the range's own top row.
    const rowSources = sourcesAfter(3, [], [1])
    expect(remapRange(range, rowSources, undefined)).toEqual({ r0: 2, c0: 2, r1: 3, c1: 2 })
  })

  it('grows a range that straddles the insertion point', () => {
    // A range covering rows 0-2; inserting at row 1 lands INSIDE it.
    const straddling: RangeBounds = { r0: 0, c0: 0, r1: 2, c1: 0 }
    const rowSources = sourcesAfter(3, [], [1])
    expect(remapRange(straddling, rowSources, undefined)).toEqual({ r0: 0, c0: 0, r1: 3, c1: 0 })
  })

  it('shrinks to the nearest surviving row when an edge row is deleted', () => {
    // Range rows 1-2; deleting original row 1 (the top of the range).
    const rowSources = sourcesAfter(3, [1], [])
    expect(remapRange(range, rowSources, undefined)).toEqual({ r0: 1, c0: 2, r1: 1, c1: 2 })
  })

  it('drops a range entirely consumed by a delete', () => {
    const singleRow: RangeBounds = { r0: 1, c0: 0, r1: 1, c1: 0 }
    const rowSources = sourcesAfter(3, [1], [])
    expect(remapRange(singleRow, rowSources, undefined)).toBeNull()
  })

  it('re-anchors columns the same way as rows', () => {
    const colSources = sourcesAfter(4, [], [0]) // a column inserted before A
    expect(remapRange(range, undefined, colSources)).toEqual({ r0: 1, c0: 3, r1: 2, c1: 3 })
  })
})

describe('encodeRangeRef', () => {
  it('encodes a single cell without a colon', () => {
    expect(encodeRangeRef({ r0: 0, c0: 0, r1: 0, c1: 0 })).toBe('A1')
  })

  it('encodes a rectangular range', () => {
    expect(encodeRangeRef({ r0: 1, c0: 2, r1: 2, c1: 2 })).toBe('C2:C3')
  })
})

describe('remapSqref', () => {
  it('re-anchors every sub-range in a space-separated union independently', () => {
    const rowSources = sourcesAfter(4, [], [1])
    expect(remapSqref('A1 C2:C3', rowSources, undefined)).toBe('A1 C3:C4')
  })

  it('drops only the sub-ranges fully consumed by a delete, keeping the rest', () => {
    const rowSources = sourcesAfter(3, [1], [])
    expect(remapSqref('A1 B2:B2', rowSources, undefined)).toBe('A1')
  })

  it('returns null when every sub-range was dropped', () => {
    const rowSources = sourcesAfter(3, [1], [])
    expect(remapSqref('B2:B2', rowSources, undefined)).toBeNull()
  })
})
