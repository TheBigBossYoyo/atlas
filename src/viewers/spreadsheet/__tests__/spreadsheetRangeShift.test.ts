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

// ---------------------------------------------------------------------------
// SHEET-3 — whole-column (`A:A`) and whole-row (`1:1`) ranges. Before this
// fix `decodeRange` (which only understands `A1`-style cell corners) could
// not parse either shape, so `remapSqref` treated them as malformed input
// and dropped them — `updateShiftedRanges` in `xlsxPassthrough.ts` then
// removed the owning conditional format / data validation / hyperlink /
// autoFilter outright, on ANY row or column insert/delete anywhere on the
// sheet, not just one that actually touched the range.
// ---------------------------------------------------------------------------

describe('remapSqref — whole-column and whole-row ranges', () => {
  it('re-anchors a whole-column range through a column insert', () => {
    const colSources = sourcesAfter(3, [], [0]) // a column inserted before A
    expect(remapSqref('A:A', undefined, colSources)).toBe('B:B')
  })

  it('re-anchors a whole-column range through a column delete', () => {
    const colSources = sourcesAfter(3, [0], []) // column A deleted
    expect(remapSqref('B:B', undefined, colSources)).toBe('A:A')
  })

  it('re-anchors a multi-column whole-column range', () => {
    const colSources = sourcesAfter(4, [], [0])
    expect(remapSqref('A:B', undefined, colSources)).toBe('B:C')
  })

  it('leaves a whole-column range alone when only rows changed (a different axis)', () => {
    const rowSources = sourcesAfter(4, [], [1])
    expect(remapSqref('A:A', rowSources, undefined)).toBe('A:A')
  })

  it('drops a whole-column range entirely consumed by a delete', () => {
    const colSources = sourcesAfter(1, [0], [])
    expect(remapSqref('A:A', undefined, colSources)).toBeNull()
  })

  it('re-anchors a whole-row range through a row insert', () => {
    const rowSources = sourcesAfter(3, [], [0]) // a row inserted before row 1
    expect(remapSqref('1:1', rowSources, undefined)).toBe('2:2')
  })

  it('re-anchors a whole-row range through a row delete', () => {
    const rowSources = sourcesAfter(3, [0], []) // row 1 deleted
    expect(remapSqref('2:2', rowSources, undefined)).toBe('1:1')
  })

  it('re-anchors a multi-row whole-row range', () => {
    const rowSources = sourcesAfter(4, [], [0])
    expect(remapSqref('1:2', rowSources, undefined)).toBe('2:3')
  })

  it('leaves a whole-row range alone when only columns changed (a different axis)', () => {
    const colSources = sourcesAfter(4, [], [1])
    expect(remapSqref('1:1', undefined, colSources)).toBe('1:1')
  })

  it('drops a whole-row range entirely consumed by a delete', () => {
    const rowSources = sourcesAfter(1, [0], [])
    expect(remapSqref('1:1', rowSources, undefined)).toBeNull()
  })

  it('keeps a whole-column range unchanged with identity sources', () => {
    expect(remapSqref('C:C', undefined, [0, 1, 2, 3])).toBe('C:C')
  })

  it('still drops a genuinely malformed piece rather than propagate garbage', () => {
    const colSources = sourcesAfter(3, [], [0])
    expect(remapSqref('A:A Sheet1!A1 1A:2B', undefined, colSources)).toBe('B:B')
  })
})
