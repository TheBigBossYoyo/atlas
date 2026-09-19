/**
 * Re-anchoring OOXML ranges through row/column inserts and deletes (USR-17).
 *
 * `xlsxPassthrough.ts` already keeps a sheet's own cell/row/column data in
 * step with an insert or delete via `rowSources`/`colSources` (new index ->
 * the ORIGINAL index it came from, or `null` for a row/column that did not
 * exist in the original file). Several other OOXML constructs are plain
 * `sqref`/`ref`-style rectangular ranges written against the sheet's cell
 * grid — conditional formatting, data validation, hyperlinks, the sheet-level
 * autoFilter, and (per-sheet) defined names — and none of them were shifted
 * when a row or column moved, so after an insert/delete they silently
 * pointed at the wrong cells (or, worse, cells that no longer existed).
 *
 * This module derives the shift purely from `rowSources`/`colSources`
 * (there is no separate "insert/delete op log" to replay): for one edge of a
 * range, if the original row/column still exists, its new index is used
 * directly; if it was deleted, the nearest SURVIVING row/column is used
 * instead (rounding the top/left edge down towards row 0, the bottom/right
 * edge up towards the deleted one) — the same "shrink to the nearest
 * survivor" and "row inserted inside the range grows it" behavior Excel
 * itself applies to merged cells and table ranges (see
 * `spreadsheetDocument.ts`'s `insertRowAt`/`deleteRowAt` merge bookkeeping
 * and `spreadsheetTables.ts`'s `tablesAfterRowInsert`/`tablesAfterRowDelete`,
 * which this reproduces for ranges that live in raw XML instead of the
 * document model). A range entirely consumed by a delete has nowhere left to
 * anchor to and is dropped (`null`), matching what Excel does to a named
 * range or a filter whose whole area was deleted.
 */
import { decodeRange, encodeCol } from './spreadsheetTables'

/** `rowSources`/`colSources` as carried on `EditableSheet`: new index -> original index, or `null` for a row/column added since load. `undefined` means "no structural changes recorded" (ranges pass through unchanged). */
export type IndexSources = ReadonlyArray<number | null> | undefined

export type RangeBounds = { readonly r0: number; readonly c0: number; readonly r1: number; readonly c1: number }

function existsMapFrom(sources: IndexSources): Map<number, number> | null {
  if (!sources) return null
  const map = new Map<number, number>()
  sources.forEach((source, newIndex) => {
    if (source !== null) map.set(source, newIndex)
  })
  return map
}

function maxDefined(sources: IndexSources): number {
  let max = -1
  if (sources) {
    for (const source of sources) if (source !== null && source > max) max = source
  }
  return max
}

/**
 * Maps one edge of a range from its original index to its new one.
 * `roundUp`: the top/left edge searches forward (towards higher indexes,
 * i.e. into the range) for the nearest surviving row/column when its own
 * was deleted; the bottom/right edge (`roundUp` false) searches backward
 * (towards lower indexes, i.e. still into the range). Returns `null` when no
 * surviving row/column exists on that side at all.
 */
function mapEdge(originalIndex: number, exists: Map<number, number> | null, maxOriginal: number, roundUp: boolean): number | null {
  if (!exists) return originalIndex
  const direct = exists.get(originalIndex)
  if (direct !== undefined) return direct
  if (roundUp) {
    for (let i = originalIndex + 1; i <= maxOriginal; i++) {
      const mapped = exists.get(i)
      if (mapped !== undefined) return mapped
    }
    return null
  }
  for (let i = originalIndex - 1; i >= 0; i--) {
    const mapped = exists.get(i)
    if (mapped !== undefined) return mapped
  }
  return null
}

/**
 * Re-anchors one rectangular range. Returns `null` when the range was fully
 * inside a deleted row/column span and has nothing left to anchor to.
 */
export function remapRange(range: RangeBounds, rowSources: IndexSources, colSources: IndexSources): RangeBounds | null {
  const rowExists = existsMapFrom(rowSources)
  const colExists = existsMapFrom(colSources)
  const maxOldRow = maxDefined(rowSources)
  const maxOldCol = maxDefined(colSources)

  const r0 = mapEdge(range.r0, rowExists, maxOldRow, true)
  const r1 = mapEdge(range.r1, rowExists, maxOldRow, false)
  const c0 = mapEdge(range.c0, colExists, maxOldCol, true)
  const c1 = mapEdge(range.c1, colExists, maxOldCol, false)
  if (r0 === null || r1 === null || c0 === null || c1 === null || r0 > r1 || c0 > c1) return null
  return { r0, c0, r1, c1 }
}

export function encodeRangeRef(range: RangeBounds): string {
  const start = `${encodeCol(range.c0)}${range.r0 + 1}`
  if (range.r0 === range.r1 && range.c0 === range.c1) return start
  return `${start}:${encodeCol(range.c1)}${range.r1 + 1}`
}

/**
 * Re-anchors a `sqref`/hyperlink-`ref`/autoFilter-`ref`-style value — one or
 * more (space-separated) ranges. A sub-range fully consumed by a delete is
 * dropped; `null` is returned only when EVERY sub-range was dropped, so the
 * caller should remove the owning element entirely (an empty `sqref` is not
 * valid OOXML).
 */
export function remapSqref(sqref: string, rowSources: IndexSources, colSources: IndexSources): string | null {
  const pieces = sqref.trim().split(/\s+/).filter(Boolean)
  const mapped: string[] = []
  for (const piece of pieces) {
    const range = decodeRange(piece)
    if (!range) continue // malformed piece in the source file — drop rather than propagate garbage
    const next = remapRange(range, rowSources, colSources)
    if (next) mapped.push(encodeRangeRef(next))
  }
  return mapped.length > 0 ? mapped.join(' ') : null
}
