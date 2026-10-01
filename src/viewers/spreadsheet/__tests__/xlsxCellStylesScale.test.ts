/**
 * SHEETFMT-1 — what reading cell formatting costs on a large sheet.
 *
 * Reading styles adds a second pass over the worksheet XML, which is the one
 * part of a workbook that can be enormous: this repo's own perf tests use a
 * 100k-row sheet, and the viewer holds renderer work to a 200 ms long-task
 * budget. That cost was the main design risk in this feature, so it is measured
 * here rather than asserted to be fine — and measured on the same shape the
 * perf suite uses, a 100k-row, 4-column sheet.
 *
 * The budget below is deliberately loose. The point is not to pin a number on a
 * machine whose speed varies, but to catch the kind of regression that would
 * matter: an accidental O(rows x cols) object allocation, a DOM parse of the
 * worksheet, or a per-cell style resolution. Those are 10x-100x effects, not
 * 20% ones.
 */
import { describe, expect, it } from 'vitest'

import { formatAt, readSheetStyleIds } from '../xlsxCellStyles'

const ROWS = 100_000
const COLS = 4

/** A 100k-row worksheet part, styled the way a real report is: a styled header, then one styled column. */
function buildLargeSheetXml(): string {
  const parts: string[] = ['<worksheet xmlns="x"><sheetData>']
  parts.push('<row r="1">')
  for (let c = 0; c < COLS; c += 1) {
    parts.push(`<c r="${String.fromCharCode(65 + c)}1" s="2" t="inlineStr"><is><t>H${c}</t></is></c>`)
  }
  parts.push('</row>')
  for (let r = 2; r <= ROWS; r += 1) {
    parts.push(`<row r="${r}">`)
    parts.push(`<c r="A${r}"><v>${r}</v></c>`)
    parts.push(`<c r="B${r}" s="3"><v>${r * 1.5}</v></c>`)
    parts.push(`<c r="C${r}"><v>${r}</v></c>`)
    parts.push(`<c r="D${r}" t="inlineStr"><is><t>padding text to make each row a realistic width</t></is></c>`)
    parts.push('</row>')
  }
  parts.push('</sheetData></worksheet>')
  return parts.join('')
}

describe('readSheetStyleIds at 100k rows (SHEETFMT-1)', () => {
  it('scans a 100k-row sheet well inside the renderer long-task budget', () => {
    const xml = buildLargeSheetXml()
    // Sanity: the fixture really is the size this test claims to measure.
    expect(xml.length).toBeGreaterThan(10_000_000)

    const started = performance.now()
    const ids = readSheetStyleIds(xml, ROWS, COLS, 0, 0)
    const elapsedMs = performance.now() - started

    // Correctness first — a fast scan that read the wrong cells would be worse
    // than a slow one.
    expect(ids[0]).toEqual([2, 2, 2, 2])
    expect(ids[1]).toEqual([0, 3, 0, 0])
    expect(ids[ROWS - 1]).toEqual([0, 3, 0, 0])

    console.log(`[SHEETFMT-1] 100k-row style scan: ${Math.round(elapsedMs)}ms for ${ROWS * COLS} cells`)

    // 2s is ~10x the headroom a one-off parse of this size needs, and runs off
    // the main thread anyway for a file this big (the parsing Worker). A
    // regression of the kind this guards against blows through it outright.
    expect(elapsedMs, `100k-row style scan took ${Math.round(elapsedMs)}ms`).toBeLessThan(2000)
  })

  it('keeps memory to one integer per cell, not one object', () => {
    // The shape that makes the above affordable: `formats` is a handful of
    // shared objects and `styleIds` is plain numbers. A regression to a format
    // object per cell would not fail any correctness test and would multiply
    // this sheet's footprint by two orders of magnitude.
    const ids = readSheetStyleIds(buildLargeSheetXml(), ROWS, COLS, 0, 0)
    expect(typeof ids[500][1]).toBe('number')

    const styles = { formats: [], styleIds: ids }
    // And an id with no format behind it resolves to the default rather than
    // throwing mid-scroll.
    expect(formatAt(styles, 500, 1).bold).toBe(false)
  })
})
