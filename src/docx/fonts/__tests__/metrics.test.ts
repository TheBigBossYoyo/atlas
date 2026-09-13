import { readFile } from 'node:fs/promises'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { buildMetrics, deserializeMetrics, measureRun, serializeMetrics } from '../metrics'
import { parseTtf } from '../ttf'

const FIXTURE_PATH = path.resolve(process.cwd(), 'public/fonts/Carlito-Regular.ttf')
const CODEPOINT_A = 'A'.codePointAt(0) ?? 0
const CODEPOINT_SPACE = ' '.codePointAt(0) ?? 0
const CODEPOINT_MISSING = 0x10ffff

async function createMetrics() {
  const bytes = await readFile(FIXTURE_PATH)
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
  return buildMetrics(parseTtf(buffer))
}

describe('FontMetrics', () => {
  it('uses OS/2 typo metrics for ascender and descender', async () => {
    const metrics = await createMetrics()
    expect(metrics.ascender).toBeGreaterThan(0)
    expect(metrics.descender).toBeLessThan(0)
  })

  it('returns the known Carlito advance width for A', async () => {
    const metrics = await createMetrics()
    expect(metrics.advanceWidth(CODEPOINT_A)).toBe(1185)
  })

  it('reports glyph presence for mapped characters', async () => {
    const metrics = await createMetrics()
    expect(metrics.hasGlyph(CODEPOINT_A)).toBe(true)
    expect(metrics.hasGlyph(CODEPOINT_SPACE)).toBe(true)
  })

  it('returns false for unmapped characters', async () => {
    const metrics = await createMetrics()
    expect(metrics.hasGlyph(CODEPOINT_MISSING)).toBe(false)
  })

  it('returns a non-zero fallback width for unmapped characters', async () => {
    const metrics = await createMetrics()
    // We intentionally do NOT return 0 for unmapped codepoints: the browser
    // will still paint a visible fallback glyph at non-zero width, so the
    // layout MUST reserve space for it. We fall back to the advance width of
    // 'o' (representative lowercase), which for Carlito Regular is 1080 funits.
    expect(metrics.advanceWidth(CODEPOINT_MISSING)).toBeGreaterThan(0)
  })

  it('measures repeated A glyphs in points', async () => {
    const metrics = await createMetrics()
    const expected = 3 * (1185 / 2048) * 12
    expect(measureRun('AAA', metrics, 12)).toBeCloseTo(expected, 3)
  })

  it('measures an empty run as zero points', async () => {
    const metrics = await createMetrics()
    expect(measureRun('', metrics, 12)).toBe(0)
  })

  it('scales linearly with point size', async () => {
    const metrics = await createMetrics()
    const widthAt12 = measureRun('Atlas', metrics, 12)
    const widthAt24 = measureRun('Atlas', metrics, 24)
    expect(widthAt24).toBeCloseTo(widthAt12 * 2, 5)
  })

  it('serializes all core numeric metrics', async () => {
    const metrics = await createMetrics()
    const serialized = serializeMetrics(metrics)
    expect(serialized.unitsPerEm).toBe(2048)
    expect(serialized.widths[CODEPOINT_A]).toBe(1185)
  })

  it('deserializes into a working FontMetrics instance', async () => {
    const metrics = await createMetrics()
    const roundTripped = deserializeMetrics(serializeMetrics(metrics))
    expect(roundTripped.advanceWidth(CODEPOINT_A)).toBe(1185)
    expect(roundTripped.hasGlyph(CODEPOINT_A)).toBe(true)
  })
})
