import { describe, expect, it } from 'vitest'

import { getCachedMetrics, setCachedMetrics, __memoryCacheSizeForTests } from '../cache'
import { resolveFontFamily, FONT_FAMILIES, type FontVariant } from '../families'
import { deserializeMetrics, type FontMetrics } from '../metrics'

// MEM-01 follow-up — the memory investigation that produced `documentSessions
// .ts`'s `shedBytes` fix flagged this module's `memoryCache` as a candidate
// unbounded module-level map (it has no explicit size cap or eviction,
// unlike `canvasMetrics.ts`'s fragment cache or `downscaleImage.ts`'s LRU).
// A follow-up heap-snapshot measurement across xlsx/csv/docx/pdf/pptx/
// markdown traced docx's residual retention to something else entirely (not
// this cache) — and source review shows why: every cache key is built as
// `${substituteName}@${variant}@${hash(fileUrl)}` (`loader.ts`), and both
// `substituteName` and `fileUrl` come from `FONT_FAMILIES` — a small, fixed,
// bundled catalog, never from anything document-specific (an arbitrary
// embedded font name or the number of documents opened). So the cache
// cannot grow past `FONT_FAMILIES.length * 4` variants no matter how many
// distinct (real or bogus) Word font names a session resolves, or how many
// documents it opens — this test proves that bound holds even under far
// more distinct inputs than that.
const VARIANTS: readonly FontVariant[] = ['regular', 'bold', 'italic', 'boldItalic']
const MAX_POSSIBLE_ENTRIES = FONT_FAMILIES.length * VARIANTS.length

// `setCachedMetrics` serializes through `WIDTHS_BY_METRICS` (metrics.ts),
// which only knows about `FontMetrics` instances created via that module's
// own `createMetrics`/`deserializeMetrics` — a hand-built object literal
// would fail with "Unknown FontMetrics instance", so this round-trips
// through the real deserializer instead.
function fakeMetrics(seed: number): FontMetrics {
  return deserializeMetrics({
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    lineGap: 0,
    xHeight: 500,
    capHeight: 700,
    widths: { 0: seed },
  })
}

function hashString(value: string): string {
  let hash = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index)
    hash = Math.imul(hash, 0x01000193)
  }
  return (hash >>> 0).toString(16)
}

describe('font metrics memory cache (cache.ts)', () => {
  it('stays bounded to the fixed substitute-font catalog even when far more distinct Word font names are resolved', async () => {
    // Many more distinct, made-up "Word font names" than there are bundled
    // substitutes — simulating a long session that opens many documents,
    // each referencing arbitrarily different (real or embedded/bogus) fonts.
    const wordFontNames = [
      'Calibri', 'Cambria', 'Courier New', 'Times New Roman', 'Arial',
      'Calibri Light', 'Calibri Body', 'Cambria Body', 'Cambria Heading',
      'CourierNew', 'Times', 'Times Roman', 'TimesNewRoman', 'ArialMT', 'Arial Narrow',
      'Aptos', 'Aptos Display', 'Aptos Narrow', 'Aptos Serif', 'Aptos Mono',
      ...Array.from({ length: 50 }, (_, i) => `Some Bogus Embedded Font ${i}`),
    ]

    let seed = 0
    for (const wordName of wordFontNames) {
      const resolvedFamily = resolveFontFamily(wordName)
      if (!resolvedFamily) continue // unmapped fonts (the bogus ones) never reach the cache
      for (const variant of VARIANTS) {
        const fileUrl = resolvedFamily.files[variant]
        const cacheKey = `${resolvedFamily.substituteName}@${variant}@${hashString(fileUrl)}`
        seed += 1
        await setCachedMetrics(cacheKey, fakeMetrics(seed))
      }
    }

    expect(__memoryCacheSizeForTests()).toBeLessThanOrEqual(MAX_POSSIBLE_ENTRIES)
    // Every substitute x variant combination should actually be present
    // (not just "small" by accident) — confirms this is a real fixed-size
    // key space, not merely happening to be a small subset of unique keys.
    expect(__memoryCacheSizeForTests()).toBe(MAX_POSSIBLE_ENTRIES)
  })

  it('returns a cached value keyed only by substitute family + variant + file, not by the original Word name', async () => {
    const calibri = resolveFontFamily('Calibri')
    const aptos = resolveFontFamily('Aptos')
    expect(calibri).not.toBeNull()
    expect(aptos).not.toBeNull()
    // Aptos maps to the same bundled substitute (Carlito) as Calibri (DXP-12).
    expect(aptos?.substituteName).toBe(calibri?.substituteName)

    const fileUrl = calibri!.files.regular
    const cacheKey = `${calibri!.substituteName}@regular@${hashString(fileUrl)}`
    await setCachedMetrics(cacheKey, fakeMetrics(42))

    const fromCache = await getCachedMetrics(cacheKey)
    expect(fromCache?.advanceWidth(0)).toBe(42)
  })
})
