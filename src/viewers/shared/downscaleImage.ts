/**
 * S18/SLD-19 — the slide thumbnail rail previously reused the exact same
 * full-resolution `data:` URL as the interactive main view for every
 * embedded image, so the browser decoded (and kept resident) a full-size
 * bitmap for each one, even though the rail only ever paints it at a few
 * dozen CSS pixels. This produces a much smaller bitmap once, up front, and
 * caches the result by `(src, target size)` so re-mounting a thumbnail row
 * (scrolling the virtualized rail) or re-rendering it (zoom, resize) reuses
 * the same downscaled image instead of re-decoding or re-resizing it.
 *
 * Feature-detects `createImageBitmap`/`OffscreenCanvas` and falls back to
 * the original `src` unchanged wherever either is unavailable (jsdom under
 * Vitest, or a genuinely unsupporting runtime) or the resize itself fails —
 * a correctness fallback (still renders the real image), not an error.
 *
 * Review fix (wave-3 shell-polish): each downscaled result that actually
 * gets re-encoded is handed back as a `blob:` object URL (`renderDownscaled`
 * below) — a *separate* browser-held resource from the cache entry itself,
 * which stays alive (backing bitmap memory and all) until explicitly
 * revoked or the document unloads. The cache is capped at
 * `MAX_CACHE_ENTRIES` and evicts oldest-inserted first, revoking that
 * entry's object URL once it's known (immediately if already resolved,
 * otherwise once its promise settles) — without this, a single long-running
 * session that opens many slide decks in turn would accumulate every
 * thumbnail bitmap it has ever decoded for the life of the process, never
 * released even after the deck that used them is long closed.
 */

/** Generous headroom over a typical single deck's embedded-image count,
 * while still bounding a long session that opens many decks in turn. */
const MAX_CACHE_ENTRIES = 200

const downscaledImageCache = new Map<string, Promise<string>>()

function revokeIfObjectUrl(url: string): void {
  if (url.startsWith('blob:')) {
    URL.revokeObjectURL(url)
  }
}

function evictOldestIfOverCapacity(): void {
  if (downscaledImageCache.size <= MAX_CACHE_ENTRIES) {
    return
  }
  // `Map` iterates in insertion order, so its first key is the
  // oldest-inserted entry — not a true LRU (a re-requested hit doesn't move
  // to the back), but sufficient to bound growth for this cache's access
  // pattern (thumbnails requested roughly once per rendered slide).
  const oldestKey = downscaledImageCache.keys().next().value
  if (oldestKey === undefined) {
    return
  }
  const evicted = downscaledImageCache.get(oldestKey)
  downscaledImageCache.delete(oldestKey)
  evicted?.then(revokeIfObjectUrl).catch(() => {
    // The evicted entry never resolved to a usable URL — nothing to revoke.
  })
}

function supportsOffscreenDownscale(): boolean {
  return typeof createImageBitmap === 'function' && typeof OffscreenCanvas === 'function'
}

async function renderDownscaled(src: string, maxDimensionPx: number): Promise<string> {
  const response = await fetch(src)
  const blob = await response.blob()
  const bitmap = await createImageBitmap(blob)

  try {
    const longestSide = Math.max(bitmap.width, bitmap.height)
    if (longestSide <= maxDimensionPx) {
      // Already at or below the target — a re-encode would only cost time
      // and (for lossy formats) quality, with no size benefit.
      return src
    }

    const scale = maxDimensionPx / longestSide
    const targetWidth = Math.max(1, Math.round(bitmap.width * scale))
    const targetHeight = Math.max(1, Math.round(bitmap.height * scale))

    const canvas = new OffscreenCanvas(targetWidth, targetHeight)
    const ctx = canvas.getContext('2d')
    if (!ctx) {
      return src
    }
    ctx.drawImage(bitmap, 0, 0, targetWidth, targetHeight)

    const outBlob = await canvas.convertToBlob({ type: 'image/png' })
    return URL.createObjectURL(outBlob)
  } finally {
    bitmap.close()
  }
}

/**
 * Returns a cached, downscaled version of `src` capped to `maxDimensionPx`
 * on its longest side. Resolves to the original `src` unchanged when
 * downscaling isn't supported, wouldn't help, or fails for any reason —
 * this never rejects, so a caller can always render its result directly.
 */
export function getDownscaledImage(src: string, maxDimensionPx: number): Promise<string> {
  const cacheKey = `${maxDimensionPx}:${src}`
  const cached = downscaledImageCache.get(cacheKey)
  if (cached) {
    return cached
  }

  const promise =
    supportsOffscreenDownscale() && src.startsWith('data:')
      ? renderDownscaled(src, maxDimensionPx).catch(() => src)
      : Promise.resolve(src)

  downscaledImageCache.set(cacheKey, promise)
  evictOldestIfOverCapacity()
  return promise
}
