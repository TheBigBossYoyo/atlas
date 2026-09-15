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
 * Known trade-off: the in-memory cache is unbounded for the life of the
 * process and keyed by the full `data:` URL string. Slide decks embed a
 * bounded number of images per file and this cache never survives a reload,
 * so this is accepted rather than building a full LRU eviction scheme for a
 * wave-3 polish item — revisit if a future deck with hundreds of large
 * embedded images makes this cache's own footprint a problem.
 */

const downscaledImageCache = new Map<string, Promise<string>>()

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
  return promise
}
