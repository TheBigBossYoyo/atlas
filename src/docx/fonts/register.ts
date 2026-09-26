/**
 * Browser-side font registration and metric resolution for the DOCX viewer.
 *
 * Extracted from `src/viewers/DocxViewer.tsx` (which was a 3,600-line file, the
 * bulk of it one component) because none of this is view code: it registers
 * `FontFace`s on the document and builds the `FontResolver` the paginator
 * measures with. Nothing here touches React, and the two callers that need it
 * are the viewer and its pagination hook.
 *
 * Three separate concerns, deliberately in one module because they share
 * `FONT_VARIANT_DESCRIPTORS` and the same fallback metrics:
 *   - `preloadDocxFonts` — one-time, app-lifetime registration of the five
 *     bundled substitute families.
 *   - `registerEmbeddedFonts`/`unregisterEmbeddedFonts` — per document, for
 *     fonts the file itself embeds (DEFER-4/DXP-13).
 *   - `createFontResolver` — what the paginator calls to measure a run.
 */
import type { FontResolver } from '../layout/types'
import {
  buildMetrics,
  FONT_FAMILIES,
  loadFontMetrics,
  parseTtf,
  resolveFontFamily,
  wrapWithCanvasAdvance,
  type EmbeddedFontFamily,
  type FontMetrics,
  type FontVariant,
} from './index'

export const DEFAULT_FONT_METRICS: FontMetrics = Object.freeze({
  unitsPerEm: 1000,
  ascender: 800,
  descender: -200,
  lineGap: 0,
  xHeight: 500,
  capHeight: 700,
  advanceWidth: () => 500,
  hasGlyph: () => true,
})

/**
 * Register every bundled substitute font via the FontFace API using URLs
 * emitted by Vite (`?url` imports in families.ts). This replaces the previous
 * CSS @font-face approach which used absolute `/fonts/...` URLs that resolve
 * to `file:///C:/fonts/...` in the packaged Electron app and 404, leaving
 * canvas measurements returning 0/NaN and producing blank pages (3.0.7 bug
 * on LaBoetie_Dossier_v2_Bac2026.docx). Registering both wordName and
 * substituteName guarantees the paginator's measured typeface is the one
 * the browser paints.
 */
const FONT_VARIANT_DESCRIPTORS: ReadonlyArray<{
  readonly variant: FontVariant
  readonly weight: string
  readonly style: string
}> = [
  { variant: 'regular', weight: '400', style: 'normal' },
  { variant: 'bold', weight: '700', style: 'normal' },
  { variant: 'italic', weight: '400', style: 'italic' },
  { variant: 'boldItalic', weight: '700', style: 'italic' },
]

let fontsRegistered = false
let fontsRegisteringPromise: Promise<void> | null = null

export async function preloadDocxFonts(): Promise<void> {
  if (typeof document === 'undefined' || document.fonts === undefined) {
    return
  }
  if (fontsRegistered) {
    return
  }
  if (fontsRegisteringPromise !== null) {
    return fontsRegisteringPromise
  }

  fontsRegisteringPromise = (async () => {
    const loads: Promise<FontFace>[] = []
    for (const family of FONT_FAMILIES) {
      for (const descriptor of FONT_VARIANT_DESCRIPTORS) {
        const url = family.files[descriptor.variant]
        for (const name of [family.wordName, family.substituteName]) {
          const face = new FontFace(name, `url(${url})`, {
            weight: descriptor.weight,
            style: descriptor.style,
            display: 'block',
          })
          loads.push(
            face.load().then((loaded) => {
              document.fonts.add(loaded)
              return loaded
            }),
          )
        }
      }
    }
    await Promise.all(loads)
    await document.fonts.ready
    fontsRegistered = true
  })()

  try {
    await fontsRegisteringPromise
  } finally {
    fontsRegisteringPromise = null
  }
}

/**
 * DEFER-4 / DXP-13 — registers a document's own embedded fonts (already
 * de-obfuscated by `loadEmbeddedFonts`) via the same `FontFace` API used for
 * the bundled substitutes, under the font's real Word name so runs that
 * reference it paint with the actual embedded typeface instead of falling
 * back to a metric-substitute or the OS default. Unlike `preloadDocxFonts`
 * (a one-time, app-lifetime registration of the 5 bundled families), this
 * runs per document — callers are responsible for un-registering the
 * returned faces (via `unregisterEmbeddedFonts`) when the document changes,
 * since two different documents can embed two different fonts under the
 * same family name.
 *
 * A face that fails to parse/load (corrupt data, unsupported table format)
 * is skipped individually rather than failing the whole document — the
 * family's other faces, or the bundled-substitute fallback, still work.
 */
export async function registerEmbeddedFonts(
  families: ReadonlyArray<EmbeddedFontFamily>,
): Promise<ReadonlyArray<FontFace>> {
  if (typeof document === 'undefined' || document.fonts === undefined || families.length === 0) {
    return []
  }

  const registered: FontFace[] = []
  for (const family of families) {
    for (const descriptor of FONT_VARIANT_DESCRIPTORS) {
      const data = family.faces[descriptor.variant]
      if (data === undefined) {
        continue
      }

      try {
        const face = new FontFace(family.name, toArrayBuffer(data), {
          weight: descriptor.weight,
          style: descriptor.style,
          display: 'block',
        })
        const loaded = await face.load()
        document.fonts.add(loaded)
        registered.push(loaded)
      } catch {
        // Corrupt/unsupported embedded font data: leave this face
        // unregistered so it falls back to the bundled substitute (or OS
        // default), same as an unresolvable font family today.
      }
    }
  }

  return registered
}

export function unregisterEmbeddedFonts(faces: ReadonlyArray<FontFace>): void {
  if (typeof document === 'undefined' || document.fonts === undefined) {
    return
  }
  for (const face of faces) {
    document.fonts.delete(face)
  }
}

export function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  // `.slice()` always allocates a fresh, non-shared ArrayBuffer (unlike
  // `.buffer.slice()`, whose return type widens to ArrayBufferLike because the
  // source buffer could in principle be a SharedArrayBuffer).
  return bytes.slice().buffer
}

export function createFontResolver(embeddedFonts: ReadonlyArray<EmbeddedFontFamily> = []): FontResolver {
  const embeddedByName = new Map<string, EmbeddedFontFamily>()
  for (const family of embeddedFonts) {
    embeddedByName.set(family.name.trim().toLowerCase(), family)
  }

  const cache = new Map<string, Promise<FontMetrics>>()

  return async (family: string, variant: FontVariant): Promise<FontMetrics> => {
    const embeddedFace = embeddedByName.get(family.trim().toLowerCase())?.faces[variant]
    if (embeddedFace !== undefined) {
      const cacheKey = `embedded:${family.trim().toLowerCase()}@${variant}`
      const existing = cache.get(cacheKey)
      if (existing !== undefined) {
        return existing
      }

      const promise = (async () => {
        try {
          const tables = parseTtf(toArrayBuffer(embeddedFace))
          return wrapWithCanvasAdvance(buildMetrics(tables), family, variant)
        } catch {
          // Corrupt embedded font data: fall back to a neutral synthetic
          // base measured under the requested name, same treatment an
          // unresolvable bundled family gets below.
          return wrapWithCanvasAdvance(DEFAULT_FONT_METRICS, family, variant)
        }
      })()
      cache.set(cacheKey, promise)
      return promise
    }

    const resolved = resolveFontFamily(family)
    if (resolved === null) {
      // Unknown family: still measure via canvas under the requested name so the
      // browser's OS-fallback width is what the paginator sees. We use a
      // synthetic 1000-em base with neutral vertical metrics; vertical metrics
      // for unknown fonts aren't critical (lines won't clip vertically).
      const syntheticBase: FontMetrics = {
        unitsPerEm: 1000,
        ascender: 800,
        descender: -200,
        lineGap: 0,
        xHeight: 500,
        capHeight: 700,
        advanceWidth: () => 500,
        hasGlyph: () => true,
      }
      return wrapWithCanvasAdvance(syntheticBase, family, variant)
    }

    const cacheKey = `${resolved.substituteName}@${variant}`
    const existing = cache.get(cacheKey)
    if (existing !== undefined) {
      return existing
    }

    // Load TTF for accurate vertical metrics, then override per-glyph advances
    // with canvas measurements taken under the requested Word font name. The
    // @font-face stack registered by fontFaces.css maps the Word name to the
    // bundled substitute, so canvas measures what the DOM will actually paint.
    const promise = loadFontMetrics(resolved.wordName, variant)
      .then((base) => wrapWithCanvasAdvance(base, resolved.wordName, variant))
      .catch(() => DEFAULT_FONT_METRICS)
    cache.set(cacheKey, promise)
    return promise
  }
}
