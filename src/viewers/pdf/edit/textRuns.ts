/**
 * Pixel and selection helpers for the edit layer (PDF-EDIT-1/2): reading the
 * paper colour behind an area from the rendered canvas, and merging a text
 * selection's rectangles into one per line for a highlight.
 *
 * Pure: they only read the values they are given.
 */

import type { RgbColor } from './types'

// ---- Colour sampling ------------------------------------------------------------------

export type PixelBlock = {
  readonly data: Uint8ClampedArray | ReadonlyArray<number>
  readonly width: number
  readonly height: number
}

const WHITE: RgbColor = { r: 1, g: 1, b: 1 }
const BLACK: RgbColor = { r: 0, g: 0, b: 0 }

function colorDistance(a: readonly [number, number, number], b: readonly [number, number, number]): number {
  return Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1]) + Math.abs(a[2] - b[2])
}

function toRgb([r, g, b]: readonly [number, number, number]): RgbColor {
  return { r: r / 255, g: g / 255, b: b / 255 }
}

/**
 * Reads a run's background and ink colours from the pixels of its box.
 *
 * Background: the most common colour along the block's border, where glyphs
 * rarely reach. Ink: the pixel furthest from that background — the core of a
 * glyph stroke, past the anti-aliased edge. Fully transparent pixels (an
 * unpainted canvas) count as white, which is what a reader shows there.
 */
export function sampleRunColors(block: PixelBlock): { readonly background: RgbColor; readonly ink: RgbColor } {
  const { data, width, height } = block
  if (width <= 0 || height <= 0) return { background: WHITE, ink: BLACK }

  const pixel = (x: number, y: number): [number, number, number] => {
    const offset = (y * width + x) * 4
    const alpha = data[offset + 3] / 255
    // Composite over white.
    return [0, 1, 2].map((channel) => Math.round(data[offset + channel] * alpha + 255 * (1 - alpha))) as [number, number, number]
  }

  const counts = new Map<string, { color: [number, number, number]; count: number }>()
  const tally = (x: number, y: number) => {
    const color = pixel(x, y)
    // Quantised so near-identical anti-aliased neighbours count together.
    const key = color.map((channel) => channel >> 3).join(',')
    const entry = counts.get(key)
    if (entry) entry.count += 1
    else counts.set(key, { color, count: 1 })
  }
  for (let x = 0; x < width; x += 1) {
    tally(x, 0)
    if (height > 1) tally(x, height - 1)
  }
  for (let y = 1; y < height - 1; y += 1) {
    tally(0, y)
    if (width > 1) tally(width - 1, y)
  }
  let background: [number, number, number] = [255, 255, 255]
  let bestCount = -1
  counts.forEach(({ color, count }) => {
    if (count > bestCount) {
      background = color
      bestCount = count
    }
  })

  let ink: [number, number, number] = background
  let inkDistance = -1
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const color = pixel(x, y)
      const distance = colorDistance(color, background)
      if (distance > inkDistance) {
        ink = color
        inkDistance = distance
      }
    }
  }
  // No glyph pixels found (blank area, or text in the background colour):
  // fall back to whichever of black/white contrasts with the background.
  if (inkDistance < 48) {
    const luminance = 0.299 * background[0] + 0.587 * background[1] + 0.114 * background[2]
    return { background: toRgb(background), ink: luminance > 128 ? BLACK : WHITE }
  }
  return { background: toRgb(background), ink: toRgb(ink) }
}

// ---- Selection rectangles ------------------------------------------------------------------

export type ScreenRect = {
  readonly left: number
  readonly top: number
  readonly right: number
  readonly bottom: number
}

/**
 * Merges the client rects of a text selection into one rect per line run.
 * The text layer yields a rect per text node — often several per line, with
 * small gaps or overlaps between words — and a highlight drawn from those
 * unmerged shows seams and double-darkened overlaps.
 */
export function mergeLineRects(rects: ReadonlyArray<ScreenRect>, gapTolerance = 3): ScreenRect[] {
  const usable = rects
    .filter((rect) => rect.right - rect.left >= 1 && rect.bottom - rect.top >= 1)
    .slice()
    .sort((a, b) => a.top - b.top || a.left - b.left)
  const merged: ScreenRect[] = []
  for (const rect of usable) {
    const height = rect.bottom - rect.top
    const line = merged.find((existing) => {
      const overlap = Math.min(existing.bottom, rect.bottom) - Math.max(existing.top, rect.top)
      const minHeight = Math.min(existing.bottom - existing.top, height)
      const horizontallyNear = rect.left <= existing.right + gapTolerance && rect.right >= existing.left - gapTolerance
      return overlap > minHeight * 0.5 && horizontallyNear
    })
    if (line) {
      merged[merged.indexOf(line)] = {
        left: Math.min(line.left, rect.left),
        top: Math.min(line.top, rect.top),
        right: Math.max(line.right, rect.right),
        bottom: Math.max(line.bottom, rect.bottom),
      }
    } else {
      merged.push(rect)
    }
  }
  return merged
}
