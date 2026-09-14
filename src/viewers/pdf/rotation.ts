/**
 * Pure page-rotation helpers (PDF-12/P10).
 */

import type { PageRotation } from './types'

export function normalizeRotation(value: number): PageRotation {
  const normalized = ((value % 360) + 360) % 360
  switch (normalized) {
    case 90:
      return 90
    case 180:
      return 180
    case 270:
      return 270
    default:
      return 0
  }
}

export function rotateClockwise(current: PageRotation): PageRotation {
  return normalizeRotation(current + 90)
}

export function rotateCounterClockwise(current: PageRotation): PageRotation {
  return normalizeRotation(current - 90)
}

/**
 * Swaps width/height for a sideways (90°/270°) rotation. pdf.js's own
 * `getViewport({ scale, rotation })` reports its rendered `width`/`height`
 * already swapped this way for sideways rotations — so a fit-width/fit-page
 * scale must be computed against these EFFECTIVE dimensions, not the page's
 * raw (unrotated) geometry, or a rotated page in fit mode renders at the
 * wrong scale (it fits against the dimension the rotation just replaced).
 */
export function applyRotationToDimensions(
  width: number,
  height: number,
  rotation: PageRotation,
): { readonly width: number; readonly height: number } {
  return rotation === 90 || rotation === 270 ? { width: height, height: width } : { width, height }
}

/**
 * Combines a page's own intrinsic `/Rotate` entry (`PDFPageProxy.rotate`)
 * with this viewer's additional user-applied rotation state into the TOTAL
 * rotation pdf.js should render at.
 *
 * pdf.js's `page.getViewport({ rotation })` only falls back to the page's own
 * `/Rotate` value when the caller OMITS `rotation` entirely (its default
 * parameter is `rotation = this.rotate`) — an explicitly-passed `rotation`
 * REPLACES that default rather than adding to it. Since this viewer must
 * always pass an explicit value (to reflect the Rotate-button state), every
 * `getViewport`/render call has to combine the two itself, or a page's own
 * baked-in rotation (common in scanned/camera-captured PDFs and any
 * individually-rotated page) is silently discarded until the user manually
 * rotates it back with the toolbar button.
 */
export function combineRotation(intrinsicRotation: number, additionalRotation: PageRotation): PageRotation {
  return normalizeRotation(intrinsicRotation + additionalRotation)
}
