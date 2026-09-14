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
