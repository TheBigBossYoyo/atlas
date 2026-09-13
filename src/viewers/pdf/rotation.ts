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
