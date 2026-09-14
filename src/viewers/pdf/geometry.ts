/**
 * Pure zoom/scale math (PDF-02/P2, PDF-09/P8).
 *
 * Kept free of any pdfjs-dist or DOM dependency: callers pass plain numbers
 * (a page's own unscaled width/height, the container's viewport size) so
 * this can be unit tested directly and reused per-page instead of once per
 * document.
 */

import type { ZoomMode } from './types'

export const MIN_ZOOM = 0.25
export const MAX_ZOOM = 3
export const DEFAULT_ZOOM = 1
export const ZOOM_STEP = 0.25

/** Padding (px) subtracted from the scrollable viewer's client size before
 * fitting a page to it, so fit-width/fit-page always leave a small margin. */
export const FIT_PADDING_PX = 48

export function clampZoom(value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    return DEFAULT_ZOOM
  }
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value))
}

/**
 * Computes the fit-width or fit-page scale for ONE page's own unscaled
 * dimensions against the given container size (PDF-09: previously this was
 * computed once from page 1 and reused for every page, so a mixed-size
 * document rendered every page after the first at the wrong scale).
 */
export function computeFitScale(
  pageWidth: number,
  pageHeight: number,
  containerWidth: number,
  containerHeight: number,
  mode: 'fit-width' | 'fit-page',
): number {
  if (pageWidth <= 0 || pageHeight <= 0 || containerWidth <= 0) {
    return DEFAULT_ZOOM
  }

  const widthScale = containerWidth / pageWidth

  if (mode === 'fit-width') {
    return clampZoom(widthScale)
  }

  const heightScale = containerHeight > 0 ? containerHeight / pageHeight : widthScale
  return clampZoom(Math.min(widthScale, heightScale))
}

/** Resolves the render scale for one page, dispatching on the zoom mode. */
export function resolveScaleForPage(
  zoomMode: ZoomMode,
  pageWidth: number,
  pageHeight: number,
  containerWidth: number,
  containerHeight: number,
): number {
  if (zoomMode === 'fit-width' || zoomMode === 'fit-page') {
    return computeFitScale(pageWidth, pageHeight, containerWidth, containerHeight, zoomMode)
  }
  return clampZoom(zoomMode)
}

export function zoomIn(current: number): number {
  // Clamp `current` itself first: stepping down from MIN_ZOOM would
  // otherwise produce 0, which clampZoom treats as an invalid input (falling
  // back to 100%) rather than as "clamp to the floor".
  return Math.min(MAX_ZOOM, clampZoom(current) + ZOOM_STEP)
}

export function zoomOut(current: number): number {
  return Math.max(MIN_ZOOM, clampZoom(current) - ZOOM_STEP)
}
