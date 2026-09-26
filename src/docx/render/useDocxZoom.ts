/**
 * The DOCX viewer's zoom level (D24/DXL-19).
 *
 * Extracted from `src/viewers/DocxViewer.tsx` as part of breaking up that file's
 * one very large component — not because zoom is complicated, but because it was
 * one more `useState`, three more `useCallback`s and three module constants in a
 * file that had far too many of each, and it depends on nothing else in there.
 *
 * `PageView` applies the value as a CSS transform, so this is presentation only:
 * pagination measures in points and is deliberately unaffected by zoom.
 */
import { useCallback, useState } from 'react'

export const MIN_ZOOM = 0.25
export const MAX_ZOOM = 3
const ZOOM_STEP = 0.1

function clampZoom(zoom: number): number {
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom))
}

export type DocxZoom = {
  /** 1 = 100%. Clamped to [`MIN_ZOOM`, `MAX_ZOOM`]. */
  readonly zoom: number
  readonly handleZoomIn: () => void
  readonly handleZoomOut: () => void
  readonly handleZoomReset: () => void
}

export function useDocxZoom(): DocxZoom {
  const [zoom, setZoom] = useState(1)

  const handleZoomIn = useCallback(() => {
    setZoom((current) => clampZoom(current + ZOOM_STEP))
  }, [])
  const handleZoomOut = useCallback(() => {
    setZoom((current) => clampZoom(current - ZOOM_STEP))
  }, [])
  const handleZoomReset = useCallback(() => {
    setZoom(1)
  }, [])

  return { zoom, handleZoomIn, handleZoomOut, handleZoomReset }
}
