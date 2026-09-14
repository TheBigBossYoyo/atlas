import type { KeyboardEvent } from 'react'
import {
  ChevronDown,
  ChevronDown as ChevronDownSmall,
  ChevronUp,
  LayoutGrid,
  Printer,
  RotateCw,
  Search,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'

import type { ZoomMode } from './types'

const ZOOM_PRESETS = [0.5, 0.75, 1, 1.25, 1.5, 2]

export type PdfToolbarProps = {
  readonly currentPage: number
  readonly pageCount: number
  readonly pageInput: string
  readonly onPageInputChange: (value: string) => void
  readonly onPageInputKeyDown: (event: KeyboardEvent<HTMLInputElement>) => void
  readonly onPageInputBlur: () => void
  readonly onPrevPage: () => void
  readonly onNextPage: () => void
  readonly zoomMode: ZoomMode
  readonly actualZoom: number
  readonly isZoomMenuOpen: boolean
  readonly onToggleZoomMenu: () => void
  readonly onSelectZoom: (mode: ZoomMode) => void
  readonly onZoomIn: () => void
  readonly onZoomOut: () => void
  readonly onRotate: () => void
  readonly isThumbnailRailOpen: boolean
  readonly onToggleThumbnailRail: () => void
  readonly onToggleFind: () => void
  readonly onPrint: () => void
}

export function PdfToolbar({
  currentPage,
  pageCount,
  pageInput,
  onPageInputChange,
  onPageInputKeyDown,
  onPageInputBlur,
  onPrevPage,
  onNextPage,
  zoomMode,
  actualZoom,
  isZoomMenuOpen,
  onToggleZoomMenu,
  onSelectZoom,
  onZoomIn,
  onZoomOut,
  onRotate,
  isThumbnailRailOpen,
  onToggleThumbnailRail,
  onToggleFind,
  onPrint,
}: PdfToolbarProps) {
  return (
    <div className="pdf-viewer__toolbar">
      <div className="pdf-viewer__toolbar-group">
        <button
          className={`pdf-viewer__toolbar-button ${isThumbnailRailOpen ? 'pdf-viewer__toolbar-button--active' : ''}`}
          onClick={onToggleThumbnailRail}
          aria-pressed={isThumbnailRailOpen}
          title="Toggle page thumbnails"
        >
          <LayoutGrid size={18} />
        </button>
      </div>

      <div className="pdf-viewer__toolbar-group">
        <button
          className="pdf-viewer__toolbar-button"
          onClick={onPrevPage}
          disabled={currentPage <= 1}
          title="Previous Page (Up / PageUp)"
        >
          <ChevronUp size={18} />
        </button>
        <div className="pdf-viewer__page-indicator">
          <input
            className="pdf-viewer__page-input"
            value={pageInput}
            onChange={(e) => onPageInputChange(e.target.value)}
            onKeyDown={onPageInputKeyDown}
            onBlur={onPageInputBlur}
            aria-label="Page number"
          />
          <span>/ {pageCount || '?'}</span>
        </div>
        <button
          className="pdf-viewer__toolbar-button"
          onClick={onNextPage}
          disabled={currentPage >= pageCount}
          title="Next Page (Down / PageDown)"
        >
          <ChevronDown size={18} />
        </button>
      </div>

      <div className="pdf-viewer__toolbar-group pdf-viewer__toolbar-group--end">
        <button className="pdf-viewer__toolbar-button" onClick={onToggleFind} title="Find (Ctrl+F)">
          <Search size={18} />
        </button>

        <button className="pdf-viewer__toolbar-button" onClick={onRotate} title="Rotate page">
          <RotateCw size={18} />
        </button>

        <button className="pdf-viewer__toolbar-button" onClick={onPrint} title="Print (Ctrl+P)">
          <Printer size={18} />
        </button>

        <button className="pdf-viewer__toolbar-button" onClick={onZoomOut} title="Zoom Out (Ctrl+-)">
          <ZoomOut size={18} />
        </button>

        <div className="pdf-viewer__zoom-menu-container">
          <button className="pdf-viewer__zoom-button" onClick={onToggleZoomMenu}>
            {Math.round(actualZoom * 100)}%
            <ChevronDownSmall size={14} />
          </button>

          {isZoomMenuOpen && (
            <div className="pdf-viewer__zoom-menu">
              {ZOOM_PRESETS.map((preset) => (
                <button
                  key={preset}
                  className={`pdf-viewer__zoom-menu-item ${zoomMode === preset ? 'pdf-viewer__zoom-menu-item--active' : ''}`}
                  onClick={() => onSelectZoom(preset)}
                >
                  {Math.round(preset * 100)}%
                </button>
              ))}
              <hr className="pdf-viewer__zoom-menu-separator" />
              <button
                className={`pdf-viewer__zoom-menu-item ${zoomMode === 'fit-width' ? 'pdf-viewer__zoom-menu-item--active' : ''}`}
                onClick={() => onSelectZoom('fit-width')}
              >
                Fit Width
              </button>
              <button
                className={`pdf-viewer__zoom-menu-item ${zoomMode === 'fit-page' ? 'pdf-viewer__zoom-menu-item--active' : ''}`}
                onClick={() => onSelectZoom('fit-page')}
              >
                Fit Page
              </button>
            </div>
          )}
        </div>

        <button className="pdf-viewer__toolbar-button" onClick={onZoomIn} title="Zoom In (Ctrl++)">
          <ZoomIn size={18} />
        </button>
      </div>
    </div>
  )
}
