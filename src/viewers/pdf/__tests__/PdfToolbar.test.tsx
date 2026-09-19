/**
 * P4.7 — PdfToolbar coverage sweep. PdfToolbar is a pure presentational
 * component (no pdfjs-dist dependency), so this exercises it directly with
 * plain prop callbacks rather than going through the full PdfViewer —
 * mirrors the "sub-component gets its own focused test file" pattern the
 * plan calls out for the other PDF sub-components.
 */
import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { PdfToolbar, type PdfToolbarProps } from '../PdfToolbar'

function baseProps(overrides: Partial<PdfToolbarProps> = {}): PdfToolbarProps {
  return {
    currentPage: 2,
    pageCount: 5,
    pageInput: '2',
    onPageInputChange: vi.fn(),
    onPageInputKeyDown: vi.fn(),
    onPageInputBlur: vi.fn(),
    onPrevPage: vi.fn(),
    onNextPage: vi.fn(),
    zoomMode: 1,
    actualZoom: 1,
    isZoomMenuOpen: false,
    onToggleZoomMenu: vi.fn(),
    onSelectZoom: vi.fn(),
    onZoomIn: vi.fn(),
    onZoomOut: vi.fn(),
    onRotate: vi.fn(),
    isThumbnailRailOpen: false,
    onToggleThumbnailRail: vi.fn(),
    onToggleFind: vi.fn(),
    onPrint: vi.fn(),
    ...overrides,
  }
}

describe('PdfToolbar', () => {
  it('shows the current page and page count', () => {
    render(<PdfToolbar {...baseProps({ currentPage: 2, pageCount: 5, pageInput: '2' })} />)

    expect(screen.getByLabelText('Page number')).toHaveValue('2')
    expect(screen.getByText('/ 5')).toBeInTheDocument()
  })

  it('shows "?" for the page count while the document is still loading (pageCount 0)', () => {
    render(<PdfToolbar {...baseProps({ pageCount: 0 })} />)

    expect(screen.getByText('/ ?')).toBeInTheDocument()
  })

  it('navigates to the previous/next page via the toolbar buttons', () => {
    const onPrevPage = vi.fn()
    const onNextPage = vi.fn()
    render(<PdfToolbar {...baseProps({ onPrevPage, onNextPage })} />)

    fireEvent.click(screen.getByTitle('Previous Page (Up / PageUp)'))
    fireEvent.click(screen.getByTitle('Next Page (Down / PageDown)'))

    expect(onPrevPage).toHaveBeenCalledTimes(1)
    expect(onNextPage).toHaveBeenCalledTimes(1)
  })

  it('disables "previous page" on the first page and enables "next page"', () => {
    render(<PdfToolbar {...baseProps({ currentPage: 1, pageCount: 5 })} />)

    expect(screen.getByTitle('Previous Page (Up / PageUp)')).toBeDisabled()
    expect(screen.getByTitle('Next Page (Down / PageDown)')).toBeEnabled()
  })

  it('disables "next page" on the last page and enables "previous page"', () => {
    render(<PdfToolbar {...baseProps({ currentPage: 5, pageCount: 5 })} />)

    expect(screen.getByTitle('Next Page (Down / PageDown)')).toBeDisabled()
    expect(screen.getByTitle('Previous Page (Up / PageUp)')).toBeEnabled()
  })

  it('disables both nav buttons while the page count is unknown (0)', () => {
    render(<PdfToolbar {...baseProps({ currentPage: 1, pageCount: 0 })} />)

    expect(screen.getByTitle('Previous Page (Up / PageUp)')).toBeDisabled()
    expect(screen.getByTitle('Next Page (Down / PageDown)')).toBeDisabled()
  })

  it('reports page-number input edits, Enter/other key handling, and blur through the given callbacks', () => {
    const onPageInputChange = vi.fn()
    const onPageInputKeyDown = vi.fn()
    const onPageInputBlur = vi.fn()
    render(
      <PdfToolbar {...baseProps({ onPageInputChange, onPageInputKeyDown, onPageInputBlur })} />,
    )

    const input = screen.getByLabelText('Page number')
    fireEvent.change(input, { target: { value: '3' } })
    expect(onPageInputChange).toHaveBeenCalledWith('3')

    fireEvent.keyDown(input, { key: 'Enter' })
    expect(onPageInputKeyDown).toHaveBeenCalledTimes(1)

    fireEvent.blur(input)
    expect(onPageInputBlur).toHaveBeenCalledTimes(1)
  })

  it('toggles the thumbnail rail and reflects its open state via aria-pressed', () => {
    const onToggleThumbnailRail = vi.fn()
    const { rerender } = render(
      <PdfToolbar {...baseProps({ isThumbnailRailOpen: false, onToggleThumbnailRail })} />,
    )

    const button = screen.getByTitle('Toggle page thumbnails')
    expect(button).toHaveAttribute('aria-pressed', 'false')

    fireEvent.click(button)
    expect(onToggleThumbnailRail).toHaveBeenCalledTimes(1)

    rerender(<PdfToolbar {...baseProps({ isThumbnailRailOpen: true, onToggleThumbnailRail })} />)
    expect(screen.getByTitle('Toggle page thumbnails')).toHaveAttribute('aria-pressed', 'true')
  })

  it('opens find, rotates, and prints via their toolbar buttons', () => {
    const onToggleFind = vi.fn()
    const onRotate = vi.fn()
    const onPrint = vi.fn()
    render(<PdfToolbar {...baseProps({ onToggleFind, onRotate, onPrint })} />)

    fireEvent.click(screen.getByTitle('Find (Ctrl+F)'))
    fireEvent.click(screen.getByTitle('Rotate page'))
    fireEvent.click(screen.getByTitle('Print (Ctrl+P)'))

    expect(onToggleFind).toHaveBeenCalledTimes(1)
    expect(onRotate).toHaveBeenCalledTimes(1)
    expect(onPrint).toHaveBeenCalledTimes(1)
  })

  it('zooms in and out via the toolbar buttons', () => {
    const onZoomIn = vi.fn()
    const onZoomOut = vi.fn()
    render(<PdfToolbar {...baseProps({ onZoomIn, onZoomOut })} />)

    fireEvent.click(screen.getByTitle('Zoom In (Ctrl++)'))
    fireEvent.click(screen.getByTitle('Zoom Out (Ctrl+-)'))

    expect(onZoomIn).toHaveBeenCalledTimes(1)
    expect(onZoomOut).toHaveBeenCalledTimes(1)
  })

  it('shows the current zoom percentage on the zoom button and toggles the zoom menu', () => {
    const onToggleZoomMenu = vi.fn()
    render(<PdfToolbar {...baseProps({ actualZoom: 1.5, isZoomMenuOpen: false, onToggleZoomMenu })} />)

    const zoomButton = screen.getByText('150%').closest('button')
    expect(zoomButton).not.toBeNull()
    // Menu is closed: none of the preset menu items should be in the DOM.
    expect(screen.queryByText('Fit Width')).not.toBeInTheDocument()

    fireEvent.click(zoomButton!)
    expect(onToggleZoomMenu).toHaveBeenCalledTimes(1)
  })

  it('lists every zoom preset plus Fit Width/Fit Page when the menu is open, and selects one on click', () => {
    const onSelectZoom = vi.fn()
    // `actualZoom` deliberately doesn't match any preset label below so the
    // zoom button's own "NN%" text can't collide with a menu item's.
    render(
      <PdfToolbar {...baseProps({ isZoomMenuOpen: true, zoomMode: 1, actualZoom: 0.6, onSelectZoom })} />,
    )

    for (const label of ['50%', '75%', '100%', '125%', '150%', '200%', 'Fit Width', 'Fit Page']) {
      expect(screen.getByText(label)).toBeInTheDocument()
    }

    fireEvent.click(screen.getByText('125%'))
    expect(onSelectZoom).toHaveBeenCalledWith(1.25)

    fireEvent.click(screen.getByText('Fit Width'))
    expect(onSelectZoom).toHaveBeenCalledWith('fit-width')

    fireEvent.click(screen.getByText('Fit Page'))
    expect(onSelectZoom).toHaveBeenCalledWith('fit-page')
  })

  it('marks the active zoom preset in the open menu', () => {
    const { rerender } = render(<PdfToolbar {...baseProps({ isZoomMenuOpen: true, zoomMode: 1 })} />)

    expect(screen.getByText('100%', { selector: '.pdf-viewer__zoom-menu-item' })).toHaveClass(
      'pdf-viewer__zoom-menu-item--active',
    )
    expect(screen.getByText('150%', { selector: '.pdf-viewer__zoom-menu-item' })).not.toHaveClass(
      'pdf-viewer__zoom-menu-item--active',
    )

    rerender(<PdfToolbar {...baseProps({ isZoomMenuOpen: true, zoomMode: 'fit-width' })} />)
    expect(screen.getByText('Fit Width')).toHaveClass('pdf-viewer__zoom-menu-item--active')
  })
})
