/** PDF find bar — closing behaviour (the bar used to stay open, swallowing the next shortcut). */
import { useState } from 'react'
import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { PdfFindBar } from '../PdfFindBar'

function renderBar(overrides: Partial<React.ComponentProps<typeof PdfFindBar>> = {}) {
  const onClose = vi.fn()
  render(
    <PdfFindBar
      isOpen
      query=""
      matchCount={0}
      currentMatchIndex={0}
      isIndexing={false}
      indexedPageCount={0}
      totalPageCount={3}
      onQueryChange={vi.fn()}
      onNext={vi.fn()}
      onPrev={vi.fn()}
      onClose={onClose}
      {...overrides}
    />,
  )
  return { onClose }
}

describe('PdfFindBar', () => {
  it('closes on Escape pressed in its input', () => {
    const { onClose } = renderBar()
    fireEvent.keyDown(screen.getByLabelText('Find in PDF'), { key: 'Escape' })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('closes on Escape even when the focus is elsewhere', () => {
    // The input only takes focus 50ms after opening, and clicking a page moves
    // focus away again — Escape must still close the bar.
    const { onClose } = renderBar()
    fireEvent.keyDown(document.body, { key: 'Escape' })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('leaves an Escape another handler already consumed alone', () => {
    const { onClose } = renderBar()
    const event = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
    event.preventDefault()
    document.body.dispatchEvent(event)
    expect(onClose).not.toHaveBeenCalled()
  })

  it('does not listen while closed', () => {
    const { onClose } = renderBar({ isOpen: false })
    fireEvent.keyDown(document.body, { key: 'Escape' })
    expect(onClose).not.toHaveBeenCalled()
  })

  it('announces the match count as a polite live region', () => {
    renderBar({ query: 'foo', matchCount: 12, currentMatchIndex: 2 })
    const status = screen.getByText('3 of 12')
    expect(status).toHaveAttribute('aria-live', 'polite')
    expect(status).toHaveAttribute('role', 'status')
  })

  it('announces "no results" the same way', () => {
    renderBar({ query: 'zzz', matchCount: 0 })
    const status = screen.getByText('No results')
    expect(status).toHaveAttribute('aria-live', 'polite')
  })

  // A11Y-4 — the find bar never contained Tab at all: it leaked straight
  // into the PDF page behind it. `matchCount: 0` (the default — no query
  // typed yet) disables the Previous/Next match buttons for real, so this
  // exercises disabled-skip with the bar's own actual state rather than an
  // injected probe.
  describe('focus trap (A11Y-4)', () => {
    it('traps Tab within the bar, skipping the disabled Previous/Next match buttons', () => {
      renderBar({ matchCount: 0 })
      const input = screen.getByLabelText('Find in PDF')
      const prevButton = screen.getByTitle('Previous match (Shift+Enter)')
      const nextButton = screen.getByTitle('Next match (Enter)')
      const closeButton = screen.getByLabelText('Close find')
      expect(prevButton).toBeDisabled()
      expect(nextButton).toBeDisabled()

      // Close is the last real (enabled) focusable — Prev/Next are excluded.
      // Tab from it must wrap back to the input (the first), not escape.
      closeButton.focus()
      fireEvent.keyDown(window, { key: 'Tab' })
      expect(input).toHaveFocus()

      // And Shift+Tab from the input (the first) must wrap back to Close.
      fireEvent.keyDown(window, { key: 'Tab', shiftKey: true })
      expect(closeButton).toHaveFocus()
    })

    it('restores focus to whatever triggered find-in-page once the bar closes', () => {
      function Harness() {
        const [open, setOpen] = useState(false)
        return (
          <div>
            <button onClick={() => setOpen(true)}>open find</button>
            {open && (
              <PdfFindBar
                isOpen
                query=""
                matchCount={0}
                currentMatchIndex={0}
                isIndexing={false}
                indexedPageCount={0}
                totalPageCount={3}
                onQueryChange={vi.fn()}
                onNext={vi.fn()}
                onPrev={vi.fn()}
                onClose={() => setOpen(false)}
              />
            )}
          </div>
        )
      }
      render(<Harness />)
      const trigger = screen.getByRole('button', { name: 'open find' })
      trigger.focus()
      fireEvent.click(trigger)

      fireEvent.click(screen.getByLabelText('Close find'))

      expect(trigger).toHaveFocus()
    })

    it('does not trap Tab while closed', () => {
      renderBar({ isOpen: false })
      const event = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
      window.dispatchEvent(event)
      expect(event.defaultPrevented).toBe(false)
    })
  })
})
