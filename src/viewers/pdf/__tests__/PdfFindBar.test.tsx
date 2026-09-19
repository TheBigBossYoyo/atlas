/** PDF find bar — closing behaviour (the bar used to stay open, swallowing the next shortcut). */
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
})
