/** USR-16 — the presenter view must own focus, or its Escape-to-exit never fires. */
import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { PresenterView } from '../PresenterView'
import type { SlideData } from '../SlideDeck.types'

const slides: ReadonlyArray<SlideData> = [
  { id: 'slide-1', index: 0, width: 1280, height: 720, shapes: [], notes: 'Say hello' },
  { id: 'slide-2', index: 1, width: 1280, height: 720, shapes: [] },
]

describe('PresenterView', () => {
  it('focuses itself even when fullscreen is refused, so Escape exits', () => {
    // jsdom has no Fullscreen API; a real browser can also refuse the request.
    const onExit = vi.fn()
    render(<PresenterView slides={slides} activeIndex={0} onSelect={vi.fn()} onExit={onExit} />)

    const dialog = screen.getByRole('dialog', { name: 'Presenter view' })
    expect(document.activeElement).toBe(dialog)

    fireEvent.keyDown(dialog, { key: 'Escape' })
    expect(onExit).toHaveBeenCalled()
  })

  it('shows the next slide, the notes and the slide counter', () => {
    render(<PresenterView slides={slides} activeIndex={0} onSelect={vi.fn()} onExit={vi.fn()} />)
    expect(screen.getByText('Next')).toBeInTheDocument()
    expect(screen.getByText('Say hello')).toBeInTheDocument()
    expect(screen.getByText('Slide 1 / 2')).toBeInTheDocument()
  })
})
