/** USR-16 — the presenter view must own focus, or its Escape-to-exit never fires. */
import { useState } from 'react'
import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { PresenterView } from '../PresenterView'
import type { SlideData } from '../SlideDeck.types'

const slides: ReadonlyArray<SlideData> = [
  { id: 'slide-1', index: 0, width: 1280, height: 720, shapes: [], notes: 'Say hello' },
  { id: 'slide-2', index: 1, width: 1280, height: 720, shapes: [] },
]

// A11Y-2 — PresenterView is only mounted while presenting (SlideDeck.tsx
// renders it conditionally on `presenterOpen`), so exercising the
// mount/unmount focus-restore behavior needs a harness that owns that state,
// mirroring SlideDeck's own "Present" button -> `presenterOpen` wiring.
function Harness() {
  const [open, setOpen] = useState(false)
  return (
    <div>
      <button onClick={() => setOpen(true)}>Present</button>
      {open && <PresenterView slides={slides} activeIndex={0} onSelect={() => {}} onExit={() => setOpen(false)} />}
    </div>
  )
}

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

  it('restores focus to the button that opened it once it exits (A11Y-2)', () => {
    render(<Harness />)
    const trigger = screen.getByRole('button', { name: 'Present' })
    trigger.focus()
    fireEvent.click(trigger)

    const dialog = screen.getByRole('dialog', { name: 'Presenter view' })
    expect(document.activeElement).toBe(dialog)

    fireEvent.click(screen.getByRole('button', { name: 'Exit presenter view' }))

    expect(screen.queryByRole('dialog', { name: 'Presenter view' })).not.toBeInTheDocument()
    expect(trigger).toHaveFocus()
  })
})
