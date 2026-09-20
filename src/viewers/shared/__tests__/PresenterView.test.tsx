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

  // A11Y-2 — PresenterView used to only capture/restore focus on mount/
  // unmount, with no Tab handling at all: Tab walked straight out into the
  // document behind it, which was worse than it sounds because
  // `requestFullscreen` (see the effect above) can be silently refused.
  it('traps Tab within its own controls, skipping the disabled Previous button on the first slide', () => {
    render(<PresenterView slides={slides} activeIndex={0} onSelect={vi.fn()} onExit={vi.fn()} />)

    const prevButton = screen.getByRole('button', { name: 'Previous slide' })
    const nextButton = screen.getByRole('button', { name: 'Next slide' })
    const exitButton = screen.getByRole('button', { name: 'Exit presenter view' })
    expect(prevButton).toBeDisabled()

    // Next is the first real (enabled) focusable — Previous is excluded.
    // Shift+Tab from it must wrap to Exit (the last), not to the disabled
    // Previous button and not out of the view entirely.
    nextButton.focus()
    fireEvent.keyDown(window, { key: 'Tab', shiftKey: true })
    expect(exitButton).toHaveFocus()

    // And Tab from Exit (the last) must wrap back to Next (the first).
    fireEvent.keyDown(window, { key: 'Tab' })
    expect(nextButton).toHaveFocus()
  })

  it('traps Tab even though jsdom has no Fullscreen API (equivalent to requestFullscreen being refused)', () => {
    // jsdom implements no Fullscreen API at all, so `container.requestFullscreen`
    // is undefined here — the same code path a real refusal takes
    // (`.catch(() => undefined)` swallows it either way). If Tab containment
    // depended on fullscreen having actually been granted, this would fail.
    render(<PresenterView slides={slides} activeIndex={0} onSelect={vi.fn()} onExit={vi.fn()} />)
    const nextButton = screen.getByRole('button', { name: 'Next slide' })
    const exitButton = screen.getByRole('button', { name: 'Exit presenter view' })

    exitButton.focus()
    fireEvent.keyDown(window, { key: 'Tab' })

    expect(nextButton).toHaveFocus()
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
