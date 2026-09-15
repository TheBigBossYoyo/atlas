import { fireEvent, render } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { ShortcutManagerProvider } from '../../../../hooks/ShortcutManagerProvider'
import { useSlideKeyboardNav } from '../useSlideKeyboardNav'

function Probe({ slideCount, setActiveIndex }: { slideCount: number; setActiveIndex: (updater: (i: number) => number) => void }) {
  useSlideKeyboardNav(slideCount, setActiveIndex)
  return null
}

function renderNav(slideCount: number, setActiveIndex: (updater: (i: number) => number) => void) {
  return render(
    <ShortcutManagerProvider>
      <Probe slideCount={slideCount} setActiveIndex={setActiveIndex} />
    </ShortcutManagerProvider>,
  )
}

function pressKey(key: string, target: Document | Element | Window = window): void {
  fireEvent.keyDown(target, { key })
}

describe('useSlideKeyboardNav', () => {
  it('S15 — ArrowRight/ArrowDown/Space advance to the next slide, clamped at the last', () => {
    const setActiveIndex = vi.fn()
    renderNav(3, setActiveIndex)

    for (const key of ['ArrowRight', 'ArrowDown', ' ']) {
      setActiveIndex.mockClear()
      pressKey(key)
      expect(setActiveIndex).toHaveBeenCalledTimes(1)
      expect(setActiveIndex.mock.calls[0][0](2)).toBe(2) // clamped: already at last index
      expect(setActiveIndex.mock.calls[0][0](0)).toBe(1)
    }
  })

  it('S15 — ArrowLeft/ArrowUp/PageUp go to the previous slide, clamped at zero', () => {
    const setActiveIndex = vi.fn()
    renderNav(3, setActiveIndex)

    for (const key of ['ArrowLeft', 'ArrowUp', 'PageUp']) {
      setActiveIndex.mockClear()
      pressKey(key)
      expect(setActiveIndex.mock.calls[0][0](0)).toBe(0)
      expect(setActiveIndex.mock.calls[0][0](2)).toBe(1)
    }
  })

  it('S15 — Home/End jump to the first/last slide', () => {
    const setActiveIndex = vi.fn()
    renderNav(5, setActiveIndex)

    pressKey('Home')
    expect(setActiveIndex.mock.calls[0][0](3)).toBe(0)

    setActiveIndex.mockClear()
    pressKey('End')
    expect(setActiveIndex.mock.calls[0][0](0)).toBe(4)
  })

  it('ignores keystrokes while an editable control has focus', () => {
    const input = document.createElement('input')
    document.body.appendChild(input)
    const setActiveIndex = vi.fn()
    renderNav(3, setActiveIndex)

    pressKey('ArrowRight', input)

    expect(setActiveIndex).not.toHaveBeenCalled()
    document.body.removeChild(input)
  })

  it('does nothing when there are no slides', () => {
    const setActiveIndex = vi.fn()
    renderNav(0, setActiveIndex)

    pressKey('ArrowRight')

    expect(setActiveIndex).not.toHaveBeenCalled()
  })

  it('lets Space activate a focused thumbnail-rail button instead of advancing the slide (shell-polish fix)', () => {
    const button = document.createElement('button')
    document.body.appendChild(button)
    const onClick = vi.fn()
    button.addEventListener('click', onClick)
    const setActiveIndex = vi.fn()
    renderNav(3, setActiveIndex)

    pressKey(' ', button)

    // The hook must not claim the event (and must not preventDefault it) —
    // it neither advances the slide nor blocks the button's own handling.
    expect(setActiveIndex).not.toHaveBeenCalled()
    document.body.removeChild(button)
  })

  it('still advances the slide on Space when focus is not on a button', () => {
    const div = document.createElement('div')
    document.body.appendChild(div)
    const setActiveIndex = vi.fn()
    renderNav(3, setActiveIndex)

    pressKey(' ', div)

    expect(setActiveIndex).toHaveBeenCalledTimes(1)
    document.body.removeChild(div)
  })
})
