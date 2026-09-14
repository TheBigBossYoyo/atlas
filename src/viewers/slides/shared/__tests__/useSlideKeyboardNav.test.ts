import { renderHook } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { useSlideKeyboardNav } from '../useSlideKeyboardNav'

function pressKey(key: string, target: EventTarget = window): void {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  Object.defineProperty(event, 'target', { value: target })
  window.dispatchEvent(event)
}

describe('useSlideKeyboardNav', () => {
  it('S15 — ArrowRight/ArrowDown/Space advance to the next slide, clamped at the last', () => {
    const setActiveIndex = vi.fn()
    renderHook(() => {
      useSlideKeyboardNav(3, setActiveIndex)
    })

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
    renderHook(() => {
      useSlideKeyboardNav(3, setActiveIndex)
    })

    for (const key of ['ArrowLeft', 'ArrowUp', 'PageUp']) {
      setActiveIndex.mockClear()
      pressKey(key)
      expect(setActiveIndex.mock.calls[0][0](0)).toBe(0)
      expect(setActiveIndex.mock.calls[0][0](2)).toBe(1)
    }
  })

  it('S15 — Home/End jump to the first/last slide', () => {
    const setActiveIndex = vi.fn()
    renderHook(() => {
      useSlideKeyboardNav(5, setActiveIndex)
    })

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
    renderHook(() => {
      useSlideKeyboardNav(3, setActiveIndex)
    })

    pressKey('ArrowRight', input)

    expect(setActiveIndex).not.toHaveBeenCalled()
    document.body.removeChild(input)
  })

  it('does nothing when there are no slides', () => {
    const setActiveIndex = vi.fn()
    renderHook(() => {
      useSlideKeyboardNav(0, setActiveIndex)
    })

    pressKey('ArrowRight')

    expect(setActiveIndex).not.toHaveBeenCalled()
  })
})
