/**
 * F6b — `KeyHold` holds the keystrokes typed between a click on the grid and
 * the grid owning focus, and replays them once it does (see
 * SpreadsheetDataEditor.tsx's module header). The real-app proof is
 * tests/e2e/spreadsheet-keystroke-seed.spec.ts (renderer CPU-throttled); these
 * pin the ordering, pass-through and fallback rules that spec can't isolate.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { KEY_HOLD_FALLBACK_MS, KeyHold } from '../keyHold'

let grid: HTMLDivElement
let canvas: HTMLCanvasElement
let received: string[]
let hold: KeyHold

function press(key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init })
  ;(document.activeElement ?? document.body).dispatchEvent(event)
  return event
}

beforeEach(() => {
  vi.useFakeTimers()
  grid = document.createElement('div')
  canvas = document.createElement('canvas')
  canvas.dataset.testid = 'data-grid-canvas'
  canvas.tabIndex = 0
  grid.appendChild(canvas)
  document.body.appendChild(grid)
  received = []
  canvas.addEventListener('keydown', (e) => received.push(e.key))
  hold = new KeyHold()
})

afterEach(() => {
  hold.stop()
  grid.remove()
  vi.useRealTimers()
})

describe('KeyHold (F6b)', () => {
  it('holds keys typed before the grid has focus and replays them to it, in order, one per task', () => {
    hold.start(grid)
    const h = press('H')
    press('ArrowRight')
    press('E')
    expect(h.defaultPrevented).toBe(true)
    expect(received).toEqual([])

    canvas.focus() // glide-data-grid's own deferred focus lands
    hold.release()
    expect(received).toEqual(['H'])
    vi.advanceTimersToNextTimer()
    expect(received).toEqual(['H', 'ArrowRight'])
    vi.advanceTimersToNextTimer()
    expect(received).toEqual(['H', 'ArrowRight', 'E'])
  })

  it('queues keys typed during the replay behind the held ones instead of letting them overtake', () => {
    hold.start(grid)
    press('A')
    press('B')
    canvas.focus()
    hold.release()
    press('C') // arrives at the focused grid while "B" is still waiting its turn
    vi.runAllTimers()
    expect(received).toEqual(['A', 'B', 'C'])
  })

  it('stops holding once the queue is empty', () => {
    hold.start(grid)
    press('A')
    canvas.focus()
    hold.release()
    vi.runAllTimers()
    const later = press('Z')
    expect(later.defaultPrevented).toBe(false)
    expect(received).toEqual(['A', 'Z'])
  })

  it('never holds Ctrl/Alt/Meta chords or bare modifiers — those are app shortcuts', () => {
    const shortcuts: string[] = []
    const onWindow = (e: KeyboardEvent): void => void shortcuts.push(e.key)
    window.addEventListener('keydown', onWindow)
    try {
      hold.start(grid)
      press('s', { ctrlKey: true })
      press('Shift', { shiftKey: true })
      press('Tab', { altKey: true })
      expect(shortcuts).toEqual(['s', 'Shift', 'Tab'])
    } finally {
      window.removeEventListener('keydown', onWindow)
    }
  })

  it('focuses the grid itself if glide-data-grid never does, then replays', () => {
    hold.start(grid)
    press('H')
    vi.advanceTimersByTime(KEY_HOLD_FALLBACK_MS)
    expect(document.activeElement).toBe(canvas)
    expect(received).toEqual(['H'])
  })

  it('does not steal focus back if the user moved it elsewhere meanwhile', () => {
    const field = document.createElement('input')
    document.body.appendChild(field)
    const inField: string[] = []
    field.addEventListener('keydown', (e) => inField.push(e.key))
    try {
      hold.start(grid)
      press('H')
      field.focus()
      vi.advanceTimersByTime(KEY_HOLD_FALLBACK_MS)
      expect(document.activeElement).toBe(field)
      expect(inField).toEqual(['H'])
    } finally {
      field.remove()
    }
  })

  it('inserts replayed characters as text when the edit overlay mounted mid-replay', () => {
    const textarea = document.createElement('textarea')
    document.body.appendChild(textarea)
    // jsdom has no editing commands; this is the call Chromium performs.
    const execCommand = vi.fn(() => true)
    Object.defineProperty(document, 'execCommand', { value: execCommand, configurable: true })
    try {
      hold.start(grid)
      press('H')
      press('I')
      canvas.focus()
      hold.release()
      expect(received).toEqual(['H'])
      textarea.focus() // the overlay opened by "H" took focus
      vi.runAllTimers()
      expect(execCommand).toHaveBeenCalledWith('insertText', false, 'I')
    } finally {
      textarea.remove()
      Reflect.deleteProperty(document, 'execCommand')
    }
  })
})
