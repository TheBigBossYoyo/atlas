/**
 * F6b — `KeyHold` keeps keystrokes in order with the grid edit they follow
 * (see SpreadsheetDataEditor.tsx's module header). The real-app proof is
 * tests/e2e/spreadsheet-keystroke-seed.spec.ts (renderer CPU-throttled); these
 * pin the ordering, pass-through and fallback rules that spec can't isolate.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { KEY_HOLD_FALLBACK_MS, KeyHold } from '../keyHold'

let grid: HTMLDivElement
let canvas: HTMLCanvasElement
let received: string[]
let inFlight: boolean
let hold: KeyHold
let detach: () => void

function press(key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init })
  ;(document.activeElement ?? document.body).dispatchEvent(event)
  return event
}

/** Records what reaches the window after KeyHold, the way the app's shortcut handler sees it. */
function recordOnWindow(): { keys: string[]; stop: () => void } {
  const keys: string[] = []
  const listener = (e: KeyboardEvent): void => void keys.push((e.ctrlKey ? 'Ctrl+' : '') + e.key)
  window.addEventListener('keydown', listener)
  return { keys, stop: () => window.removeEventListener('keydown', listener) }
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
  inFlight = false
  hold = new KeyHold(() => inFlight)
  detach = hold.attach()
})

afterEach(() => {
  detach()
  grid.remove()
  vi.useRealTimers()
})

describe('KeyHold (F6b) — between a click and the grid owning focus', () => {
  it('holds every key and replays them to the grid, in order, one per task', () => {
    hold.start(grid)
    const h = press('H')
    press('ArrowRight')
    press('E')
    expect(h.defaultPrevented).toBe(true)
    expect(received).toEqual([])

    canvas.focus() // glide-data-grid's own deferred focus lands
    hold.release()
    vi.advanceTimersToNextTimer()
    expect(received).toEqual(['H'])
    vi.advanceTimersToNextTimer()
    expect(received).toEqual(['H', 'ArrowRight'])
    vi.advanceTimersToNextTimer()
    expect(received).toEqual(['H', 'ArrowRight', 'E'])
  })

  it('holds shortcuts too, so Ctrl+S cannot save before the text typed ahead of it', () => {
    const onWindow = recordOnWindow()
    try {
      hold.start(grid)
      press('H')
      press('s', { ctrlKey: true })
      expect(onWindow.keys).toEqual([])
      canvas.focus()
      hold.release()
      vi.runAllTimers()
      expect(onWindow.keys).toEqual(['H', 'Ctrl+s'])
    } finally {
      onWindow.stop()
    }
  })

  it('queues keys typed during the replay behind the held ones instead of letting them overtake', () => {
    hold.start(grid)
    press('A')
    press('B')
    canvas.focus()
    hold.release()
    press('C') // arrives at the focused grid while "A" and "B" are still waiting their turn
    vi.runAllTimers()
    expect(received).toEqual(['A', 'B', 'C'])
  })

  it('lets bare modifier keys through', () => {
    const onWindow = recordOnWindow()
    try {
      hold.start(grid)
      press('Shift', { shiftKey: true })
      expect(onWindow.keys).toEqual(['Shift'])
    } finally {
      onWindow.stop()
    }
  })

  it('focuses the grid itself if glide-data-grid never does, then replays', () => {
    hold.start(grid)
    press('H')
    vi.advanceTimersByTime(KEY_HOLD_FALLBACK_MS)
    expect(document.activeElement).toBe(canvas)
    vi.runAllTimers()
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
      vi.runAllTimers()
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
      vi.advanceTimersToNextTimer()
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

describe('KeyHold (F6b) — while an edit is in flight', () => {
  beforeEach(() => canvas.focus())

  it('holds a shortcut until the edit commits, then replays it', () => {
    const onWindow = recordOnWindow()
    try {
      inFlight = true
      const save = press('s', { ctrlKey: true })
      expect(save.defaultPrevented).toBe(true)
      vi.advanceTimersByTime(200)
      expect(onWindow.keys).toEqual([])

      inFlight = false // the overlay mounted and committed
      vi.advanceTimersByTime(50)
      expect(onWindow.keys).toEqual(['Ctrl+s'])
    } finally {
      onWindow.stop()
    }
  })

  it('lets ordinary keys straight through to the grid, which buffers them itself', () => {
    inFlight = true
    const e = press('E')
    expect(e.defaultPrevented).toBe(false)
    expect(received).toEqual(['E'])
  })

  it('queues ordinary keys typed after a held shortcut behind it', () => {
    inFlight = true
    press('s', { ctrlKey: true })
    press('X')
    expect(received).toEqual([])
    inFlight = false
    vi.runAllTimers()
    expect(received).toEqual(['s', 'X'])
  })

  it('touches nothing when no click is pending and no edit is in flight', () => {
    const save = press('s', { ctrlKey: true })
    const x = press('X')
    expect(save.defaultPrevented).toBe(false)
    expect(x.defaultPrevented).toBe(false)
    expect(received).toEqual(['s', 'X'])
  })

  it('stops listening once detached', () => {
    detach()
    inFlight = true
    const save = press('s', { ctrlKey: true })
    expect(save.defaultPrevented).toBe(false)
    detach = () => undefined
  })
})
