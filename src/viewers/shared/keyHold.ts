/**
 * F6b — see SpreadsheetDataEditor.tsx's module header. Kept separate from that
 * component module (fast refresh) and free of glide-data-grid, so it can be
 * unit-tested on its own.
 */

/** Enough of a keydown to replay it: glide-data-grid reads `key`, `code`, `keyCode` and the modifier flags. */
type HeldKey = Pick<globalThis.KeyboardEvent, 'key' | 'code' | 'keyCode' | 'shiftKey' | 'location'>

const MODIFIER_KEYS = new Set(['Shift', 'Control', 'Alt', 'Meta', 'AltGraph', 'CapsLock'])

/** If glide-data-grid's own deferred focus never lands (a stalled animation frame), focus the grid ourselves after this long. */
export const KEY_HOLD_FALLBACK_MS = 250

/** Our own replays, which must pass straight through rather than be held again. */
const replayed = new WeakSet<Event>()

function replayKey(key: HeldKey): void {
  const target = document.activeElement ?? document.body
  if (target instanceof HTMLTextAreaElement && key.key.length === 1) {
    // The edit overlay mounted mid-replay, and a synthetic keydown never
    // inserts text: insert it the way typing would (caret, input event).
    document.execCommand('insertText', false, key.key)
    return
  }
  const event = new window.KeyboardEvent('keydown', { ...key, bubbles: true, cancelable: true, composed: true })
  replayed.add(event)
  target.dispatchEvent(event)
}

/**
 * F6b — holds the keystrokes typed between a click on the grid and the grid
 * owning DOM focus, then replays them in order — see module header. Replay
 * is one key per task, so each key's state update (an arrow moving the
 * selection, the first character opening the overlay) commits before the
 * next key reads it; keys typed meanwhile queue behind the held ones.
 * Ctrl/Alt/Meta chords are never held: they are app shortcuts (and Electron
 * menu accelerators would not fire for a replayed event).
 */
export class KeyHold {
  private keys: HeldKey[] = []
  private grid: HTMLElement | null = null
  private focusAtStart: Element | null = null
  private draining = false
  private fallback: ReturnType<typeof setTimeout> | undefined

  start(grid: HTMLElement): void {
    if (this.grid !== null) return
    this.grid = grid
    this.focusAtStart = document.activeElement
    window.addEventListener('keydown', this.hold, true)
    this.fallback = setTimeout(this.release, KEY_HOLD_FALLBACK_MS)
  }

  /** Called when focus enters the grid (glide's own deferred focus), or by the fallback timer. */
  readonly release = (): void => {
    const grid = this.grid
    if (grid === null || this.draining) return
    clearTimeout(this.fallback)
    // Only take focus if nothing else has (the user may have clicked elsewhere since).
    if (document.activeElement === this.focusAtStart && !grid.contains(document.activeElement)) {
      grid.querySelector<HTMLCanvasElement>('canvas[data-testid="data-grid-canvas"]')?.focus({ preventScroll: true })
    }
    this.draining = true
    this.drain()
  }

  readonly stop = (): void => {
    clearTimeout(this.fallback)
    window.removeEventListener('keydown', this.hold, true)
    this.keys = []
    this.grid = null
    this.focusAtStart = null
    this.draining = false
  }

  private readonly hold = (event: globalThis.KeyboardEvent): void => {
    if (replayed.has(event) || event.ctrlKey || event.altKey || event.metaKey || MODIFIER_KEYS.has(event.key)) return
    event.preventDefault()
    event.stopImmediatePropagation()
    const { key, code, keyCode, shiftKey, location } = event
    this.keys.push({ key, code, keyCode, shiftKey, location })
  }

  private readonly drain = (): void => {
    const next = this.keys.shift()
    if (next === undefined) {
      this.stop()
      return
    }
    replayKey(next)
    setTimeout(this.drain, 0)
  }
}
