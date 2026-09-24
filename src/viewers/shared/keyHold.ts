/**
 * F6b — see SpreadsheetDataEditor.tsx's module header. Kept separate from that
 * component module (fast refresh) and free of glide-data-grid, so it can be
 * unit-tested on its own.
 */

/** Enough of a keydown to replay it: glide-data-grid reads `key`, `code`, `keyCode` and the modifier flags. */
type HeldKey = Pick<
  globalThis.KeyboardEvent,
  'key' | 'code' | 'keyCode' | 'shiftKey' | 'ctrlKey' | 'altKey' | 'metaKey' | 'location'
>

const MODIFIER_KEYS = new Set(['Shift', 'Control', 'Alt', 'Meta', 'AltGraph', 'CapsLock'])

/** If glide-data-grid's own deferred focus never lands (a stalled animation frame), focus the grid ourselves after this long. */
export const KEY_HOLD_FALLBACK_MS = 250

/** How often a held shortcut re-checks whether the in-flight edit has committed. */
const IN_FLIGHT_RECHECK_MS = 16

/** Our own replays, which must pass straight through rather than be held again. */
const replayed = new WeakSet<Event>()

function isChord(key: Pick<HeldKey, 'ctrlKey' | 'altKey' | 'metaKey'>): boolean {
  return key.ctrlKey || key.altKey || key.metaKey
}

function replayKey(key: HeldKey): void {
  const target = document.activeElement ?? document.body
  if (target instanceof HTMLTextAreaElement && key.key.length === 1 && !isChord(key)) {
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
 * F6b — keeps keystrokes in order with the grid edit they follow, so that no
 * key acts before the edit typed ahead of it has reached the grid:
 *
 * - Between a click on the grid and the grid owning DOM focus, every key is
 *   held (glide-data-grid focuses one animation frame late; until then keys
 *   go to `<body>` and are lost). Shortcuts too: Ctrl+S must not save before
 *   the text typed ahead of it.
 * - While an edit is in flight (`editInFlight` — typed, but the overlay has
 *   not mounted and committed yet), shortcuts are held until it commits.
 *   Ordinary keys still go straight to the grid, which buffers them itself.
 * - Once anything is held, everything after it queues behind it.
 *
 * Replay is one key per task, so each key's state update (an arrow moving the
 * selection, the first character opening the overlay) commits before the
 * next key reads it.
 */
export class KeyHold {
  private keys: HeldKey[] = []
  /** Set from a click on the grid until the grid owns focus. */
  private awaitingFocus: HTMLElement | null = null
  private focusAtStart: Element | null = null
  private fallback: ReturnType<typeof setTimeout> | undefined
  private drainTimer: ReturnType<typeof setTimeout> | undefined

  private readonly editInFlight: () => boolean

  constructor(editInFlight: () => boolean = () => false) {
    this.editInFlight = editInFlight
  }

  /** Starts listening; returns the function that stops. One per mounted grid. */
  attach(): () => void {
    window.addEventListener('keydown', this.hold, true)
    return () => {
      window.removeEventListener('keydown', this.hold, true)
      clearTimeout(this.fallback)
      clearTimeout(this.drainTimer)
      this.drainTimer = undefined
      this.keys = []
      this.awaitingFocus = null
      this.focusAtStart = null
    }
  }

  /** A click on the grid: hold keys until the grid owns focus. */
  start(grid: HTMLElement): void {
    if (this.awaitingFocus !== null) return
    this.awaitingFocus = grid
    this.focusAtStart = document.activeElement
    this.fallback = setTimeout(this.release, KEY_HOLD_FALLBACK_MS)
  }

  /** Called when focus enters the grid (glide's own deferred focus), or by the fallback timer. */
  readonly release = (): void => {
    const grid = this.awaitingFocus
    if (grid === null) return
    clearTimeout(this.fallback)
    // Only take focus if nothing else has (the user may have clicked elsewhere since).
    if (document.activeElement === this.focusAtStart && !grid.contains(document.activeElement)) {
      grid.querySelector<HTMLCanvasElement>('canvas[data-testid="data-grid-canvas"]')?.focus({ preventScroll: true })
    }
    this.awaitingFocus = null
    this.focusAtStart = null
    this.scheduleDrain(0)
  }

  private readonly hold = (event: globalThis.KeyboardEvent): void => {
    if (replayed.has(event) || MODIFIER_KEYS.has(event.key)) return
    const mustHold =
      this.awaitingFocus !== null || this.keys.length > 0 || (isChord(event) && this.editInFlight())
    if (!mustHold) return
    event.preventDefault()
    event.stopImmediatePropagation()
    const { key, code, keyCode, shiftKey, ctrlKey, altKey, metaKey, location } = event
    this.keys.push({ key, code, keyCode, shiftKey, ctrlKey, altKey, metaKey, location })
    if (this.awaitingFocus === null) this.scheduleDrain(0)
  }

  private scheduleDrain(delay: number): void {
    if (this.drainTimer !== undefined) return
    this.drainTimer = setTimeout(this.drain, delay)
  }

  private readonly drain = (): void => {
    this.drainTimer = undefined
    const next = this.keys[0]
    if (next === undefined || this.awaitingFocus !== null) return
    if (isChord(next) && this.editInFlight()) {
      this.scheduleDrain(IN_FLIGHT_RECHECK_MS)
      return
    }
    this.keys.shift()
    replayKey(next)
    if (this.keys.length > 0) this.scheduleDrain(0)
  }
}
