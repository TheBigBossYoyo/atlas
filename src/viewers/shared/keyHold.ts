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

/**
 * The element that actually receives grid keystrokes.
 *
 * F6c — exported because "does the grid have focus?" must be asked about THIS
 * element and not about the grid container. glide-data-grid renders an
 * accessibility table inside the container, and after an edit commits it leaves
 * focus on one of that table's `<td>`s. A `<td>` is inside the grid but receives
 * no useful keystrokes, so treating "container contains focus" as "grid has
 * focus" silently drops everything typed next. Both this module and
 * `SpreadsheetDataEditor`'s mouse-down hook must agree on the answer.
 */
export const GRID_CANVAS_SELECTOR = 'canvas[data-testid="data-grid-canvas"]'

/** The canvas within `grid`, or null before glide has rendered it. */
export function gridCanvas(grid: Element): HTMLCanvasElement | null {
  return grid.querySelector<HTMLCanvasElement>(GRID_CANVAS_SELECTOR)
}

/**
 * glide's cell-edit overlay lives in `#portal`, OUTSIDE the grid container.
 *
 * MATRIX-FLAKE-1 — "an overlay exists" and "the overlay has focus" are different
 * questions, and conflating them cost a silent data-loss bug. See `drain`.
 */
export const OVERLAY_TEXTAREA_SELECTOR = '#portal textarea'

export function overlayTextarea(): HTMLTextAreaElement | null {
  return document.querySelector<HTMLTextAreaElement>(OVERLAY_TEXTAREA_SELECTOR)
}

/** How often a held shortcut re-checks whether the in-flight edit has committed. */
const IN_FLIGHT_RECHECK_MS = 16

/**
 * How long `drain` waits for a mounted overlay to accept focus before replaying
 * into it anyway.
 *
 * Bounded on purpose: if focus can never land there, dispatching the key is still
 * better than holding it forever, which would lose it just as surely and with no
 * trace. A second is far longer than the mount-then-focus gap this exists for.
 */
const OVERLAY_FOCUS_MAX_WAIT_MS = 1000

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
  /** When the current wait for an unfocused overlay began; null when not waiting. */
  private overlayWaitStartedAt: number | null = null

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
      this.overlayWaitStartedAt = null
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
    const canvas = gridCanvas(grid)
    // Take focus unless something outside the grid has it — the user may have
    // clicked away, or an edit overlay (which lives in `#portal`, outside the
    // grid) may have opened and must keep its own focus.
    //
    // F6c — this used to require `!grid.contains(document.activeElement)`, i.e. it
    // did nothing whenever focus was anywhere inside the grid. After an edit
    // commits, focus sits on a `<td>` of glide's accessibility table, which IS
    // inside the grid and is NOT the canvas — so the canvas was never focused and
    // every replayed key went to that `<td>` and did nothing.
    const active = document.activeElement
    const heldOutsideGrid = active !== this.focusAtStart && !grid.contains(active)
    const overlay = overlayTextarea()
    if (overlay !== null) {
      // MATRIX-FLAKE-1 — an edit overlay is already open, so IT owns focus and the
      // canvas must not take it. The old code asked only whether focus sat inside
      // the grid: after a commit it sits on a `<td>` of glide's accessibility
      // table, which is inside the grid, so the canvas was focused out from under
      // a freshly mounted overlay and every key after the first went to that
      // `<td>` and was lost.
      if (active !== overlay) overlay.focus({ preventScroll: true })
    } else if (!heldOutsideGrid && active !== canvas) {
      canvas?.focus({ preventScroll: true })
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
    // MATRIX-FLAKE-1 — see `awaitOverlayFocus`. Replaying before a mounted overlay
    // has focus delivers the key to whatever glide left focused instead, and it is
    // gone.
    if (!this.awaitOverlayFocus()) {
      this.scheduleDrain(IN_FLIGHT_RECHECK_MS)
      return
    }
    this.keys.shift()
    replayKey(next)
    if (this.keys.length > 0) this.scheduleDrain(0)
  }

  /**
   * MATRIX-FLAKE-1 — whether it is safe to replay the next key now.
   *
   * The overlay mounts one task before it takes focus. Replaying the first
   * character to the canvas is what OPENS it, so the keys immediately after that
   * one arrive in the gap: the overlay exists, and `document.activeElement` is
   * still the canvas or a `<td>` of glide's accessibility table. Dispatching into
   * either loses the key silently — measured as an overlay left holding `"t"`
   * after `"two"` was typed, with the Enter gone too, so the edit could never
   * commit and the Ctrl+S behind it never saved.
   *
   * So: if an overlay is mounted, it must own focus first. Asking for focus and
   * re-checking (rather than assuming `focus()` worked) matters because it can
   * fail while the element is still being attached.
   */
  private awaitOverlayFocus(): boolean {
    const overlay = overlayTextarea()
    if (overlay === null || document.activeElement === overlay) {
      this.overlayWaitStartedAt = null
      return true
    }

    overlay.focus({ preventScroll: true })
    if (document.activeElement === overlay) {
      this.overlayWaitStartedAt = null
      return true
    }

    const startedAt = this.overlayWaitStartedAt ?? performance.now()
    this.overlayWaitStartedAt = startedAt
    if (performance.now() - startedAt >= OVERLAY_FOCUS_MAX_WAIT_MS) {
      // Out of patience. Replaying is still better than holding the key forever.
      this.overlayWaitStartedAt = null
      return true
    }
    return false
  }
}
