/**
 * glide-data-grid's `DataEditor` as used by SpreadsheetViewer and CsvViewer,
 * loaded lazily by both (this module statically pulls in the grid and its CSS).
 *
 * USR-17 — cell edits typed quickly were silently dropped: the grid's overlay
 * commits Enter/Tab with a value captured from its previous render, so a key
 * typed just before Enter (or the whole edit, when the overlay had not
 * re-rendered yet) never reached `onCellEdited`. The editor below swaps in
 * the same text entry with its own Enter/Tab handler that commits the
 * textarea's live value directly. Escape and click-outside keep the grid's
 * own behavior.
 *
 * F6 — the FIRST character typed into a freshly-selected cell was silently
 * dropped (confirmed against saved `.xlsx` bytes: click a cell, type "HELLO",
 * save, and the cell held "ELLO" — deterministically, not a rare race, at
 * every typing speed from 0ms to 120ms/char). Root cause, confirmed by
 * tracing real keydown/focus/DOM-mutation events in the running app:
 * glide-data-grid's own "type on a selected cell to start editing" feature
 * (`editOnType`, on by default) opens its edit overlay by calling an internal
 * `reselect(bounds, true, key)` on the FIRST qualifying keydown, which seeds
 * the new cell's `data` with that key — this part already works. But mounting
 * that overlay and moving DOM focus into its textarea is not synchronous: our
 * own tracing showed the seeded "H" overlay get rendered (React called the
 * editor component) yet never actually committed/focused for over 100ms,
 * because glide-data-grid still has DOM focus sitting on the grid's
 * accessibility-tree cell element (`<canvas>`'s own fallback content) at that
 * point. A second keystroke arriving before that focus handoff completes —
 * which ordinary typing (well under 120ms/char) always does — lands back on
 * that STILL-focused accessibility element instead of the new textarea, so
 * glide-data-grid treats it as ANOTHER "start editing" trigger: `reselect()`
 * runs again, recomputes the cell's content from scratch, and overwrites the
 * just-seeded "H" with just that second key. Only the LAST such racing
 * keystroke before the overlay finally mounts survives as the seed — for
 * ordinary typing speed that's always the second character, which is exactly
 * the observed "first character missing" bug (and why 0ms/char, where several
 * keystrokes can race before the very first mount, loses more than one).
 *
 * glide-data-grid exposes no public API to open its edit overlay
 * imperatively (only `DataEditorRef.emit()` for delete/fill, nothing for
 * "start editing with this seed"), so the fix can't avoid using its own
 * `editOnType` to open the FIRST keystroke's overlay. Instead, `pendingSeedRef`
 * below buffers every character that arrives for the cell currently in the
 * process of opening (still targeting the grid, not yet the mounted
 * textarea) rather than letting it re-trigger `reselect()`, and
 * `TextCellEditor` drains that buffer into its initial value the one time it
 * mounts for a type-triggered edit — so no matter how many keystrokes race
 * ahead of the overlay's own focus handoff, all of them end up in the
 * committed text, in order, instead of only the last one winning.
 *
 * `PendingSeedContext` (rather than closing over the ref directly while
 * building `provideEditor`, or handing it to a factory function) is a
 * `useRef` value threaded through a stable, module-level `TextCellEditor` —
 * `provideEditor`'s callback must be safe to invoke during render, and both
 * "read a ref during render" and "pass a ref into a function during render"
 * are disallowed (`react-hooks/refs`); routing it through context instead
 * means the ref is only ever read inside `TextCellEditor`'s own effect,
 * never during any parent's render.
 *
 * The mount itself is not fast: tracing a real run showed 200-300ms between
 * the seeding keystroke and the overlay's textarea actually gaining DOM
 * focus (glide-data-grid re-renders the whole grid as part of opening it).
 * That is longer than an entire short word typed at ordinary speed — typing
 * "HELLO" and pressing Enter right after, at 30ms/char, finishes in ~230ms,
 * before the overlay has mounted at all. So Enter/Tab can ALSO arrive while
 * still targeting glide-data-grid's own accessibility element rather than
 * the (not yet mounted) textarea, where glide-data-grid treats them as plain
 * navigation (move the selection down/across) — silently discarding every
 * buffered character, seed included, instead of committing them. `pending`'s
 * `commitOnMount` below records that intent (which movement Enter/Tab asked
 * for) so `TextCellEditor` can finish editing immediately once it mounts,
 * instead of only seeding it and leaving the commit to a keystroke that will
 * now never come (the overlay has focus, but nothing is still "racing" to
 * retrigger it).
 *
 * F6b — the buffer above only helps keystrokes that reach the grid. On a slow
 * or busy CPU the FIRST keystroke after a click did not: glide-data-grid's
 * pointer-down handler calls `preventDefault()` (suppressing the browser's
 * own focus move) and only focuses the grid in a `requestAnimationFrame`
 * (`focus()` in data-editor.js). Until that frame runs, DOM focus is still on
 * `<body>` (or wherever it was), so a key typed in that window goes there and
 * is silently discarded — measured with the renderer CPU-throttled 8x:
 * click, type "HELLO" at 120ms/char, save, and the cell held "ELLO" in every
 * trial, with a keydown trace showing "H" targeting BODY and "E" the canvas.
 *
 * Focusing the grid early is NOT a fix: that frame is also when the click's
 * selection has taken effect, and keys delivered before it act on a grid
 * with no selection yet (tried: click, ArrowRight, ArrowDown x2, type "42"
 * put 42 in A1 — every arrow ignored). Instead `KeyHold` (keyHold.ts) holds
 * the keystrokes typed between a click on the grid and the grid owning focus,
 * then replays them in order once it does, one per task. It only engages when
 * focus is outside the grid and its edit overlay (glide's `#portal`), so
 * typing and clicking inside an already-focused grid are untouched.
 *
 * Shortcuts must keep their place in that order too. At 24x throttling, a
 * 0ms/char "HELLO", Enter, Ctrl+S saved "Name": the save either overtook
 * held keys still waiting to replay, or ran after they replayed but before
 * the overlay mounted and applied `commitOnMount`. The edit landed a moment
 * later (a second save had "HELLO"), but the file saved at the user's
 * Ctrl+S lacked it. So `KeyHold` also holds Ctrl/Alt/Meta chords while a
 * pending seed is in flight, and replays them only once it has committed.
 */
import { createContext, useCallback, useContext, useEffect, useRef, type ChangeEvent, type KeyboardEvent, type MouseEvent } from 'react'
import {
  DataEditor,
  GridCellKind,
  TextCellEntry,
  type DataEditorProps,
  type GridCell,
  type GridKeyEventArgs,
  type ProvideEditorCallback,
  type ProvideEditorComponent,
  type TextCell,
} from '@glideapps/glide-data-grid'
import '@glideapps/glide-data-grid/dist/index.css'

import { KeyHold } from './keyHold'

/** A movement tuple, matching `TextCellEditor`'s own Enter/Tab handler below. */
type Movement = readonly [-1 | 0 | 1, -1 | 0 | 1]

/** One in-flight "typing opened this cell's overlay, but it hasn't mounted yet" buffer — see module header (F6). */
type PendingSeed = {
  readonly row: number
  readonly col: number
  chars: string[]
  /** Set when Enter/Tab arrived before the overlay mounted — see module header. */
  commitOnMount?: Movement
  /** `performance.now()` when the first key opened it — bounds how long it counts as in flight (F6b). */
  readonly startedAt: number
}

/**
 * F6b — last-resort bound: a pending seed older than this no longer counts as
 * an edit in flight, so a seed stranded by some path nobody foresaw can never
 * hold shortcuts like Ctrl+S indefinitely. (The known stranding path — glide
 * declining to open an overlay — is cleared precisely in `handleKeyDown`.)
 * Overlay mounts take 100-300ms normally; at 24x CPU throttling one took
 * 2.45s locally (a 2s bound let a held Ctrl+S overtake it), and on the slower
 * CI runner the 24x test exceeded 5s. Once glide has confirmed it opened an
 * editor the edit really is coming (the only other way to strand it is the
 * grid unmounting, which detaches KeyHold anyway), so this bound is generous.
 */
const PENDING_EDIT_MAX_MS = 30_000

/** Holds the current `SpreadsheetDataEditor` instance's pending-seed ref — see module header on why this goes through context rather than a closure. */
const PendingSeedContext = createContext<{ current: PendingSeed | null } | null>(null)

/** Mirrors glide-data-grid's own `editOnType` key filter (data-editor.js) exactly, so we only ever intervene on keys it would itself treat as "start editing". */
function isEditOnTypeKey(event: GridKeyEventArgs): boolean {
  return !event.metaKey && !event.ctrlKey && event.key.length === 1 && /[ -~]/.test(event.key)
}

/** Same commit-key/movement mapping as `TextCellEditor`'s own post-mount Enter/Tab handler below. */
function commitMovementFor(event: GridKeyEventArgs): Movement | undefined {
  if (event.key === 'Enter' && !event.shiftKey) return [0, 1]
  if (event.key === 'Tab') return [event.shiftKey ? -1 : 1, 0]
  return undefined
}

const TextCellEditor: ProvideEditorComponent<TextCell> = ({
  value,
  onChange,
  onFinishedEditing,
  isHighlighted,
  validatedSelection,
  forceEditMode,
  initialValue,
}) => {
  const pendingRef = useContext(PendingSeedContext)

  // Runs once per mount only (see module header): a type-triggered open
  // (forceEditMode + initialValue, per glide-data-grid's own `reselect`) may
  // have raced ahead keystrokes (and an Enter/Tab requesting an immediate
  // commit) buffered in `pendingRef` while this overlay was still mounting.
  // Apply them, in order, to the seed glide already put in `value.data`, and
  // finish editing right away if a commit was requested. A double-click/F2
  // open (forceEditMode false, no initialValue) never applies the buffer,
  // but ANY overlay mounting — including that one — still clears
  // `pendingRef`: once editing has genuinely started (by whatever path), a
  // leftover buffer from an earlier attempt that never made it to mount (the
  // user clicked away, say) is stale and must not silently attach itself to
  // some unrelated later edit of the same cell.
  const appliedPendingSeed = useRef(false)
  useEffect(() => {
    if (appliedPendingSeed.current) return
    appliedPendingSeed.current = true
    if (pendingRef === null) return
    const pending = pendingRef.current
    pendingRef.current = null
    if (pending && forceEditMode && initialValue !== undefined) {
      const seeded = pending.chars.length > 0 ? { ...value, data: value.data + pending.chars.join('') } : value
      if (pending.commitOnMount !== undefined) {
        onFinishedEditing(seeded, pending.commitOnMount)
      } else if (seeded !== value) {
        onChange(seeded)
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const handleKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>): void => {
    const isCommitKey = (event.key === 'Enter' && !event.shiftKey) || event.key === 'Tab'
    if (!isCommitKey) return
    event.preventDefault()
    event.stopPropagation()
    const movement = event.key === 'Tab' ? ([event.shiftKey ? -1 : 1, 0] as const) : ([0, 1] as const)
    onFinishedEditing({ ...value, data: event.currentTarget.value }, movement)
  }

  return (
    <TextCellEntry
      highlight={isHighlighted}
      autoFocus={value.readonly !== true}
      disabled={value.readonly === true}
      altNewline
      value={value.data}
      validatedSelection={validatedSelection}
      onChange={(event: ChangeEvent<HTMLTextAreaElement>) => onChange({ ...value, data: event.target.value })}
      onKeyDown={handleKeyDown}
    />
  )
}

const provideCellEditor: ProvideEditorCallback<GridCell> = (cell) =>
  cell.kind === GridCellKind.Text ? (TextCellEditor as ProvideEditorComponent<GridCell>) : undefined

export function SpreadsheetDataEditor(props: DataEditorProps) {
  // Not React state: mutated synchronously inside a native-event handler and
  // read back inside `TextCellEditor`'s effect a moment later — see module
  // header (F6). A single ref is enough because glide-data-grid only ever
  // has one overlay open at a time.
  const pendingSeedRef = useRef<PendingSeed | null>(null)

  const onKeyDownIn = props.onKeyDown
  const handleKeyDown = useCallback(
    (event: GridKeyEventArgs) => {
      onKeyDownIn?.(event)
      if (event.location === undefined) return
      const [col, row] = event.location
      const pending = pendingSeedRef.current
      const pendingHere = pending !== null && pending.row === row && pending.col === col ? pending : null

      if (pendingHere !== null) {
        const movement = commitMovementFor(event)
        if (movement !== undefined) {
          // Enter/Tab arrived before the overlay mounted — glide-data-grid
          // would otherwise treat it as plain navigation (see module header)
          // and silently discard the seed and everything buffered so far.
          // Record the commit; `TextCellEditor` finishes editing with it the
          // moment it mounts.
          pendingHere.commitOnMount = movement
          event.cancel()
          event.preventDefault()
          event.stopPropagation()
          return
        }
        if (event.key === 'Escape') {
          // Not a race to guard — glide-data-grid has no overlay open yet to
          // close — but an abandoned buffer must not leak into whatever
          // happens next, so drop it and let Escape proceed normally.
          pendingSeedRef.current = null
          return
        }
        if (isEditOnTypeKey(event)) {
          // The overlay for this exact cell is still opening (see module
          // header) — buffer this keystroke instead of letting
          // glide-data-grid's own `editOnType` re-trigger `reselect()` and
          // silently drop everything seeded/buffered so far.
          pendingHere.chars.push(event.key)
          event.cancel()
          event.preventDefault()
          event.stopPropagation()
        }
        return
      }

      if (!isEditOnTypeKey(event)) return
      // First qualifying keystroke for this cell: let glide-data-grid's own
      // `editOnType` open the overlay as usual (there is no public API to do
      // that ourselves — see module header), and start tracking it.
      const seed: PendingSeed = { row, col, chars: [], startedAt: performance.now() }
      pendingSeedRef.current = seed
      // glide-data-grid preventDefault()s this keydown when `editOnType`
      // really opens the overlay (right after this handler returns). If it
      // didn't — a read-only grid or cell, a cell scrolled out of view — no
      // editor is coming, and a seed left behind would hold every shortcut
      // (F6b) until PENDING_EDIT_MAX_MS.
      const raw = event.rawEvent
      if (raw !== undefined) {
        queueMicrotask(() => {
          if (!raw.isDefaultPrevented() && pendingSeedRef.current === seed) pendingSeedRef.current = null
        })
      }
    },
    [onKeyDownIn],
  )

  // F6b — see module header.
  const keyHoldRef = useRef<KeyHold | null>(null)
  useEffect(() => {
    const keyHold = new KeyHold(() => {
      const pending = pendingSeedRef.current
      return pending !== null && performance.now() - pending.startedAt < PENDING_EDIT_MAX_MS
    })
    keyHoldRef.current = keyHold
    const detach = keyHold.attach()
    return () => {
      detach()
      keyHoldRef.current = null
    }
  }, [])

  const holdKeysOnMouseDown = useCallback((event: MouseEvent<HTMLDivElement>) => {
    // The click lands on glide's `.dvn-scroller`, layered over the canvas.
    if (event.button !== 0 || !(event.target instanceof Element) || event.target.closest('.dvn-scroller') === null) return
    const active = document.activeElement
    if (active !== null && (event.currentTarget.contains(active) || active.closest('#portal') !== null)) return
    keyHoldRef.current?.start(event.currentTarget)
  }, [])

  const releaseKeysOnGridFocus = useCallback(() => keyHoldRef.current?.release(), [])

  return (
    <PendingSeedContext.Provider value={pendingSeedRef}>
      <div style={{ display: 'contents' }} onMouseDownCapture={holdKeysOnMouseDown} onFocusCapture={releaseKeysOnGridFocus}>
        <DataEditor provideEditor={provideCellEditor} {...props} onKeyDown={handleKeyDown} />
      </div>
    </PendingSeedContext.Provider>
  )
}
