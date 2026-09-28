/**
 * SHELL-17 — the open-documents strip.
 *
 * Switching away from a document with unsaved changes goes through the
 * shell's existing Save/Discard/Cancel prompt (see `App.tsx`), so a tab only
 * ever shows the dirty dot for the document currently being edited.
 */
import { memo, useEffect, useRef, useState, type DragEvent, type KeyboardEvent } from 'react'
import { X } from 'lucide-react'

import type { DocumentSession } from '../session/documentSessions'
import { useTranslate } from '../i18n'

export type TabBarProps = {
  readonly sessions: ReadonlyArray<DocumentSession>
  readonly activeId: string | null
  readonly isActiveDirty: boolean
  readonly onSelect: (id: string) => void
  readonly onClose: (id: string) => void
  readonly onReorder: (fromIndex: number, toIndex: number) => void
}

function TabBarBase({ sessions, activeId, isActiveDirty, onSelect, onClose, onReorder }: TabBarProps) {
  const t = useTranslate()
  const [draggingIndex, setDraggingIndex] = useState<number | null>(null)

  // A11Y pass 3 — role="tab" implies the ARIA APG tablist keyboard pattern
  // (Left/Right/Home/End move focus among tabs, exactly one tab is ever a
  // Tab stop at a time), which this bar never implemented: every tab button
  // just sat in normal Tab order, so a screen reader announcing "tab 1 of N"
  // was promising arrow-key navigation that silently did nothing. This is
  // "manual activation" (arrowing only MOVES focus; Enter/Space or a click
  // actually switches documents) rather than the simpler "focus follows
  // selection" variant, because selecting away from a dirty document opens
  // the shell's Save/Discard/Cancel prompt (see this file's own header
  // comment) — arrow-key browsing through open tabs must never trigger that
  // as a side effect.
  //
  // A11Y pass 4 (axe audit, 2026-09-28) — `role="tab"` moved from the inner
  // button onto the tab's own wrapper element. `tablist` may only own `tab`
  // children, and the close button sat right beside the tab as a second child of
  // the list: axe reported it as a CRITICAL `aria-required-children` violation on
  // every screen in the app, because the tab bar is always mounted. With the
  // wrapper being the `tab`, the close button is a DESCENDANT of a tab rather
  // than a stray child of the list, which is the structure Deque's own guidance
  // for closable tabs describes.
  //
  // That exposed the other half of the same problem: a FOCUSABLE control inside a
  // `role="tab"` is a `nested-interactive` violation (WCAG 4.1.2), and axe is
  // explicit that `tabIndex={-1}` plus `aria-hidden` does not excuse it. There is
  // no arrangement that keeps the tablist pattern AND a focusable close button,
  // so the X became a mouse-only `<span>` and closing by keyboard is Delete or
  // Backspace on the focused tab (plus the existing Ctrl+W) — which is what
  // browser tab strips do, is fewer keystrokes than tabbing to a button, and is
  // listed in the shortcuts dialog. The alternative considered and rejected was
  // dropping `tablist`/`tab` for `list`/`listitem`: conformant, but it gives up
  // "tab 3 of 8, selected" and puts two Tab stops per open document back in the
  // way of reaching the document itself.
  const [focusedId, setFocusedId] = useState<string | null>(activeId)
  const tabRefs = useRef<Array<HTMLDivElement | null>>([])

  // A11Y-3 — closing a tab unmounts the element that had focus, leaving it with
  // nowhere to go, so it fell back to
  // `document.body` and the next Tab press started over from the top of the
  // window. `useRestoreFocusOnClose` doesn't fit here — it restores focus to
  // one fixed trigger across a single open/close toggle, but a tab list is a
  // dynamic array where the "trigger" (the closed tab itself) is gone for
  // good. Instead: remember which neighbouring tab should pick up focus
  // *before* asking the parent to remove the session, then once `sessions`
  // actually reflects the removal, focus that neighbour — mirroring how browser
  // tab strips let you keep closing tabs by pressing the same key in place, which
  // is more useful here than restoring focus to whatever was focused before the
  // tab bar existed. (A11Y pass 4 moved this target from the neighbour's close
  // button to the neighbour tab itself, the close button no longer being
  // focusable.)
  //
  // CTRLW-FOCUS-1 — the `pendingFocusIdRef` handshake below only fires for closes
  // that go through this component's own `handleClose`. The global Ctrl+W in
  // `App.tsx` calls `closeSessionById` directly, so that route left focus on
  // `document.body` exactly as described above. Rather than have every caller
  // remember to announce intent, the effect now ALSO re-homes focus whenever the
  // session list shrank and the tab that had focus is gone — which covers Ctrl+W,
  // the toolbar Close and anything added later, with no cooperation needed.
  //
  // Guarded on focus having been inside the tab bar at the time. Ctrl+W pressed
  // while the caret is in the document must NOT yank focus up to a tab's close
  // button; that would be worse than the bug. Where focus *should* go for that case
  // (the newly-active document, presumably) is a larger question about ViewerRouter
  // and is deliberately left alone here rather than guessed at.
  const pendingFocusIdRef = useRef<string | null>(null)
  const previousSessionsRef = useRef<ReadonlyArray<{ id: string }>>(sessions)

  useEffect(() => {
    const previous = previousSessionsRef.current
    previousSessionsRef.current = sessions

    const explicitTargetId = pendingFocusIdRef.current
    pendingFocusIdRef.current = null

    const focusTabByIndex = (index: number): void => {
      const target = sessions[index]
      if (target === undefined) return
      setFocusedId(target.id)
      tabRefs.current[index]?.focus()
    }

    if (explicitTargetId !== null) {
      const targetIndex = sessions.findIndex((s) => s.id === explicitTargetId)
      if (targetIndex !== -1) focusTabByIndex(targetIndex)
      return
    }

    // No explicit handshake: infer it. Only act on a shrink, and only when focus was
    // in this tab bar — see the note above.
    if (sessions.length >= previous.length) return
    const activeElement = document.activeElement
    const focusWasInTabBar =
      activeElement === null ||
      activeElement === document.body ||
      // A11Y pass 4 — the tab itself is what holds focus in this bar now; the
      // close buttons this used to check are not focusable any more.
      tabRefs.current.includes(activeElement as HTMLDivElement)
    if (!focusWasInTabBar) return

    const closedIndex = previous.findIndex((p) => !sessions.some((s) => s.id === p.id))
    if (closedIndex === -1) return
    // The tab that slid into the closed one's place, or the last one if it was last.
    focusTabByIndex(Math.min(closedIndex, sessions.length - 1))
  }, [sessions])

  if (sessions.length === 0) return null

  const handleClose = (index: number, id: string): void => {
    const neighbor = sessions[index + 1] ?? sessions[index - 1] ?? null
    pendingFocusIdRef.current = neighbor ? neighbor.id : null
    onClose(id)
  }

  const focusedIndex = (() => {
    const idx = sessions.findIndex((s) => s.id === focusedId)
    if (idx !== -1) return idx
    const activeIdx = sessions.findIndex((s) => s.id === activeId)
    return activeIdx !== -1 ? activeIdx : 0
  })()

  const focusTabAt = (index: number): void => {
    const target = sessions[index]
    if (!target) return
    setFocusedId(target.id)
    tabRefs.current[index]?.focus()
  }

  const handleTabKeyDown = (event: KeyboardEvent<HTMLDivElement>, index: number): void => {
    switch (event.key) {
      // A11Y pass 4 — the close button is deliberately not a Tab stop (see the
      // note by `tabRefs`), so closing gets a key on the tab itself. Both keys,
      // because both are what people try.
      case 'Delete':
      case 'Backspace': {
        event.preventDefault()
        const session = sessions[index]
        if (session !== undefined) handleClose(index, session.id)
        break
      }
      // Enter and Space activate a manually-activated tab. A `<button>` used to
      // give this for free; a `div[role="tab"]` does not.
      case 'Enter':
      case ' ': {
        event.preventDefault()
        const session = sessions[index]
        if (session !== undefined) onSelect(session.id)
        break
      }
      case 'ArrowRight':
        event.preventDefault()
        focusTabAt((index + 1) % sessions.length)
        break
      case 'ArrowLeft':
        event.preventDefault()
        focusTabAt((index - 1 + sessions.length) % sessions.length)
        break
      case 'Home':
        event.preventDefault()
        focusTabAt(0)
        break
      case 'End':
        event.preventDefault()
        focusTabAt(sessions.length - 1)
        break
      default:
        break
    }
  }

  const handleDrop = (event: DragEvent<HTMLDivElement>, index: number): void => {
    event.preventDefault()
    if (draggingIndex !== null) onReorder(draggingIndex, index)
    setDraggingIndex(null)
  }

  return (
    <div className="tab-bar" role="tablist" aria-label={t('tabBar.label')}>
      {sessions.map((session, index) => {
        const isActive = session.id === activeId
        const showsDirtyDot = isActive && isActiveDirty
        return (
          <div
            key={session.id}
            ref={(node) => {
              tabRefs.current[index] = node
            }}
            role="tab"
            id={`tab-${session.id}`}
            aria-selected={isActive}
            // An explicit name rather than name-from-content: the close button's
            // own `aria-label` would otherwise be concatenated into the tab's
            // name ("sample.docx, Close sample.docx"), and the unsaved dot needs
            // saying in words either way.
            aria-label={showsDirtyDot ? t('tabBar.tabDirtyAria', { name: session.name }) : session.name}
            tabIndex={index === focusedIndex ? 0 : -1}
            title={session.id}
            onClick={() => onSelect(session.id)}
            onFocus={() => setFocusedId(session.id)}
            onKeyDown={(event) => handleTabKeyDown(event, index)}
            className={isActive ? 'tab-bar__tab tab-bar__tab--active' : 'tab-bar__tab'}
            draggable
            onDragStart={() => setDraggingIndex(index)}
            onDragOver={(event) => event.preventDefault()}
            onDrop={(event) => handleDrop(event, index)}
            onDragEnd={() => setDraggingIndex(null)}
            // Middle-click closes, as in every browser and editor.
            onAuxClick={(event) => {
              if (event.button === 1) {
                event.preventDefault()
                handleClose(index, session.id)
              }
            }}
          >
            <span className="tab-bar__label">
              {/* Decorative: the unsaved state is in the tab's own `aria-label`,
                  so labelling the dot too would say it twice. */}
              {showsDirtyDot && <span className="tab-bar__dirty" aria-hidden="true" />}
              <span className="tab-bar__name">{session.name}</span>
            </span>
            {/* A `<span>`, not a `<button>` — see the `nested-interactive` note
                by `tabRefs`. `data-close-tab` is how tests reach it now that it
                has no role and no accessible name. */}
            <span
              className="tab-bar__close"
              data-close-tab={session.id}
              aria-hidden="true"
              title={t('tabBar.closeTitle', { name: session.name })}
              onClick={(event) => {
                // Without this the click also reaches the tab wrapper, which owns
                // the select handler — closing a background tab would select it
                // on the way out.
                event.stopPropagation()
                handleClose(index, session.id)
              }}
            >
              <X size={13} />
            </span>
          </div>
        )
      })}
    </div>
  )
}

export const TabBar = memo(TabBarBase)
