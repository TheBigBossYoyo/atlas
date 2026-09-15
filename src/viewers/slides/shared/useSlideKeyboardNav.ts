/**
 * S15 — keyboard navigation shared by `PptxViewer`/`OdpViewer`: PageUp/PageDown
 * (existing), plus Arrow keys, Space (next/previous), and Home/End (first/last).
 * Ignored while focus is inside an editable control.
 *
 * Registered through the shell's centralized shortcut dispatcher
 * (`useViewerShortcuts`, wave-3 shell-polish follow-up) instead of its own
 * raw `window.addEventListener('keydown', ...)`, for two reasons: it
 * participates in the same viewer/shell precedence as every other shortcut
 * (a modal opened on top of the slide viewer gets first refusal), and — the
 * concrete bug this migration fixes — a focused thumbnail-rail `<button>`
 * (`SlideDeck`'s rail) must still respond to Space as a native button
 * activation (select that slide via `onClick`) rather than having this hook
 * unconditionally `preventDefault()` it into "advance to the next slide"
 * instead, which silently broke keyboard activation of the thumbnail rail.
 */

import { useCallback } from 'react'

import { useViewerShortcuts } from '../../../hooks/useShortcutManager'
import type { ShortcutHandler } from '../../../hooks/shortcutManagerContext'

type SetActiveIndex = (updater: (currentIndex: number) => number) => void

function isEditableTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement
    && (target.isContentEditable
      || target.tagName === 'INPUT'
      || target.tagName === 'TEXTAREA'
      || target.tagName === 'SELECT')
  )
}

/** A focused `<button>` (the thumbnail rail's own slide-select buttons) owns
 * Space/Enter itself — this hook must not hijack Space away from it. */
function isFocusedButtonTarget(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && target.tagName === 'BUTTON'
}

export function useSlideKeyboardNav(slideCount: number, setActiveIndex: SetActiveIndex): void {
  const handler = useCallback<ShortcutHandler>(
    (event) => {
      if (slideCount === 0 || isEditableTarget(event.target)) {
        return false
      }

      switch (event.key) {
        case 'PageDown':
        case 'ArrowRight':
        case 'ArrowDown':
          event.preventDefault()
          setActiveIndex(currentIndex => Math.min(currentIndex + 1, slideCount - 1))
          return true
        case ' ':
          if (isFocusedButtonTarget(event.target)) {
            // Let the browser activate the focused thumbnail button instead.
            return false
          }
          event.preventDefault()
          setActiveIndex(currentIndex => Math.min(currentIndex + 1, slideCount - 1))
          return true
        case 'PageUp':
        case 'ArrowLeft':
        case 'ArrowUp':
          event.preventDefault()
          setActiveIndex(currentIndex => Math.max(currentIndex - 1, 0))
          return true
        case 'Home':
          event.preventDefault()
          setActiveIndex(() => 0)
          return true
        case 'End':
          event.preventDefault()
          setActiveIndex(() => slideCount - 1)
          return true
        default:
          return false
      }
    },
    [setActiveIndex, slideCount],
  )

  useViewerShortcuts(handler)
}
