/**
 * SHELL-17 — the open-documents strip.
 *
 * Switching away from a document with unsaved changes goes through the
 * shell's existing Save/Discard/Cancel prompt (see `App.tsx`), so a tab only
 * ever shows the dirty dot for the document currently being edited.
 */
import { memo, useRef, useState, type DragEvent, type KeyboardEvent } from 'react'
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
  const [focusedId, setFocusedId] = useState<string | null>(activeId)
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([])

  if (sessions.length === 0) return null

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

  const handleTabKeyDown = (event: KeyboardEvent<HTMLButtonElement>, index: number): void => {
    switch (event.key) {
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
                onClose(session.id)
              }
            }}
          >
            <button
              ref={(node) => {
                tabRefs.current[index] = node
              }}
              type="button"
              role="tab"
              id={`tab-${session.id}`}
              aria-selected={isActive}
              tabIndex={index === focusedIndex ? 0 : -1}
              className="tab-bar__label"
              title={session.id}
              onClick={() => onSelect(session.id)}
              onFocus={() => setFocusedId(session.id)}
              onKeyDown={(event) => handleTabKeyDown(event, index)}
            >
              {showsDirtyDot && <span className="tab-bar__dirty" aria-label={t('common.unsavedChanges')} />}
              <span className="tab-bar__name">{session.name}</span>
            </button>
            <button
              type="button"
              className="tab-bar__close"
              aria-label={t('tabBar.closeAria', { name: session.name })}
              title={t('tabBar.closeTitle', { name: session.name })}
              onClick={() => onClose(session.id)}
            >
              <X size={13} />
            </button>
          </div>
        )
      })}
    </div>
  )
}

export const TabBar = memo(TabBarBase)
