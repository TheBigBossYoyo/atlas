/**
 * SHELL-17 — the open-documents strip.
 *
 * Switching away from a document with unsaved changes goes through the
 * shell's existing Save/Discard/Cancel prompt (see `App.tsx`), so a tab only
 * ever shows the dirty dot for the document currently being edited.
 */
import { memo, useState, type DragEvent } from 'react'
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

  if (sessions.length === 0) return null

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
              type="button"
              role="tab"
              id={`tab-${session.id}`}
              aria-selected={isActive}
              className="tab-bar__label"
              title={session.id}
              onClick={() => onSelect(session.id)}
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
