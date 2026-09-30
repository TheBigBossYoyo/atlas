/**
 * VERSIONS-1 — the version history panel: what the document looked like at each
 * point, newest first, with a way back to any of them.
 *
 * RESTORING IS GATED ON THE DOCUMENT BEING SAVED, and that is a deliberate
 * safety decision rather than an omission. Restoring writes the chosen version's
 * bytes over the document; main records a version of whatever it writes, so the
 * state being replaced survives ONLY if it was itself already written. Allowing a
 * restore over unsaved edits would therefore destroy work that exists nowhere
 * else, silently. With unsaved changes the action is disabled and says why.
 *
 * A plain `<ul>` of `<li>`s with real buttons, deliberately not a `role="menu"` /
 * `role="listbox"` construction: A11Y pass 4 was a full day of learning that a
 * container role which forbids interactive children (`tablist` did) turns every
 * control inside it into something assistive software describes wrongly. This
 * markup has no such trap.
 */
import { History, RotateCcw, Trash2, X } from 'lucide-react'
import { useCallback, useRef, useState } from 'react'

import type { DocumentVersion } from '../electron'
import { useFocusTrap } from '../hooks/useFocusTrap'
import { useShellShortcut } from '../hooks/useShortcutManager'
import type { VersionHistoryState } from '../hooks/useVersionHistory'
import { useTranslate } from '../i18n'
import { formatBytes, formatRelativeTime } from '../utils/formatRelativeTime'

export interface VersionHistoryPanelProps {
  readonly open: boolean
  readonly onClose: () => void
  readonly history: VersionHistoryState
  /** Puts a version back. Only offered when the document has no unsaved changes. */
  readonly onRestore: (version: DocumentVersion) => Promise<void>
  /** Unsaved changes block restoring — see this file's header. */
  readonly documentDirty: boolean
  /** Records the state the user has right now, for the "save a version" action. */
  readonly onCaptureNow: (() => Promise<void>) | null
}

export function VersionHistoryPanel({
  open,
  onClose,
  history,
  onRestore,
  documentDirty,
  onCaptureNow,
}: VersionHistoryPanelProps) {
  const t = useTranslate()
  const panelRef = useRef<HTMLDivElement>(null)
  const [busyId, setBusyId] = useState<string | null>(null)
  const [confirmingClear, setConfirmingClear] = useState(false)

  useShellShortcut(
    useCallback(
      (event) => {
        if (event.key !== 'Escape') return false
        onClose()
        return true
      },
      [onClose],
    ),
    open,
  )

  useFocusTrap(panelRef, open)

  const restore = useCallback(
    async (version: DocumentVersion) => {
      setBusyId(version.id)
      try {
        await onRestore(version)
      } finally {
        setBusyId(null)
      }
    },
    [onRestore],
  )

  if (!open) return null

  return (
    <aside className="version-history" ref={panelRef} aria-label={t('history.panelAria')}>
      <header className="version-history__header">
        <h2 className="version-history__title">
          <History size={16} aria-hidden="true" />
          {t('history.title')}
        </h2>
        <button type="button" className="version-history__close" onClick={onClose} aria-label={t('history.closeAria')}>
          <X size={16} />
        </button>
      </header>

      {onCaptureNow !== null && (
        <button type="button" className="version-history__capture" onClick={() => void onCaptureNow()}>
          {t('history.captureNow')}
        </button>
      )}

      {documentDirty && <p className="version-history__notice">{t('history.unsavedBlocksRestore')}</p>}
      {history.error !== null && <p className="version-history__error">{history.error}</p>}

      {history.versions.length === 0 ? (
        <p className="version-history__empty">
          {history.loading ? t('history.loading') : t('history.empty')}
        </p>
      ) : (
        <ul className="version-history__list">
          {history.versions.map((version, index) => (
            <li key={`${version.id}-${version.at}`} className="version-history__item">
              <div className="version-history__meta">
                <span className="version-history__when">{formatRelativeTime(version.at, t)}</span>
                {/* The exact time as a tooltip: "3 days ago" is the right thing to
                    read at a glance and the wrong thing to rely on. */}
                <span className="version-history__exact" title={new Date(version.at).toLocaleString()}>
                  {new Date(version.at).toLocaleTimeString()}
                </span>
                <span className="version-history__size">{formatBytes(version.bytes)}</span>
                {index === 0 && <span className="version-history__badge">{t('history.latest')}</span>}
              </div>
              {version.label !== null && <p className="version-history__label">{version.label}</p>}
              <button
                type="button"
                className="version-history__restore"
                disabled={documentDirty || busyId !== null}
                onClick={() => void restore(version)}
              >
                <RotateCcw size={14} aria-hidden="true" />
                {busyId === version.id ? t('history.restoring') : t('history.restore')}
              </button>
            </li>
          ))}
        </ul>
      )}

      {history.versions.length > 0 && (
        <div className="version-history__footer">
          {confirmingClear ? (
            <>
              <span className="version-history__confirm">{t('history.clearConfirm')}</span>
              <button
                type="button"
                className="version-history__clear-confirm"
                onClick={() => {
                  setConfirmingClear(false)
                  void history.clear()
                }}
              >
                {t('history.clearYes')}
              </button>
              <button type="button" onClick={() => setConfirmingClear(false)}>
                {t('history.clearNo')}
              </button>
            </>
          ) : (
            // Two steps on purpose: this throws away the only record of how the
            // document got here, and it cannot be undone.
            <button type="button" className="version-history__clear" onClick={() => setConfirmingClear(true)}>
              <Trash2 size={14} aria-hidden="true" />
              {t('history.clear')}
            </button>
          )}
        </div>
      )}
    </aside>
  )
}
