import { useCallback, useRef } from 'react';
import { AlertTriangle } from 'lucide-react';
import { useShellShortcut } from '../hooks/useShortcutManager';
import { useFocusTrap } from '../hooks/useFocusTrap';
import { useTranslate } from '../i18n';

interface UnsavedChangesDialogProps {
  isOpen: boolean;
  isSaving: boolean;
  errorMessage: string | null;
  onSave: () => void;
  onDiscard: () => void;
  onCancel: () => void;
}

/**
 * P1.1 — the Save / Discard / Cancel confirmation shown before any open path
 * (dialog, Recent click, drag-drop, OS "Open with") would otherwise silently
 * replace unsaved edits. "Save" triggers the active viewer's own save and
 * only proceeds with the pending open once it actually succeeds.
 */
export function UnsavedChangesDialog({
  isOpen,
  isSaving,
  errorMessage,
  onSave,
  onDiscard,
  onCancel,
}: UnsavedChangesDialogProps) {
  const t = useTranslate();
  const dialogRef = useRef<HTMLDivElement>(null);

  // P2.1 — Escape-close registered with the shared dispatcher instead of its
  // own ad hoc `window` listener; only listens while the dialog is open.
  useShellShortcut(
    useCallback(
      (e) => {
        if (e.key !== 'Escape' || isSaving) return false;
        onCancel();
        return true;
      },
      [isSaving, onCancel],
    ),
    isOpen,
  );

  // A11Y-1 — focus trap: move focus into the dialog on open (this is a
  // data-loss-adjacent alertdialog — every open path funnels through it — so
  // it must not be possible to Tab straight past it into the document behind
  // it), keep Tab/Shift+Tab cycling within it while open (skipping disabled
  // controls, e.g. while isSaving), and restore focus to whatever triggered
  // it once it closes. Migrated to the shared useFocusTrap hook — this used
  // to be its own hand-rolled copy of the pattern ShortcutsModal also had
  // (UX-13), now unified.
  useFocusTrap(dialogRef, isOpen);

  if (!isOpen) return null;

  return (
    <div className="modal-backdrop" onClick={() => !isSaving && onCancel()}>
      <div
        ref={dialogRef}
        className="modal confirm-dialog"
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="unsaved-changes-title"
        onClick={e => e.stopPropagation()}
      >
        <div className="confirm-dialog__icon">
          <AlertTriangle size={22} aria-hidden="true" />
        </div>
        <h2 id="unsaved-changes-title" className="modal__title">{t('unsavedDialog.title')}</h2>
        <p className="confirm-dialog__body">
          {t('unsavedDialog.body')}
        </p>
        {errorMessage && (
          <p className="confirm-dialog__error" role="alert">{errorMessage}</p>
        )}
        <div className="confirm-dialog__actions">
          <button
            className="confirm-dialog__btn confirm-dialog__btn--secondary"
            onClick={onCancel}
            disabled={isSaving}
          >
            {t('unsavedDialog.cancel')}
          </button>
          <button
            className="confirm-dialog__btn confirm-dialog__btn--danger"
            onClick={onDiscard}
            disabled={isSaving}
          >
            {t('unsavedDialog.discard')}
          </button>
          <button
            className="confirm-dialog__btn confirm-dialog__btn--primary"
            onClick={onSave}
            disabled={isSaving}
          >
            {isSaving ? t('unsavedDialog.saving') : t('unsavedDialog.save')}
          </button>
        </div>
      </div>
    </div>
  );
}
