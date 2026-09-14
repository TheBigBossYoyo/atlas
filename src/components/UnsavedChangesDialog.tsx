import { useCallback } from 'react';
import { AlertTriangle } from 'lucide-react';
import { useShellShortcut } from '../hooks/useShortcutManager';

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

  if (!isOpen) return null;

  return (
    <div className="modal-backdrop" onClick={() => !isSaving && onCancel()}>
      <div
        className="modal confirm-dialog"
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="unsaved-changes-title"
        onClick={e => e.stopPropagation()}
      >
        <div className="confirm-dialog__icon">
          <AlertTriangle size={22} aria-hidden="true" />
        </div>
        <h2 id="unsaved-changes-title" className="modal__title">Unsaved changes</h2>
        <p className="confirm-dialog__body">
          This document has unsaved changes. Save them before continuing, discard them, or cancel.
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
            Cancel
          </button>
          <button
            className="confirm-dialog__btn confirm-dialog__btn--danger"
            onClick={onDiscard}
            disabled={isSaving}
          >
            Discard
          </button>
          <button
            className="confirm-dialog__btn confirm-dialog__btn--primary"
            onClick={onSave}
            disabled={isSaving}
          >
            {isSaving ? 'Saving…' : 'Save'}
          </button>
        </div>
      </div>
    </div>
  );
}
