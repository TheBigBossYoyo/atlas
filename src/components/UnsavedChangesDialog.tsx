import { useCallback, useEffect, useRef } from 'react';
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

const FOCUSABLE_SELECTOR = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';

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
  const dialogRef = useRef<HTMLDivElement>(null);
  const previouslyFocusedRef = useRef<HTMLElement | null>(null);

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

  // Focus trap: move focus into the dialog on open (this is a data-loss-
  // adjacent alertdialog — every open path funnels through it — so it must
  // not be possible to Tab straight past it into the document behind it),
  // keep Tab/Shift+Tab cycling within it while open, and restore focus to
  // whatever triggered it once it closes. Same pattern as ShortcutsModal's
  // UX-13 fix; this dialog never had it.
  useEffect(() => {
    if (!isOpen) return;

    previouslyFocusedRef.current = document.activeElement as HTMLElement | null;
    const dialog = dialogRef.current;
    dialog?.querySelector<HTMLElement>(FOCUSABLE_SELECTOR)?.focus();

    const handleTab = (e: KeyboardEvent) => {
      if (e.key !== 'Tab' || !dialog) return;
      const focusables = Array.from(dialog.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)).filter(
        el => !el.hasAttribute('disabled'),
      );
      if (focusables.length === 0) return;
      const first = focusables[0]!;
      const last = focusables[focusables.length - 1]!;
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    };

    window.addEventListener('keydown', handleTab);
    return () => {
      window.removeEventListener('keydown', handleTab);
      previouslyFocusedRef.current?.focus();
    };
  }, [isOpen]);

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
