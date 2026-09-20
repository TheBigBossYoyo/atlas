import { useCallback, useId, useRef, useState, type FormEvent } from 'react'

import { useFocusTrap } from '../hooks/useFocusTrap'
import { useShellShortcut } from '../hooks/useShortcutManager'
import { useTranslate } from '../i18n'

export type DocxPromptDialogProps = {
  readonly isOpen: boolean
  readonly titleText: string
  readonly label: string
  readonly initialValue: string
  readonly confirmLabel: string
  readonly onConfirm: (value: string) => void
  readonly onCancel: () => void
}

/**
 * F1 — a real in-app modal replacing the three `window.prompt` calls
 * DocxViewer used for Insert Hyperlink (toolbar + Ctrl+K), Add Comment, and
 * Reply to Comment. Electron does not implement `window.prompt` (it throws
 * "prompt() is not supported"), so all three previously did nothing at all
 * in the packaged app — no dialog ever appeared, nothing was inserted, and
 * the error landed silently in the console.
 *
 * A single shared component handles all three call sites; DocxViewer picks
 * the copy (title/label/confirm button text) and initial value per call
 * site and remounts this component (via a `key` bumped on every open,
 * mirroring `PdfPasswordDialog`'s documented pattern) so a value typed for
 * one prompt never bleeds into the next — e.g. a half-typed reply to one
 * comment must not reappear when replying to a different comment.
 */
export function DocxPromptDialog({
  isOpen,
  titleText,
  label,
  initialValue,
  confirmLabel,
  onConfirm,
  onCancel,
}: DocxPromptDialogProps) {
  const t = useTranslate()
  const [value, setValue] = useState(initialValue)
  const dialogRef = useRef<HTMLFormElement | null>(null)
  const titleId = useId()
  const inputId = useId()

  // Escape cancels — registered with the shared shortcut dispatcher instead
  // of an ad hoc `window` listener, matching UnsavedChangesDialog; only
  // listens while the dialog is open.
  useShellShortcut(
    useCallback(
      (e) => {
        if (e.key !== 'Escape') return false
        onCancel()
        return true
      },
      [onCancel],
    ),
    isOpen,
  )

  // A11Y-2 — shared focus trap: focuses the text input on open (it's the
  // first focusable element in the dialog), keeps Tab/Shift+Tab cycling
  // within it while open, and restores focus to whatever triggered it once
  // it closes.
  useFocusTrap(dialogRef, isOpen)

  if (!isOpen) return null

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    onConfirm(value)
  }

  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <form
        ref={dialogRef}
        className="modal docx-prompt-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        onClick={(event) => event.stopPropagation()}
        onSubmit={handleSubmit}
      >
        <h2 id={titleId} className="modal__title">{titleText}</h2>
        <div className="docx-prompt-dialog__field">
          <label htmlFor={inputId} className="docx-prompt-dialog__label">
            {label}
          </label>
          <input
            id={inputId}
            type="text"
            className="docx-prompt-dialog__input"
            value={value}
            onChange={(event) => setValue(event.target.value)}
            autoComplete="off"
            spellCheck={false}
          />
        </div>
        <div className="docx-prompt-dialog__actions">
          <button
            type="button"
            className="confirm-dialog__btn confirm-dialog__btn--secondary"
            onClick={onCancel}
          >
            {t('unsavedDialog.cancel')}
          </button>
          <button type="submit" className="confirm-dialog__btn confirm-dialog__btn--primary">
            {confirmLabel}
          </button>
        </div>
      </form>
    </div>
  )
}
