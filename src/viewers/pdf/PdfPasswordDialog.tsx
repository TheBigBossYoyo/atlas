import { useEffect, useRef, useState } from 'react'
import type { FormEvent } from 'react'
import { KeyRound } from 'lucide-react'
import { useFocusTrap } from '../../hooks/useFocusTrap'

export type PdfPasswordDialogProps = {
  readonly isOpen: boolean
  readonly isIncorrect: boolean
  readonly onSubmit: (password: string) => void
  readonly onCancel: () => void
}

/**
 * Modal prompt wired to pdf.js's `loadingTask.onPassword` (PDF-14/P10).
 *
 * The caller should remount this component (via a changing `key`) each time
 * a fresh password prompt starts, rather than this component resetting its
 * own `password` state in an effect — the latter causes an extra render on
 * every open and is the pattern React's own hooks lint now flags.
 */
export function PdfPasswordDialog({ isOpen, isIncorrect, onSubmit, onCancel }: PdfPasswordDialogProps) {
  const [password, setPassword] = useState('')
  const inputRef = useRef<HTMLInputElement | null>(null)
  const dialogRef = useRef<HTMLFormElement | null>(null)

  useEffect(() => {
    if (isOpen) {
      const timer = setTimeout(() => inputRef.current?.focus(), 50)
      return () => clearTimeout(timer)
    }
    return undefined
  }, [isOpen])

  // A11Y-2 — Tab was never contained inside the dialog, so it leaked focus
  // straight into the PDF page behind it. The input above already handles
  // its own (animation-delayed) autofocus, so this only adds Tab
  // containment and restores focus to whatever triggered the prompt on close.
  useFocusTrap(dialogRef, isOpen, { focusOnOpen: false })

  if (!isOpen) return null

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (password.length > 0) {
      onSubmit(password)
    }
  }

  return (
    <div className="pdf-viewer__password-overlay" role="dialog" aria-modal="true" aria-label="Password required">
      <form ref={dialogRef} className="pdf-viewer__password-dialog" onSubmit={handleSubmit}>
        <KeyRound size={24} className="pdf-viewer__password-icon" aria-hidden="true" />
        <p className="pdf-viewer__password-title">This PDF is password protected</p>
        <input
          ref={inputRef}
          type="password"
          className="pdf-viewer__password-input"
          placeholder="Enter password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          aria-label="PDF password"
        />
        {isIncorrect && <p className="pdf-viewer__password-error">Incorrect password. Try again.</p>}
        <div className="pdf-viewer__password-actions">
          <button type="button" className="pdf-viewer__password-cancel" onClick={onCancel}>
            Cancel
          </button>
          <button type="submit" className="pdf-viewer__password-submit" disabled={password.length === 0}>
            Unlock
          </button>
        </div>
      </form>
    </div>
  )
}
