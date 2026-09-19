/**
 * P4.7 — coverage for the PDF password-prompt dialog wired to pdfjs's
 * `loadingTask.onPassword` (PDF-14/P10). The plan's status notes call this
 * file "PdfWordDialog.tsx" at ~41% statement coverage; no such file exists
 * in `src/viewers/pdf` — the low-coverage dialog component there is actually
 * `PdfPasswordDialog.tsx` (a password prompt, not a word-lookup dialog), so
 * this test file targets that.
 */
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { PdfPasswordDialog } from '../PdfPasswordDialog'

describe('PdfPasswordDialog', () => {
  it('renders nothing when closed', () => {
    const { container } = render(
      <PdfPasswordDialog isOpen={false} isIncorrect={false} onSubmit={vi.fn()} onCancel={vi.fn()} />,
    )

    expect(container).toBeEmptyDOMElement()
  })

  it('renders an accessible modal prompt with an empty, focused password field when open', async () => {
    render(<PdfPasswordDialog isOpen isIncorrect={false} onSubmit={vi.fn()} onCancel={vi.fn()} />)

    const dialog = screen.getByRole('dialog', { name: 'Password required' })
    expect(dialog).toHaveAttribute('aria-modal', 'true')
    expect(screen.getByText('This PDF is password protected')).toBeInTheDocument()

    const input = screen.getByLabelText('PDF password')
    expect(input).toHaveAttribute('type', 'password')
    expect(input).toHaveValue('')

    // Focus is grabbed asynchronously (a 50ms timer) so the caller's
    // remount-via-key doesn't fight an in-progress render.
    await waitFor(() => expect(input).toHaveFocus())
  })

  it('does not show an error message when the password has not been rejected yet', () => {
    render(<PdfPasswordDialog isOpen isIncorrect={false} onSubmit={vi.fn()} onCancel={vi.fn()} />)

    expect(screen.queryByText('Incorrect password. Try again.')).not.toBeInTheDocument()
  })

  it('shows an error message when the previous attempt was incorrect', () => {
    render(<PdfPasswordDialog isOpen isIncorrect onSubmit={vi.fn()} onCancel={vi.fn()} />)

    expect(screen.getByText('Incorrect password. Try again.')).toBeInTheDocument()
  })

  it('disables Unlock until a password is typed, and enables it once there is input', () => {
    render(<PdfPasswordDialog isOpen isIncorrect={false} onSubmit={vi.fn()} onCancel={vi.fn()} />)

    const submit = screen.getByRole('button', { name: 'Unlock' })
    expect(submit).toBeDisabled()

    fireEvent.change(screen.getByLabelText('PDF password'), { target: { value: 'hunter2' } })
    expect(submit).toBeEnabled()
  })

  it('submits the typed password and does not call onSubmit for an empty value', () => {
    const onSubmit = vi.fn()
    render(<PdfPasswordDialog isOpen isIncorrect={false} onSubmit={onSubmit} onCancel={vi.fn()} />)

    const input = screen.getByLabelText('PDF password')
    const form = input.closest('form')
    expect(form).not.toBeNull()

    // Submitting the form while empty (e.g. a stray Enter before the guard
    // button-disabled state would even matter) must not call onSubmit.
    fireEvent.submit(form!)
    expect(onSubmit).not.toHaveBeenCalled()

    fireEvent.change(input, { target: { value: 'correct horse battery staple' } })
    fireEvent.click(screen.getByRole('button', { name: 'Unlock' }))

    expect(onSubmit).toHaveBeenCalledTimes(1)
    expect(onSubmit).toHaveBeenCalledWith('correct horse battery staple')
  })

  it('calls onCancel when Cancel is clicked', () => {
    const onCancel = vi.fn()
    render(<PdfPasswordDialog isOpen isIncorrect={false} onSubmit={vi.fn()} onCancel={onCancel} />)

    fireEvent.click(screen.getByRole('button', { name: 'Cancel' }))
    expect(onCancel).toHaveBeenCalledTimes(1)
  })

  it('starts a fresh, empty password field when remounted (via a changing key) for a new attempt', () => {
    const { rerender } = render(
      <PdfPasswordDialog key={1} isOpen isIncorrect={false} onSubmit={vi.fn()} onCancel={vi.fn()} />,
    )
    fireEvent.change(screen.getByLabelText('PDF password'), { target: { value: 'wrong-guess' } })
    expect(screen.getByLabelText('PDF password')).toHaveValue('wrong-guess')

    // Simulates PdfViewer bumping `passwordAttemptId` (its `key`) after an
    // INCORRECT_PASSWORD callback from pdfjs.
    rerender(<PdfPasswordDialog key={2} isOpen isIncorrect onSubmit={vi.fn()} onCancel={vi.fn()} />)

    expect(screen.getByLabelText('PDF password')).toHaveValue('')
    expect(screen.getByText('Incorrect password. Try again.')).toBeInTheDocument()
  })
})
