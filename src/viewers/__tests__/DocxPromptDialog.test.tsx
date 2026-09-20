/**
 * DocxPromptDialog — the shared single-line text-input modal behind Insert
 * Hyperlink (toolbar + Ctrl+K), Add Comment, and Reply to Comment (F1).
 * Electron doesn't implement `window.prompt` (it throws "prompt() is not
 * supported"), so all three previously did nothing at all in the packaged
 * app; this is what replaces that call. See DocxViewer.editor.test.tsx for
 * coverage of the three call sites actually applying/cancelling a value.
 */
import { useState } from 'react'
import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { ShortcutManagerProvider } from '../../hooks/ShortcutManagerProvider'
import { DocxPromptDialog } from '../DocxPromptDialog'

// Mirrors UnsavedChangesDialog.test.tsx's Harness — the dialog's own
// Escape-close registers through the shared shortcut dispatcher, so it
// needs a real <ShortcutManagerProvider> ancestor to exercise for real.
function Harness({ initialValue = 'https://' }: { readonly initialValue?: string }) {
  const [open, setOpen] = useState(false)
  const [confirmed, setConfirmed] = useState<string | null>(null)
  return (
    <ShortcutManagerProvider>
      <div>
        <button onClick={() => setOpen(true)}>trigger prompt dialog</button>
        <span data-testid="confirmed-value">{confirmed ?? ''}</span>
        <DocxPromptDialog
          isOpen={open}
          titleText="Insert hyperlink"
          label="URL"
          initialValue={initialValue}
          confirmLabel="Insert"
          onConfirm={(value) => {
            setConfirmed(value)
            setOpen(false)
          }}
          onCancel={() => setOpen(false)}
        />
      </div>
    </ShortcutManagerProvider>
  )
}

describe('DocxPromptDialog', () => {
  it('renders nothing when closed', () => {
    const { container } = render(
      <DocxPromptDialog
        isOpen={false}
        titleText="Insert hyperlink"
        label="URL"
        initialValue="https://"
        confirmLabel="Insert"
        onConfirm={vi.fn()}
        onCancel={vi.fn()}
      />,
    )
    expect(container).toBeEmptyDOMElement()
  })

  it('is a labeled, accessible modal with a real <label> for the input', () => {
    render(
      <DocxPromptDialog
        isOpen
        titleText="Add comment"
        label="Comment"
        initialValue=""
        confirmLabel="Add comment"
        onConfirm={vi.fn()}
        onCancel={vi.fn()}
      />,
    )

    const dialog = screen.getByRole('dialog', { name: 'Add comment' })
    expect(dialog).toHaveAttribute('aria-modal', 'true')

    // A real <label for="…"> pointing at the input's id — not a placeholder
    // standing in for one.
    const input = screen.getByLabelText('Comment')
    expect(input.tagName).toBe('INPUT')
    const label = dialog.querySelector('label')
    expect(label).not.toBeNull()
    expect(label).toHaveAttribute('for', input.id)
    expect(input).not.toHaveAttribute('placeholder')
  })

  it('prefills the input with initialValue and focuses it on open', () => {
    render(<Harness initialValue="https://" />)
    fireEvent.click(screen.getByRole('button', { name: 'trigger prompt dialog' }))

    const input = screen.getByLabelText('URL')
    expect(input).toHaveValue('https://')
    expect(input).toHaveFocus()
  })

  it('Enter (submitting the form) confirms the current value', () => {
    render(<Harness />)
    fireEvent.click(screen.getByRole('button', { name: 'trigger prompt dialog' }))

    const input = screen.getByLabelText('URL')
    fireEvent.change(input, { target: { value: 'https://example.com' } })
    fireEvent.submit(input.closest('form')!)

    expect(screen.getByTestId('confirmed-value')).toHaveTextContent('https://example.com')
    expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  })

  it('clicking the confirm button confirms the current value, including an empty one', () => {
    const onConfirm = vi.fn()
    render(
      <DocxPromptDialog
        isOpen
        titleText="Add comment"
        label="Comment"
        initialValue=""
        confirmLabel="Add comment"
        onConfirm={onConfirm}
        onCancel={vi.fn()}
      />,
    )

    // Add Comment/Reply apply even an empty value (only Cancel/Escape must
    // leave the document untouched) — matching the old `window.prompt`
    // callers, which only bailed on a `null` (cancelled) result, not on an
    // empty string.
    fireEvent.click(screen.getByRole('button', { name: 'Add comment' }))
    expect(onConfirm).toHaveBeenCalledWith('')
  })

  it('Cancel closes the dialog without confirming', () => {
    render(<Harness />)
    fireEvent.click(screen.getByRole('button', { name: 'trigger prompt dialog' }))
    fireEvent.change(screen.getByLabelText('URL'), { target: { value: 'https://example.com' } })

    fireEvent.click(screen.getByRole('button', { name: 'Cancel' }))

    expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
    expect(screen.getByTestId('confirmed-value')).toHaveTextContent('')
  })

  it('clicking the backdrop closes the dialog without confirming', () => {
    const onConfirm = vi.fn()
    const onCancel = vi.fn()
    render(
      <DocxPromptDialog
        isOpen
        titleText="Insert hyperlink"
        label="URL"
        initialValue="https://"
        confirmLabel="Insert"
        onConfirm={onConfirm}
        onCancel={onCancel}
      />,
    )

    // The backdrop is the dialog form's own parent; clicking it (not the
    // form itself, which stops propagation) must cancel.
    fireEvent.click(screen.getByRole('dialog').parentElement!)

    expect(onCancel).toHaveBeenCalledTimes(1)
    expect(onConfirm).not.toHaveBeenCalled()
  })

  it('closes on Escape without confirming', () => {
    render(<Harness />)
    fireEvent.click(screen.getByRole('button', { name: 'trigger prompt dialog' }))
    fireEvent.change(screen.getByLabelText('URL'), { target: { value: 'https://example.com' } })
    expect(screen.getByRole('dialog')).toBeInTheDocument()

    fireEvent.keyDown(window, { key: 'Escape' })

    expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
    expect(screen.getByTestId('confirmed-value')).toHaveTextContent('')
  })

  it('restores focus to the element that opened it once closed', () => {
    render(<Harness />)
    const trigger = screen.getByRole('button', { name: 'trigger prompt dialog' })
    trigger.focus()
    fireEvent.click(trigger)
    expect(screen.getByLabelText('URL')).toHaveFocus()

    fireEvent.click(screen.getByRole('button', { name: 'Cancel' }))

    expect(trigger).toHaveFocus()
  })

  it('keeps Tab from the last focusable element wrapping back to the first (focus trap), excluding a disabled control', () => {
    render(<Harness />)
    fireEvent.click(screen.getByRole('button', { name: 'trigger prompt dialog' }))

    const dialog = screen.getByRole('dialog')
    const input = screen.getByLabelText('URL')
    const insertButton = screen.getByRole('button', { name: 'Insert' })

    // A disabled control injected into the dialog (as if a future variant
    // of this dialog rendered one) must be skipped by the trap, matching
    // UnsavedChangesDialog's own "excludes disabled controls" coverage of
    // the shared useFocusTrap hook.
    const disabledProbe = document.createElement('button')
    disabledProbe.textContent = 'disabled probe'
    disabledProbe.disabled = true
    dialog.appendChild(disabledProbe)

    insertButton.focus()
    expect(insertButton).toHaveFocus()

    fireEvent.keyDown(window, { key: 'Tab' })
    expect(input).toHaveFocus()
    expect(disabledProbe).not.toHaveFocus()

    // Shift+Tab from the first (the input) must wrap back to the last real
    // (enabled) focusable — Insert, not the disabled probe.
    fireEvent.keyDown(window, { key: 'Tab', shiftKey: true })
    expect(insertButton).toHaveFocus()
  })
})
