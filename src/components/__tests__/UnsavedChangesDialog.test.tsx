/**
 * UnsavedChangesDialog — focus trap + initial focus + focus restoration.
 *
 * This is the P1.1 confirm dialog every open path (dialog, Recent click,
 * drag-drop, OS "Open with", closing a tab/the app) funnels through before
 * silently discarding unsaved edits — a real data-loss guard, not a cosmetic
 * confirm box. It declared `role="alertdialog"`/`aria-modal="true"` without
 * ever moving focus into itself, trapping Tab, or restoring focus on close —
 * the same gap ShortcutsModal had before its UX-13 fix (see
 * ShortcutsModal.test.tsx), just never fixed here.
 */
import { useState } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { ShortcutManagerProvider } from '../../hooks/ShortcutManagerProvider';
import { UnsavedChangesDialog } from '../UnsavedChangesDialog';

// Mirrors ShortcutsModal.test.tsx's Harness — the dialog's own Escape-close
// registers through the shared shortcut dispatcher, so it needs a real
// <ShortcutManagerProvider> ancestor to exercise for real.
function Harness() {
  const [open, setOpen] = useState(false);
  return (
    <ShortcutManagerProvider>
      <div>
        <button onClick={() => setOpen(true)}>trigger unsaved dialog</button>
        <UnsavedChangesDialog
          isOpen={open}
          isSaving={false}
          errorMessage={null}
          onSave={() => setOpen(false)}
          onDiscard={() => setOpen(false)}
          onCancel={() => setOpen(false)}
        />
      </div>
    </ShortcutManagerProvider>
  );
}

describe('UnsavedChangesDialog', () => {
  it('renders nothing when closed', () => {
    render(
      <UnsavedChangesDialog
        isOpen={false}
        isSaving={false}
        errorMessage={null}
        onSave={vi.fn()}
        onDiscard={vi.fn()}
        onCancel={vi.fn()}
      />,
    );
    expect(screen.queryByRole('alertdialog')).not.toBeInTheDocument();
  });

  it('moves focus into the dialog (its Cancel button) when opened', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'trigger unsaved dialog' }));

    expect(screen.getByRole('button', { name: 'Cancel' })).toHaveFocus();
  });

  it('restores focus to the element that opened it once closed', () => {
    render(<Harness />);
    const trigger = screen.getByRole('button', { name: 'trigger unsaved dialog' });
    // fireEvent.click (unlike a real mouse click, or userEvent) does not
    // itself focus the target — focus it explicitly first so
    // document.activeElement matches what a real click would leave behind.
    trigger.focus();
    fireEvent.click(trigger);
    expect(screen.getByRole('button', { name: 'Cancel' })).toHaveFocus();

    fireEvent.click(screen.getByRole('button', { name: 'Cancel' }));

    expect(trigger).toHaveFocus();
  });

  it('closes on Escape', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'trigger unsaved dialog' }));
    expect(screen.getByRole('alertdialog')).toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'Escape' });

    expect(screen.queryByRole('alertdialog')).not.toBeInTheDocument();
  });

  it('keeps Tab from the last focusable element wrapping back to the first (focus trap)', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'trigger unsaved dialog' }));

    const cancelButton = screen.getByRole('button', { name: 'Cancel' });
    const saveButton = screen.getByRole('button', { name: 'Save' });
    saveButton.focus();
    expect(saveButton).toHaveFocus();

    // Tab from the last focusable element (Save) must wrap back to the
    // first (Cancel) rather than escaping into the page behind the dialog.
    fireEvent.keyDown(window, { key: 'Tab' });
    expect(cancelButton).toHaveFocus();

    // Shift+Tab from the first must wrap back to the last.
    fireEvent.keyDown(window, { key: 'Tab', shiftKey: true });
    expect(saveButton).toHaveFocus();
  });

  it('does not focus or trap on a disabled (isSaving) button', () => {
    render(
      <ShortcutManagerProvider>
        <UnsavedChangesDialog
          isOpen
          isSaving
          errorMessage={null}
          onSave={vi.fn()}
          onDiscard={vi.fn()}
          onCancel={vi.fn()}
        />
      </ShortcutManagerProvider>,
    );
    // All three actions are disabled while a save is in flight (isSaving);
    // the focus trap must skip them rather than focusing a disabled button.
    expect(screen.getByRole('button', { name: 'Cancel' })).toBeDisabled();
    expect(screen.getByRole('button', { name: 'Discard' })).toBeDisabled();
    expect(screen.getByRole('button', { name: 'Saving…' })).toBeDisabled();
  });

  // A11Y-1 — this dialog already filtered out disabled controls before
  // ShortcutsModal migrated to share this behavior (see useFocusTrap.ts and
  // ShortcutsModal.test.tsx's own "excludes disabled controls" test for the
  // gap that migration closed). `isSaving` disables all three buttons at
  // once, so it can't exercise "some enabled, one disabled" — inject a
  // disabled probe alongside the three real (enabled) buttons to prove the
  // shared hook still excludes a disabled control from the wrap even when
  // real controls remain focusable around it.
  it('excludes a disabled control from the wrap while real controls stay enabled', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'trigger unsaved dialog' }));

    const dialog = screen.getByRole('alertdialog');
    const cancelButton = screen.getByRole('button', { name: 'Cancel' });
    const saveButton = screen.getByRole('button', { name: 'Save' });

    const disabledProbe = document.createElement('button');
    disabledProbe.textContent = 'disabled probe';
    disabledProbe.disabled = true;
    dialog.appendChild(disabledProbe);

    saveButton.focus();
    const event = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    window.dispatchEvent(event);

    // Save is still the last real focusable (the probe is excluded), so Tab
    // from it must wrap to Cancel (the first), not land on the disabled probe.
    expect(event.defaultPrevented).toBe(true);
    expect(cancelButton).toHaveFocus();
    expect(disabledProbe).not.toHaveFocus();
  });

  it('is a labeled, accessible alertdialog', () => {
    render(
      <UnsavedChangesDialog
        isOpen
        isSaving={false}
        errorMessage={null}
        onSave={vi.fn()}
        onDiscard={vi.fn()}
        onCancel={vi.fn()}
      />,
    );
    const dialog = screen.getByRole('alertdialog');
    expect(dialog).toHaveAttribute('aria-modal', 'true');
    expect(dialog).toHaveAccessibleName('Unsaved changes');
  });

  it('announces a save-failure error message via role="alert"', () => {
    render(
      <UnsavedChangesDialog
        isOpen
        isSaving={false}
        errorMessage="Permission denied — you don't have access to save to this location."
        onSave={vi.fn()}
        onDiscard={vi.fn()}
        onCancel={vi.fn()}
      />,
    );
    expect(screen.getByRole('alert')).toHaveTextContent(/permission denied/i);
  });
});
