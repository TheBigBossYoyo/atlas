/**
 * ShortcutsModal — focus trap + initial focus + focus restoration (UX-13).
 *
 * The modal previously declared `aria-modal="true"` without actually
 * trapping focus, moving it on open, or restoring it on close — a
 * keyboard/screen-reader user tabbing past the dialog escaped straight back
 * into the page behind it.
 */
import { useState } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { ShortcutManagerProvider } from '../../hooks/ShortcutManagerProvider';
import { ShortcutsModal } from '../ShortcutsModal';

// P2.1 — ShortcutsModal's own Escape-close now registers through the shared
// dispatcher (useShellShortcut) instead of an ad hoc `window` listener, so
// exercising it for real requires a real <ShortcutManagerProvider> ancestor
// (without one, the registration silently no-ops — see useShortcutManager.ts).
function Harness() {
  const [open, setOpen] = useState(false);
  return (
    <ShortcutManagerProvider>
      <div>
        <button onClick={() => setOpen(true)}>open shortcuts</button>
        <ShortcutsModal isOpen={open} onClose={() => setOpen(false)} />
      </div>
    </ShortcutManagerProvider>
  );
}

describe('ShortcutsModal (UX-13)', () => {
  it('renders nothing when closed', () => {
    render(<ShortcutsModal isOpen={false} onClose={() => {}} />);
    expect(screen.queryByRole('dialog')).not.toBeInTheDocument();
  });

  it('moves focus into the modal (its close button) when opened', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'open shortcuts' }));

    expect(screen.getByRole('button', { name: 'Close' })).toHaveFocus();
  });

  it('restores focus to the element that opened it once closed', () => {
    render(<Harness />);
    const trigger = screen.getByRole('button', { name: 'open shortcuts' });
    // fireEvent.click (unlike a real mouse click, or userEvent) does not
    // itself focus the target — focus it explicitly first so
    // document.activeElement matches what a real click would leave behind
    // for the modal's open effect to capture as "what to restore focus to".
    trigger.focus();
    fireEvent.click(trigger);
    expect(screen.getByRole('button', { name: 'Close' })).toHaveFocus();

    fireEvent.click(screen.getByRole('button', { name: 'Close' }));

    expect(trigger).toHaveFocus();
  });

  it('closes on Escape', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'open shortcuts' }));
    expect(screen.getByRole('dialog')).toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'Escape' });

    expect(screen.queryByRole('dialog')).not.toBeInTheDocument();
  });

  it('keeps Tab from the last focusable element wrapping back to the first (focus trap)', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'open shortcuts' }));
    const closeButton = screen.getByRole('button', { name: 'Close' });
    expect(closeButton).toHaveFocus();

    // The close button is the modal's only focusable element, so Tab from
    // it must wrap right back to itself rather than escaping to the page.
    fireEvent.keyDown(window, { key: 'Tab' });

    expect(closeButton).toHaveFocus();
  });

  // A11Y-1 — this modal's own hand-rolled trap (before its migration to the
  // shared useFocusTrap hook) never filtered out disabled elements when
  // collecting focusable candidates, unlike UnsavedChangesDialog's. That was
  // latent because the modal ships only one focusable control; it becomes a
  // real escape-from-trap the moment a disabled one exists alongside it. The
  // real markup has no disabled control today, so this injects one to prove
  // the shared hook (now the only trap implementation) actually excludes it.
  it('excludes disabled controls when trapping Tab', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'open shortcuts' }));
    const dialog = screen.getByRole('dialog');
    const closeButton = screen.getByRole('button', { name: 'Close' });

    const disabledProbe = document.createElement('button');
    disabledProbe.textContent = 'disabled probe';
    disabledProbe.disabled = true;
    dialog.appendChild(disabledProbe);

    expect(closeButton).toHaveFocus();

    const event = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    window.dispatchEvent(event);

    // With the disabled probe excluded, the close button is both "first"
    // and "last", so Tab from it must wrap straight back to itself — and
    // the trap must actually have intercepted the key (defaultPrevented) to
    // prove it didn't just silently do nothing, which is what would let Tab
    // escape into the page once a real disabled control existed here.
    expect(event.defaultPrevented).toBe(true);
    expect(closeButton).toHaveFocus();
    expect(disabledProbe).not.toHaveFocus();
  });

  it('is a labeled, accessible dialog', () => {
    render(<ShortcutsModal isOpen onClose={() => {}} />);
    const dialog = screen.getByRole('dialog');
    expect(dialog).toHaveAttribute('aria-modal', 'true');
    expect(dialog).toHaveAccessibleName('Keyboard Shortcuts');
  });

  // Regression for a doc/behavior mismatch: `useUniversalShortcuts.ts`
  // (Ctrl+N -> openNewMenu, Ctrl+Shift+T -> reopenClosedSession per SHELL-17)
  // both actually work, but this modal — the app's own "authoritative list"
  // per README.md — omitted the Ctrl+Shift+T row entirely, so there was no
  // way for a keyboard user to discover it short of reading source.
  it('documents every shell shortcut useUniversalShortcuts.ts actually implements', () => {
    render(<ShortcutsModal isOpen onClose={() => {}} />);
    expect(screen.getByText('New document menu')).toBeInTheDocument();
    expect(screen.getByText('Reopen last closed document')).toBeInTheDocument();
  });

  // Same class of gap, found again: App.tsx registers Ctrl+Tab/Ctrl+Shift+Tab
  // directly (SHELL-17's tab-cycling shortcut, not routed through
  // useUniversalShortcuts.ts), and this modal omitted it too — nothing here
  // is discoverable by a keyboard user unless it's actually listed.
  it('documents the Ctrl+Tab / Ctrl+Shift+Tab tab-cycling shortcut App.tsx implements', () => {
    render(<ShortcutsModal isOpen onClose={() => {}} />);
    expect(screen.getByText('Next / previous tab')).toBeInTheDocument();
  });
});
