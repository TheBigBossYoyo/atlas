/**
 * useRestoreFocusOnClose — shared focus-restore hook for the toolbar's
 * non-modal dropdown menus (ThemeMenu/ExportMenu/NewDocumentMenu). See the
 * hook's own doc comment for the "document.activeElement === body" heuristic
 * this relies on to distinguish "the menu closed out from under the user's
 * focus" from "the user clicked a different focusable element", which should
 * keep whatever it already claimed.
 */
import { useRef, useState } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { useRestoreFocusOnClose } from '../useRestoreFocusOnClose';

function Harness() {
  const [open, setOpen] = useState(true);
  const triggerRef = useRef<HTMLButtonElement>(null);
  useRestoreFocusOnClose(open, triggerRef);

  return (
    <div>
      <button ref={triggerRef}>trigger</button>
      <button>elsewhere</button>
      {open && (
        <div>
          <button
            onClick={() => setOpen(false)}
            data-testid="item"
          >
            item
          </button>
        </div>
      )}
    </div>
  );
}

describe('useRestoreFocusOnClose', () => {
  it('refocuses the trigger once open flips to false and the closing content took focus with it', () => {
    render(<Harness />);
    const item = screen.getByTestId('item');
    item.focus();
    expect(item).toHaveFocus();

    // Clicking removes the item from the DOM (open -> false); jsdom (like a
    // real browser) drops focus to document.body when the focused node is
    // unmounted, which is exactly the case the hook should catch.
    fireEvent.click(item);

    expect(screen.getByText('trigger')).toHaveFocus();
  });

  it('does nothing while still open', () => {
    render(<Harness />);
    const other = screen.getByText('elsewhere');
    other.focus();
    expect(other).toHaveFocus();
  });

  it('does not steal focus from an element the user already focused elsewhere', () => {
    function OutsideClickHarness() {
      const [open, setOpen] = useState(true);
      const triggerRef = useRef<HTMLButtonElement>(null);
      useRestoreFocusOnClose(open, triggerRef);
      return (
        <div>
          <button ref={triggerRef}>trigger</button>
          <button
            onClick={() => {
              // Simulates an "outside click" handler: something ELSE claims
              // focus in the same tick that closes the menu.
              setOpen(false);
            }}
          >
            elsewhere
          </button>
        </div>
      );
    }

    render(<OutsideClickHarness />);
    const elsewhere = screen.getByText('elsewhere');
    elsewhere.focus();
    fireEvent.click(elsewhere);

    // `elsewhere` was never removed from the DOM, so it keeps focus — the
    // hook must not yank it back to the trigger.
    expect(elsewhere).toHaveFocus();
  });
});
