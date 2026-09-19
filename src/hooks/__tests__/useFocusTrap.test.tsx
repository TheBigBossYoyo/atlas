/**
 * useFocusTrap — shared Tab-containment + focus-restore hook (A11Y-2),
 * factored out of the pattern ShortcutsModal/UnsavedChangesDialog each
 * hand-rolled (UX-13) so SearchOverlay and PdfPasswordDialog could reuse it
 * without copy-pasting another ~25-line effect.
 */
import { useRef, useState } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { useFocusTrap } from '../useFocusTrap';

function Harness({ focusOnOpen }: { focusOnOpen?: boolean } = {}) {
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);
  useFocusTrap(containerRef, open, { focusOnOpen });

  return (
    <div>
      <button onClick={() => setOpen(true)}>open trigger</button>
      {open && (
        <div ref={containerRef}>
          <button>first</button>
          <button disabled>disabled middle</button>
          <button onClick={() => setOpen(false)}>last</button>
        </div>
      )}
    </div>
  );
}

describe('useFocusTrap', () => {
  it('moves focus to the first focusable element inside the container when it opens (default)', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'open trigger' }));

    expect(screen.getByRole('button', { name: 'first' })).toHaveFocus();
  });

  it('does not move focus on open when focusOnOpen is false', () => {
    render(<Harness focusOnOpen={false} />);
    const trigger = screen.getByRole('button', { name: 'open trigger' });
    trigger.focus();
    fireEvent.click(trigger);

    // Nothing inside the (now-rendered) container should have stolen focus.
    expect(screen.getByRole('button', { name: 'first' })).not.toHaveFocus();
  });

  it('wraps Tab from the last focusable element back to the first, skipping disabled ones', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'open trigger' }));
    screen.getByRole('button', { name: 'last' }).focus();

    fireEvent.keyDown(window, { key: 'Tab' });

    expect(screen.getByRole('button', { name: 'first' })).toHaveFocus();
  });

  it('wraps Shift+Tab from the first focusable element back to the last, skipping disabled ones', () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole('button', { name: 'open trigger' }));
    screen.getByRole('button', { name: 'first' }).focus();

    fireEvent.keyDown(window, { key: 'Tab', shiftKey: true });

    expect(screen.getByRole('button', { name: 'last' })).toHaveFocus();
  });

  it('restores focus to whatever triggered it once isOpen flips back to false', () => {
    render(<Harness />);
    const trigger = screen.getByRole('button', { name: 'open trigger' });
    trigger.focus();
    fireEvent.click(trigger);
    expect(screen.getByRole('button', { name: 'first' })).toHaveFocus();

    fireEvent.click(screen.getByRole('button', { name: 'last' }));

    expect(trigger).toHaveFocus();
  });

  it('does nothing while closed', () => {
    render(<Harness />);
    expect(screen.queryByRole('button', { name: 'first' })).not.toBeInTheDocument();
  });
});
