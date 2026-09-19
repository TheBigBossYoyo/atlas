/**
 * SearchOverlay — aria-live match count and visible focus (UX-15), plus the
 * Tab focus trap + focus restore added for A11Y-2.
 */
import { useState } from 'react';
import type { ComponentProps } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { SearchOverlay } from '../SearchOverlay';

function renderOverlay(overrides: Partial<ComponentProps<typeof SearchOverlay>> = {}) {
  return render(
    <SearchOverlay
      isOpen={true}
      query="foo"
      matchCount={3}
      currentMatch={0}
      onQueryChange={vi.fn()}
      onNext={vi.fn()}
      onPrev={vi.fn()}
      onClose={vi.fn()}
      {...overrides}
    />,
  );
}

// A11Y-2 — SearchOverlay's isOpen is controlled by its parent (Ctrl+F opens
// it, the close button or Escape closes it), so exercising open/close focus
// behavior needs a harness that owns that state, mirroring how App.tsx does.
function Harness() {
  const [open, setOpen] = useState(false);
  return (
    <div>
      <button onClick={() => setOpen(true)}>open search</button>
      <SearchOverlay
        isOpen={open}
        query="foo"
        matchCount={3}
        currentMatch={0}
        onQueryChange={() => {}}
        onNext={() => {}}
        onPrev={() => {}}
        onClose={() => setOpen(false)}
      />
    </div>
  );
}

describe('SearchOverlay (UX-15)', () => {
  it('announces the match count via aria-live="polite"', () => {
    renderOverlay({ matchCount: 3, currentMatch: 1 });
    const status = screen.getByRole('status');
    expect(status).toHaveAttribute('aria-live', 'polite');
    expect(status).toHaveTextContent('2 of 3');
  });

  it('announces "No results" when nothing matches', () => {
    renderOverlay({ matchCount: 0 });
    expect(screen.getByRole('status')).toHaveTextContent('No results');
  });

  it('renders nothing (no match-count element) when the query is empty', () => {
    renderOverlay({ query: '' });
    expect(screen.queryByRole('status')).not.toBeInTheDocument();
  });

  it('renders nothing at all when closed', () => {
    renderOverlay({ isOpen: false });
    expect(screen.queryByPlaceholderText('Search in document...')).not.toBeInTheDocument();
  });

  it('keeps Tab from the last focusable element wrapping back to the first (focus trap, A11Y-2)', () => {
    renderOverlay();
    const closeButton = screen.getByRole('button', { name: 'Close search' });
    closeButton.focus();
    expect(closeButton).toHaveFocus();

    fireEvent.keyDown(window, { key: 'Tab' });

    expect(screen.getByPlaceholderText('Search in document...')).toHaveFocus();
  });

  it('keeps Shift+Tab from the first focusable element wrapping back to the last (A11Y-2)', () => {
    renderOverlay();
    const input = screen.getByPlaceholderText('Search in document...');
    input.focus();
    expect(input).toHaveFocus();

    fireEvent.keyDown(window, { key: 'Tab', shiftKey: true });

    expect(screen.getByRole('button', { name: 'Close search' })).toHaveFocus();
  });

  it('restores focus to the element that opened it once closed (A11Y-2)', () => {
    render(<Harness />);
    const trigger = screen.getByRole('button', { name: 'open search' });
    trigger.focus();
    fireEvent.click(trigger);

    expect(screen.getByRole('button', { name: 'Close search' })).toBeInTheDocument();

    fireEvent.click(screen.getByRole('button', { name: 'Close search' }));

    expect(trigger).toHaveFocus();
  });
});
