/**
 * SearchOverlay — aria-live match count and visible focus (UX-15).
 */
import type { ComponentProps } from 'react';
import { render, screen } from '@testing-library/react';
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
});
