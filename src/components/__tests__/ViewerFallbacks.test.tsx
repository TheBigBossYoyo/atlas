/**
 * Themed fallback UIs (UX-17): ViewerErrorBoundary's crash screen and
 * ViewerLoading. Both used to be unstyled — plain div/pre/button, or bare
 * "Loading X…" text — with no shared visual treatment. UnknownViewer's own
 * `.unknown-viewer` scroll-container regression is already covered by
 * src/viewers/__tests__/viewerScrollContainers.test.tsx.
 */
import { render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { ViewerErrorBoundary } from '../ViewerErrorBoundary';
import { ViewerLoading } from '../ViewerLoading';

function Boom(): never {
  throw new Error('synthetic viewer crash');
}

describe('ViewerErrorBoundary (UX-17)', () => {
  it('renders the themed fallback with the shared viewer-fallback classes when a child throws', () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});

    render(
      <ViewerErrorBoundary>
        <Boom />
      </ViewerErrorBoundary>,
    );

    const alert = screen.getByRole('alert');
    expect(alert).toHaveClass('viewer-fallback', 'viewer-error');
    expect(screen.getByText('Viewer crashed')).toBeInTheDocument();
    expect(screen.getByText('synthetic viewer crash')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Try again' })).toBeInTheDocument();

    consoleError.mockRestore();
  });

  it('renders children normally when nothing throws', () => {
    render(
      <ViewerErrorBoundary>
        <div>all good</div>
      </ViewerErrorBoundary>,
    );
    expect(screen.getByText('all good')).toBeInTheDocument();
    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
  });
});

describe('ViewerLoading (UX-17)', () => {
  it('renders the themed fallback with the shared viewer-fallback classes', () => {
    render(<ViewerLoading format="docx" />);
    const status = screen.getByRole('status');
    expect(status).toHaveClass('viewer-fallback', 'viewer-loading');
    expect(status).toHaveAttribute('aria-live', 'polite');
    expect(screen.getByText('Loading docx…')).toBeInTheDocument();
  });

  it('falls back to a generic label when no format is given', () => {
    render(<ViewerLoading />);
    expect(screen.getByText('Loading document…')).toBeInTheDocument();
  });
});
