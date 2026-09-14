/**
 * Toast notification system (UX-18) — replaces every blocking, theme-ignoring
 * `alert()` with a themed, auto-dismissing, non-blocking notification.
 */
import { act, fireEvent, render, screen, within } from '@testing-library/react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ToastProvider } from '../ToastProvider';
import { useToast } from '../../hooks/useToast';

function ShowToastButton({ message, variant }: { message: string; variant?: 'error' | 'success' | 'info' }) {
  const showToast = useToast();
  return (
    <button onClick={() => showToast(message, variant)}>show</button>
  );
}

beforeEach(() => {
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
});

describe('useToast / ToastProvider', () => {
  it('throws a clear error when used outside <ToastProvider>', () => {
    // Swallow the expected React error-boundary console noise for this one
    // deliberately-invalid-usage assertion.
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});
    function Broken() {
      useToast();
      return null;
    }
    expect(() => render(<Broken />)).toThrow('useToast: missing <ToastProvider>');
    consoleError.mockRestore();
  });

  it('renders nothing when there are no active toasts', () => {
    render(<ToastProvider><div>app</div></ToastProvider>);
    expect(screen.queryByLabelText('Notifications')).not.toBeInTheDocument();
  });

  it('queues and displays a toast when showToast is called', () => {
    render(
      <ToastProvider>
        <ShowToastButton message="Export failed: disk full" variant="error" />
      </ToastProvider>,
    );

    act(() => { fireEvent.click(screen.getByRole('button', { name: 'show' })); });

    expect(screen.getByText('Export failed: disk full')).toBeInTheDocument();
  });

  it('gives an error toast role="alert"/aria-live="assertive" (replacing alert())', () => {
    render(
      <ToastProvider>
        <ShowToastButton message="boom" variant="error" />
      </ToastProvider>,
    );
    act(() => { fireEvent.click(screen.getByRole('button', { name: 'show' })); });

    const toast = screen.getByRole('alert');
    expect(toast).toHaveAttribute('aria-live', 'assertive');
    expect(within(toast).getByText('boom')).toBeInTheDocument();
  });

  it('gives an info/success toast the calmer role="status"/aria-live="polite"', () => {
    render(
      <ToastProvider>
        <ShowToastButton message="Saved" variant="success" />
      </ToastProvider>,
    );
    act(() => { fireEvent.click(screen.getByRole('button', { name: 'show' })); });

    const toast = screen.getByRole('status');
    expect(toast).toHaveAttribute('aria-live', 'polite');
  });

  it('dismisses a toast when its close button is clicked', () => {
    render(
      <ToastProvider>
        <ShowToastButton message="dismiss me" variant="info" />
      </ToastProvider>,
    );
    act(() => { fireEvent.click(screen.getByRole('button', { name: 'show' })); });
    expect(screen.getByText('dismiss me')).toBeInTheDocument();

    act(() => { fireEvent.click(screen.getByRole('button', { name: 'Dismiss notification' })); });

    expect(screen.queryByText('dismiss me')).not.toBeInTheDocument();
  });

  it('auto-dismisses an error toast after its timeout', () => {
    render(
      <ToastProvider>
        <ShowToastButton message="auto-dismiss error" variant="error" />
      </ToastProvider>,
    );
    act(() => { fireEvent.click(screen.getByRole('button', { name: 'show' })); });
    expect(screen.getByText('auto-dismiss error')).toBeInTheDocument();

    act(() => { vi.advanceTimersByTime(8000); });

    expect(screen.queryByText('auto-dismiss error')).not.toBeInTheDocument();
  });

  it('does not auto-dismiss a success toast before its (shorter) timeout elapses', () => {
    render(
      <ToastProvider>
        <ShowToastButton message="not yet" variant="success" />
      </ToastProvider>,
    );
    act(() => { fireEvent.click(screen.getByRole('button', { name: 'show' })); });

    act(() => { vi.advanceTimersByTime(3000); });
    expect(screen.getByText('not yet')).toBeInTheDocument();

    act(() => { vi.advanceTimersByTime(1000); });
    expect(screen.queryByText('not yet')).not.toBeInTheDocument();
  });

  it('stacks multiple toasts independently', () => {
    render(
      <ToastProvider>
        <ShowToastButton message="first" variant="info" />
        <ShowToastButton message="second" variant="info" />
      </ToastProvider>,
    );

    const buttons = screen.getAllByRole('button', { name: 'show' });
    act(() => { fireEvent.click(buttons[0]!); });
    act(() => { fireEvent.click(buttons[1]!); });

    expect(screen.getByText('first')).toBeInTheDocument();
    expect(screen.getByText('second')).toBeInTheDocument();
  });
});
