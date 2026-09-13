import { AlertCircle, X } from 'lucide-react';

interface FileStatusBannerProps {
  loading: boolean;
  error: string | null;
  onDismissError: () => void;
}

/**
 * P2.3/SHELL-07/LOAD-06/RUN-17 — `useFileHandler`'s `loading`/`error` state
 * was already implemented and unit-tested but never rendered, so a failed
 * open (deleted/locked/corrupted file, a stale Recent entry) looked like
 * nothing happened. Renders a thin top-of-content loading bar while a load
 * is in flight and a dismissible error banner when one fails.
 */
export function FileStatusBanner({ loading, error, onDismissError }: FileStatusBannerProps) {
  if (!loading && !error) return null;

  return (
    <div className="file-status">
      {loading && (
        <div className="file-status__loading" role="status" aria-live="polite">
          <span className="file-status__loading-bar" aria-hidden="true" />
          <span className="file-status__loading-label">Opening file…</span>
        </div>
      )}
      {error && (
        <div className="file-status__error" role="alert">
          <AlertCircle size={16} aria-hidden="true" />
          <span className="file-status__error-message">{error}</span>
          <button
            className="file-status__error-dismiss"
            onClick={onDismissError}
            aria-label="Dismiss error"
            title="Dismiss"
          >
            <X size={14} />
          </button>
        </div>
      )}
    </div>
  );
}
