import { AlertCircle, CheckCircle2, Info, X } from 'lucide-react';
import type { ToastItem } from './ToastContext';

interface ToastViewportProps {
  readonly toasts: readonly ToastItem[];
  readonly onDismiss: (id: string) => void;
}

const VARIANT_ICON = { error: AlertCircle, success: CheckCircle2, info: Info } as const;

/**
 * Renders the stack of active toasts, fixed to a corner of the viewport.
 * Error toasts get `role="alert"`/`aria-live="assertive"` (announced
 * immediately, replacing the native `alert()` they stand in for); success/
 * info toasts use the calmer `role="status"`/`aria-live="polite"`.
 */
export function ToastViewport({ toasts, onDismiss }: ToastViewportProps) {
  if (toasts.length === 0) return null;

  return (
    <div className="toast-viewport" aria-label="Notifications">
      {toasts.map(({ id, message, variant }) => {
        const Icon = VARIANT_ICON[variant];
        return (
          <div
            key={id}
            className={`toast toast--${variant}`}
            role={variant === 'error' ? 'alert' : 'status'}
            aria-live={variant === 'error' ? 'assertive' : 'polite'}
          >
            <Icon size={16} className="toast__icon" aria-hidden="true" />
            <span className="toast__message">{message}</span>
            <button className="toast__close" onClick={() => onDismiss(id)} aria-label="Dismiss notification">
              <X size={14} />
            </button>
          </div>
        );
      })}
    </div>
  );
}
