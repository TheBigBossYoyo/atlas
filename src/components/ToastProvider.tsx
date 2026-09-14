import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ToastContext, type ToastItem, type ToastVariant } from './ToastContext';
import { ToastViewport } from './Toast';

const AUTO_DISMISS_MS: Record<ToastVariant, number> = {
  error: 8000,
  success: 4000,
  info: 4000,
};

let nextToastId = 0;

interface ToastProviderProps {
  readonly children: React.ReactNode;
}

/**
 * UX-18 — app-wide toast notification system, replacing every blocking
 * native `alert()` (export failures, save failures, ...) with a themed,
 * auto-dismissing, non-blocking notification. Wrap the app root with this
 * once; any descendant calls `useToast()` to queue one.
 */
export function ToastProvider({ children }: ToastProviderProps) {
  const [toasts, setToasts] = useState<readonly ToastItem[]>([]);
  const timersRef = useRef<Map<string, number>>(new Map());

  const dismissToast = useCallback((id: string) => {
    setToasts(prev => prev.filter(t => t.id !== id));
    const timer = timersRef.current.get(id);
    if (timer !== undefined) {
      window.clearTimeout(timer);
      timersRef.current.delete(id);
    }
  }, []);

  const showToast = useCallback((message: string, variant: ToastVariant = 'info') => {
    nextToastId += 1;
    const id = `toast-${nextToastId}`;
    setToasts(prev => [...prev, { id, message, variant }]);
    const timer = window.setTimeout(() => dismissToast(id), AUTO_DISMISS_MS[variant]);
    timersRef.current.set(id, timer);
  }, [dismissToast]);

  useEffect(() => {
    const timers = timersRef.current;
    return () => {
      for (const timer of timers.values()) window.clearTimeout(timer);
      timers.clear();
    };
  }, []);

  const value = useMemo(() => ({ toasts, showToast, dismissToast }), [toasts, showToast, dismissToast]);

  return (
    <ToastContext.Provider value={value}>
      {children}
      <ToastViewport toasts={toasts} onDismiss={dismissToast} />
    </ToastContext.Provider>
  );
}
