import { useContext } from 'react';
import { ToastContext, type ToastVariant } from '../components/ToastContext';

/**
 * Returns a stable `showToast(message, variant?)` function that queues a
 * themed, auto-dismissing notification (UX-18 — replaces every blocking,
 * theme-ignoring `alert()` in the app). Must be called from a component
 * rendered inside `<ToastProvider>`.
 */
export function useToast(): (message: string, variant?: ToastVariant) => void {
  const ctx = useContext(ToastContext);
  if (!ctx) {
    throw new Error('useToast: missing <ToastProvider>');
  }
  return ctx.showToast;
}
