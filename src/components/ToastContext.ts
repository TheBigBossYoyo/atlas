/**
 * Toast notification system (UX-18) — context + types.
 *
 * Kept separate from the Provider component and the `useToast` hook so that
 * `ToastProvider.tsx` only exports a component and `useToast.ts` only
 * exports a hook (satisfies `react-refresh/only-export-components`), the
 * same split already used for `ViewerContext`.
 */
import { createContext } from 'react';

export type ToastVariant = 'error' | 'success' | 'info';

export interface ToastItem {
  readonly id: string;
  readonly message: string;
  readonly variant: ToastVariant;
}

export interface ToastContextValue {
  readonly toasts: readonly ToastItem[];
  readonly showToast: (message: string, variant?: ToastVariant) => void;
  readonly dismissToast: (id: string) => void;
}

export const ToastContext = createContext<ToastContextValue | null>(null);
