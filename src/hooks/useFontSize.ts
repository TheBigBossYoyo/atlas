import { useState, useCallback, useEffect } from 'react';
import { useShellShortcut } from './useShortcutManager';
import type { ShortcutHandler } from './shortcutManagerContext';

const KEY = 'atlas-font-size';
const MIN = 12;
const MAX = 24;
const DEFAULT = 15;

function clampSize(value: number): number {
  return Math.min(MAX, Math.max(MIN, value));
}

function getInitialSize(): number {
  try {
    const stored = localStorage.getItem(KEY);
    if (stored == null) return DEFAULT;
    const parsed = Number(stored);
    return Number.isNaN(parsed) ? DEFAULT : clampSize(parsed);
  } catch {
    return DEFAULT;
  }
}

export function useFontSize() {
  const [size, setSize] = useState<number>(getInitialSize);

  useEffect(() => {
    document.documentElement.style.setProperty('--md-font-size', `${size}px`);
    localStorage.setItem(KEY, String(size));
  }, [size]);

  const increase = useCallback(() => {
    setSize(current => Math.min(MAX, current + 1));
  }, []);

  const decrease = useCallback(() => {
    setSize(current => Math.max(MIN, current - 1));
  }, []);

  const reset = useCallback(() => {
    setSize(DEFAULT);
  }, []);

  // P2.1 — shell-global tier: skipped while a plain field/contentEditable has
  // focus (e.g. Ctrl+0 while editing a DOCX no longer resets *markdown's*
  // font size out from under it — the same collision class as SHELL-08/09,
  // just never independently ticketed).
  const handler = useCallback<ShortcutHandler>(
    (e, ctx) => {
      if (!(e.ctrlKey || e.metaKey)) return false;
      if (ctx.inPlainField) return false;

      if (e.key === '=' || e.key === '+') {
        e.preventDefault();
        increase();
        return true;
      }
      if (e.key === '-') {
        e.preventDefault();
        decrease();
        return true;
      }
      if (e.key === '0') {
        e.preventDefault();
        reset();
        return true;
      }
      return false;
    },
    [decrease, increase, reset],
  );

  useShellShortcut(handler);

  return { size, increase, decrease, reset } as const;
}
