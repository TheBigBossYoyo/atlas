import { useState, useCallback, useEffect } from 'react';

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

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent): void => {
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) return;

      if (e.key === '=' || e.key === '+') {
        e.preventDefault();
        increase();
      } else if (e.key === '-') {
        e.preventDefault();
        decrease();
      } else if (e.key === '0') {
        e.preventDefault();
        reset();
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [decrease, increase, reset]);

  return { size, increase, decrease, reset } as const;
}
