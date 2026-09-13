import { useState, useEffect, useCallback } from 'react';
import { THEMES, type Theme } from '../types';

function isTheme(value: string | null): value is Theme {
  return value !== null && THEMES.some(t => t.id === value);
}

function getInitialTheme(): Theme {
  const stored = localStorage.getItem('atlas-theme');
  if (isTheme(stored)) return stored;
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

export function useTheme() {
  const [theme, setTheme] = useState<Theme>(getInitialTheme);

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('atlas-theme', theme);
    window.electronAPI?.setTheme?.(theme as never);
  }, [theme]);

  const cycleTheme = useCallback(() => {
    setTheme(prev => {
      const currentIndex = THEMES.findIndex(t => t.id === prev);
      const nextIndex = (currentIndex + 1) % THEMES.length;
      return THEMES[nextIndex]?.id ?? 'dark';
    });
  }, []);

  return { theme, setTheme, cycleTheme } as const;
}
