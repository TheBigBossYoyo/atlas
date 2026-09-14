import { useState, useEffect, useCallback, useRef } from 'react';
import { THEMES, type Theme } from '../types';

const EXPLICIT_KEY = 'atlas-theme-explicit';

function isTheme(value: string | null): value is Theme {
  return value !== null && THEMES.some(t => t.id === value);
}

function systemTheme(): Theme {
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function getInitialTheme(): Theme {
  const stored = localStorage.getItem('atlas-theme');
  if (isTheme(stored)) return stored;
  return systemTheme();
}

function hasExplicitChoice(): boolean {
  try {
    return localStorage.getItem(EXPLICIT_KEY) === '1';
  } catch {
    return false;
  }
}

export function useTheme() {
  const [theme, setThemeState] = useState<Theme>(getInitialTheme);
  // Tracks whether the user has ever explicitly picked a theme (ThemeMenu or
  // Ctrl+T) — a ref so the SHELL-25 matchMedia listener below always reads
  // the latest value without needing to re-subscribe every time it changes.
  const explicitRef = useRef(hasExplicitChoice());

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('atlas-theme', theme);
    window.electronAPI?.setTheme?.(theme as never);
  }, [theme]);

  // SHELL-25 — the initial system preference was previously only read once,
  // at mount; a live OS dark/light toggle had no effect until the app
  // restarted. Reapply the system preference on every change, but only while
  // the user hasn't made an explicit theme choice of their own — once they
  // have (ThemeMenu or Ctrl+T), their choice sticks regardless of what the OS
  // does afterward.
  useEffect(() => {
    const mql = window.matchMedia('(prefers-color-scheme: dark)');
    const onChange = () => {
      if (explicitRef.current) return;
      setThemeState(systemTheme());
    };
    mql.addEventListener('change', onChange);
    return () => mql.removeEventListener('change', onChange);
  }, []);

  const markExplicit = useCallback(() => {
    explicitRef.current = true;
    try {
      localStorage.setItem(EXPLICIT_KEY, '1');
    } catch {
      // localStorage unavailable (private mode, etc.) — the in-memory ref
      // still correctly stops this session's live system-theme updates.
    }
  }, []);

  const setTheme = useCallback((next: Theme) => {
    markExplicit();
    setThemeState(next);
  }, [markExplicit]);

  const cycleTheme = useCallback(() => {
    markExplicit();
    setThemeState(prev => {
      const currentIndex = THEMES.findIndex(t => t.id === prev);
      const nextIndex = (currentIndex + 1) % THEMES.length;
      return THEMES[nextIndex]?.id ?? 'dark';
    });
  }, [markExplicit]);

  return { theme, setTheme, cycleTheme } as const;
}
