import { useState, useEffect, useCallback, useRef } from 'react';
import { THEMES, type Theme } from '../types';

const THEME_KEY = 'atlas-theme';
const EXPLICIT_KEY = 'atlas-theme-explicit';

function isTheme(value: string | null): value is Theme {
  return value !== null && THEMES.some(t => t.id === value);
}

function systemTheme(): Theme {
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function getInitialTheme(): Theme {
  const stored = localStorage.getItem(THEME_KEY);
  if (isTheme(stored)) return stored;
  return systemTheme();
}

/**
 * Migration (wave 3 shell-polish follow-up to SHELL-25): an install that
 * already had a persisted theme before `EXPLICIT_KEY` existed must keep
 * behaving the way it always did — sticky across launches, never silently
 * re-derived from a live OS change — even though a stored value alone can no
 * longer tell us whether it came from a deliberate ThemeMenu/Ctrl+T pick or
 * just the very first cold-start OS read. Without this, every upgraded
 * install would start unexpectedly retheming itself on the next OS light/
 * dark toggle, which is a bigger surprise than the smaller risk of also
 * grandfathering a same-version install that never touched the theme picker
 * before its first relaunch. A brand-new install (no stored theme at all
 * yet) is unaffected and still follows the OS live until the user actually
 * picks one.
 */
function hasExplicitChoice(): boolean {
  try {
    if (localStorage.getItem(EXPLICIT_KEY) === '1') return true;
    if (localStorage.getItem(THEME_KEY) !== null) {
      localStorage.setItem(EXPLICIT_KEY, '1');
      return true;
    }
    return false;
  } catch {
    return false;
  }
}

export function useTheme() {
  const [theme, setThemeState] = useState<Theme>(getInitialTheme);
  // Tracks whether the user has ever explicitly picked a theme (ThemeMenu or
  // Ctrl+T) — a ref so the SHELL-25 matchMedia listener below always reads
  // the latest value without needing to re-subscribe every time it changes.
  //
  // Lazily initialized (`useRef(undefined)` + a render-time backfill) rather
  // than `useRef(hasExplicitChoice())`: that argument is a plain expression,
  // re-evaluated on EVERY render (React only uses its value on the very
  // first one) — and `hasExplicitChoice()` is side-effecting, persisting
  // `EXPLICIT_KEY` once it sees a stored theme. The mount effect below
  // writes `THEME_KEY` right after the first render, so the SECOND render
  // for any reason at all (an OS-driven theme change included) would call
  // hasExplicitChoice() again, see the just-written THEME_KEY, and
  // permanently mark the install "explicit" — breaking live OS-following
  // for brand-new installs too, not just migrating existing ones. Computing
  // it once here, only on the render that actually initializes the ref,
  // keeps the migration's side effect from firing on unrelated re-renders.
  const explicitRef = useRef<boolean | undefined>(undefined);
  if (explicitRef.current === undefined) {
    explicitRef.current = hasExplicitChoice();
  }

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem(THEME_KEY, theme);
    window.electronAPI?.setTheme?.(theme);
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
