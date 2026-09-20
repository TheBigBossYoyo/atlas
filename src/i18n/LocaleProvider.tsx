import { useCallback, useEffect, useMemo, useState, type ReactNode } from 'react';
import type { LocalePreference, MessageParams, TranslateFn } from './types';
import { getStoredPreference, storePreference, resolveLocale } from './locale';
import { translate } from './translate';
import { setActiveLocale } from './activeLocale';
import { LocaleContext, type LocaleContextValue } from './context';

/**
 * Wraps the shell (see `App.tsx`) so every component below it can call
 * `useTranslate()`. Mirrors `useTheme.ts`'s persistence pattern: the
 * preference is read from `localStorage` on mount and written back on every
 * change. Unlike the theme, the DEFAULT preference is the literal `'en'`
 * (see `locale.ts`'s doc comment) — only an explicit `'system'` choice ever
 * asks the main process for `app.getLocale()`.
 */
export function LocaleProvider({ children }: { children: ReactNode }) {
  const [preference, setPreferenceState] = useState<LocalePreference>(getStoredPreference);
  const [osTag, setOsTag] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    window.electronAPI?.getLocale?.()
      .then((tag) => {
        if (!cancelled) setOsTag(tag);
      })
      .catch(() => {
        // No electronAPI (plain browser tab) or the IPC call failed — a
        // 'system' preference simply falls back to English (see
        // `resolveLocale`) until/unless it resolves.
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const locale = resolveLocale(preference, osTag);

  useEffect(() => {
    document.documentElement.lang = locale;
    setActiveLocale(locale);
  }, [locale]);

  const setPreference = useCallback((next: LocalePreference) => {
    storePreference(next);
    setPreferenceState(next);
  }, []);

  const t = useCallback<TranslateFn>(
    (key, params?: MessageParams) => translate(locale, key, params),
    [locale],
  );

  const value = useMemo<LocaleContextValue>(
    () => ({ locale, preference, setPreference, t }),
    [locale, preference, setPreference, t],
  );

  return <LocaleContext.Provider value={value}>{children}</LocaleContext.Provider>;
}
