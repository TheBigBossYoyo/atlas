import type { Locale } from './types';
import { getStoredPreference, resolveLocale } from './locale';

/**
 * A small module-level mirror of `LocaleProvider`'s resolved locale, for the
 * handful of call sites that build a user-facing message outside a React
 * render (`useFileHandler.ts`'s error strings, `friendlyLibraryError.ts`,
 * `useCodeRun.ts`) and so can't call `useTranslate()`. `LocaleProvider`
 * keeps this in sync via `setActiveLocale` every time its resolved locale
 * changes; until it mounts, this resolves the stored preference against no
 * OS tag, which is exactly right for the (default, and only synchronously
 * knowable) `'en'`/`'fr'` preferences and only briefly wrong for `'system'`
 * on an OS whose locale hasn't been fetched yet (falls back to English,
 * corrected on the next render once `LocaleProvider`'s effect resolves it).
 */
let activeLocale: Locale = resolveLocale(getStoredPreference(), null);

export function setActiveLocale(locale: Locale): void {
  activeLocale = locale;
}

export function getActiveLocale(): Locale {
  return activeLocale;
}
