import type { Locale, LocalePreference } from './types';

/**
 * Persistence key, mirroring `useTheme.ts`'s `atlas-theme`. Unlike the
 * theme (which defaults to the OS preference), the language default is a
 * literal `'en'` — never derived from `app.getLocale()` — so a fresh
 * profile and the Playwright e2e suite (which never sets a preference)
 * both always start in English, regardless of the host OS's locale. Only
 * an explicit `'system'` choice ever consults the OS locale.
 */
export const LOCALE_KEY = 'atlas-locale';

function isLocalePreference(value: string | null): value is LocalePreference {
  return value === 'en' || value === 'fr' || value === 'system';
}

/** Reads the persisted preference, defaulting to `'en'` — see module doc comment. */
export function getStoredPreference(): LocalePreference {
  try {
    const stored = window.localStorage.getItem(LOCALE_KEY);
    return isLocalePreference(stored) ? stored : 'en';
  } catch {
    return 'en';
  }
}

export function storePreference(preference: LocalePreference): void {
  try {
    window.localStorage.setItem(LOCALE_KEY, preference);
  } catch {
    // localStorage unavailable (private mode, etc.) — the in-memory state
    // set by the caller still governs this session.
  }
}

/** Maps an OS locale tag (e.g. `"fr-FR"`, `"fr"`) to one of Atlas's supported locales. */
export function localeFromOsTag(osTag: string | null | undefined): Locale {
  return typeof osTag === 'string' && osTag.toLowerCase().startsWith('fr') ? 'fr' : 'en';
}

/** Resolves a stored preference to a renderable `Locale`, consulting the OS tag only for `'system'`. */
export function resolveLocale(preference: LocalePreference, osTag: string | null | undefined): Locale {
  if (preference === 'system') return localeFromOsTag(osTag);
  return preference;
}
