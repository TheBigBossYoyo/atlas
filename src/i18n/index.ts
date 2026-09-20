export type { Locale, LocalePreference, MessageParams, MessageValue, PluralForms, TranslateFn } from './types';
export type { MessageKey } from './messages.en';
export { LocaleProvider } from './LocaleProvider';
export { useLocale, useTranslate } from './useLocale';
export { LOCALE_KEY } from './locale';

import type { MessageParams } from './types';
import { getActiveLocale } from './activeLocale';
import { translate } from './translate';

/**
 * A free (non-hook) translate function for the handful of call sites that
 * build a user-facing string outside a React render — `useFileHandler.ts`,
 * `friendlyLibraryError.ts`, `useCodeRun.ts`. Reads whatever locale
 * `LocaleProvider` last resolved (see `activeLocale.ts`); everywhere else,
 * prefer `useTranslate()`.
 */
export function t(key: string, params?: MessageParams): string {
  return translate(getActiveLocale(), key, params);
}
