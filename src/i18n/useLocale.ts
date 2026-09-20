import { useContext } from 'react';
import type { TranslateFn } from './types';
import { LocaleContext, type LocaleContextValue } from './context';

function useLocaleContext(): LocaleContextValue {
  return useContext(LocaleContext);
}

/** The resolved locale, the raw preference, and a setter — for the `LanguageMenu`. */
export function useLocale(): Pick<LocaleContextValue, 'locale' | 'preference' | 'setPreference'> {
  return useLocaleContext();
}

/** The translate function every other component needs. */
export function useTranslate(): TranslateFn {
  return useLocaleContext().t;
}
