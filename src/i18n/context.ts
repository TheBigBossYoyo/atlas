import { createContext } from 'react';
import type { Locale, LocalePreference, TranslateFn } from './types';
import { translate } from './translate';

/**
 * Context value + the context object itself — kept separate from
 * `LocaleProvider.tsx` (component only) and `useLocale.ts` (hooks only) so
 * each file exports only one kind of thing, satisfying
 * `react-refresh/only-export-components`. Same split already used for
 * `ToastContext.ts`/`ToastProvider.tsx`/`useToast.ts`.
 */
export interface LocaleContextValue {
  readonly locale: Locale;
  readonly preference: LocalePreference;
  readonly setPreference: (preference: LocalePreference) => void;
  readonly t: TranslateFn;
}

/**
 * A real (non-null) English default, not `null` — matches
 * `useShortcutManager.ts`'s "tolerate rendering without the provider" rule:
 * many existing component tests (`Toolbar.test.tsx`, `ThemeMenu.test.tsx`,
 * `NewDocumentMenu.test.tsx`, ...) render a single component standalone,
 * without the app's full provider stack, and `useTranslate()` there must
 * still return readable English text instead of throwing. It also means a
 * fresh render before `<LocaleProvider>` mounts (there isn't one in this
 * app, but a future caller might) reads as English, consistent with the
 * default-English requirement everywhere else in `src/i18n`.
 */
const DEFAULT_LOCALE_CONTEXT: LocaleContextValue = {
  locale: 'en',
  preference: 'en',
  setPreference: () => {
    // No-op outside a real <LocaleProvider> — nothing persists a preference
    // choice made without one.
  },
  t: (key, params) => translate('en', key, params),
};

export const LocaleContext = createContext<LocaleContextValue>(DEFAULT_LOCALE_CONTEXT);
