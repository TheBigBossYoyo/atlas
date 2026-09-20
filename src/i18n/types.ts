/**
 * i18n — shared types for Atlas's small, dependency-free translation layer.
 *
 * `Locale` is a resolved, renderable language. `LocalePreference` is what the
 * user (or a fresh profile) actually has stored — `'system'` isn't itself a
 * renderable language, it means "resolve from `app.getLocale()` at run
 * time" (see `resolveLocale` in `locale.ts`).
 */
export type Locale = 'en' | 'fr';

export type LocalePreference = Locale | 'system';

/**
 * A pluralized message. `zero` is optional (falls back to `other` when
 * absent, which is correct for both English and French outside a handful of
 * explicitly-authored zero-count messages). Selection uses `Intl.PluralRules`
 * for the active locale, so this also works for locales with richer plural
 * systems if Atlas ever adds one.
 */
export interface PluralForms {
  readonly zero?: string;
  readonly one: string;
  readonly other: string;
}

export type MessageValue = string | PluralForms;

export type MessageCatalogue = Readonly<Record<string, MessageValue>>;

/** Values interpolation accepts for a `{placeholder}` in a message template. */
export type MessageParams = Readonly<Record<string, string | number>>;

/** Signature of the `t()` function handed out by `useTranslate()`. */
export type TranslateFn = (key: string, params?: MessageParams) => string;
