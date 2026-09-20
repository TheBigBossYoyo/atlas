import type { Locale, MessageParams, MessageValue, PluralForms } from './types';
import { messages as enMessages } from './messages.en';
import { messages as frMessages } from './messages.fr';

const CATALOGUES: Readonly<Record<Locale, Readonly<Record<string, MessageValue>>>> = {
  en: enMessages,
  fr: frMessages,
};

/** BCP-47 tag `Intl.PluralRules`/`Intl.NumberFormat` expect for each locale. */
export function localeTag(locale: Locale): string {
  return locale === 'fr' ? 'fr-FR' : 'en-US';
}

const pluralRulesCache = new Map<Locale, Intl.PluralRules>();

function pluralRulesFor(locale: Locale): Intl.PluralRules {
  let rules = pluralRulesCache.get(locale);
  if (!rules) {
    rules = new Intl.PluralRules(localeTag(locale));
    pluralRulesCache.set(locale, rules);
  }
  return rules;
}

function isPluralForms(value: MessageValue): value is PluralForms {
  return typeof value !== 'string';
}

const numberFormatterCache = new Map<Locale, Intl.NumberFormat>();

function numberFormatterFor(locale: Locale): Intl.NumberFormat {
  let formatter = numberFormatterCache.get(locale);
  if (!formatter) {
    formatter = new Intl.NumberFormat(localeTag(locale));
    numberFormatterCache.set(locale, formatter);
  }
  return formatter;
}

/**
 * Replaces every `{name}` in `template` with `params.name`, left untouched
 * when no param matches. A numeric value is formatted with `Intl.NumberFormat`
 * for the given locale (thousands separators, etc.) rather than stringified
 * raw — matching what every count in the shell showed before i18n (e.g.
 * `StatusBar`'s word/row counts).
 */
export function interpolate(template: string, locale: Locale, params?: MessageParams): string {
  if (!params) return template;
  const formatter = numberFormatterFor(locale);
  return template.replace(/\{(\w+)\}/g, (match, key: string) => {
    const value = params[key];
    if (value === undefined) return match;
    return typeof value === 'number' ? formatter.format(value) : String(value);
  });
}

function selectPluralTemplate(value: PluralForms, locale: Locale, count: number): string {
  if (count === 0 && value.zero !== undefined) return value.zero;
  const category = pluralRulesFor(locale).select(count);
  if (category === 'one') return value.one;
  return value.other;
}

/**
 * Resolves one catalogue entry (plain string or plural forms) to its final
 * display string for `locale`, with `{placeholder}` interpolation applied.
 * A plural entry needs `params.count`; missing it is treated as `0`.
 */
export function formatMessage(value: MessageValue, locale: Locale, params?: MessageParams): string {
  if (!isPluralForms(value)) {
    return interpolate(value, locale, params);
  }
  const count = Number(params?.count ?? 0);
  const template = selectPluralTemplate(value, locale, count);
  // A plural template's own `{count}` must interpolate even when the caller
  // passed no `params` at all (count then defaults to 0 above) — merge it in
  // rather than passing `params` through unchanged.
  return interpolate(template, locale, { ...params, count });
}

/**
 * Looks a key up in `locale`'s catalogue (falling back to English for a key
 * missing from a non-English catalogue — shouldn't happen once
 * `catalogueParity.test.ts` passes, but keeps the UI showing *something*
 * translatable rather than the raw key if it ever does) and formats it.
 * Returns the bare key when it exists in neither catalogue, so a typo'd key
 * is obvious in the rendered UI instead of throwing.
 */
export function translate(locale: Locale, key: string, params?: MessageParams): string {
  const value = CATALOGUES[locale][key] ?? enMessages[key as keyof typeof enMessages];
  if (value === undefined) return key;
  return formatMessage(value, locale, params);
}
