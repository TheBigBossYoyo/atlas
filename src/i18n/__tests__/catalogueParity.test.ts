import { describe, expect, it } from 'vitest';
import { messages as en } from '../messages.en';
import { messages as fr } from '../messages.fr';

/**
 * Runtime counterpart to the compile-time check `messages.fr.ts` already gets
 * from `satisfies Record<MessageKey, MessageValue>` — that catches a missing
 * French key at build time, but not silently (e.g. a stray key typo'd
 * differently in each file would still satisfy both types independently).
 * This test compares the two actual key sets directly, so a mismatch fails
 * loudly with the exact offending key(s) rather than a generic type error.
 */
describe('i18n catalogue parity', () => {
  const enKeys = Object.keys(en);
  const frKeys = Object.keys(fr);

  it('has no keys in the English catalogue missing from French', () => {
    const missing = enKeys.filter((key) => !(key in fr));
    expect(missing).toEqual([]);
  });

  it('has no orphan keys in the French catalogue absent from English', () => {
    const orphans = frKeys.filter((key) => !(key in en));
    expect(orphans).toEqual([]);
  });

  it('has the exact same key count in both catalogues', () => {
    expect(frKeys.length).toBe(enKeys.length);
  });

  it('matches plural-vs-plain shape for every key (a plural entry in one locale must be plural in the other)', () => {
    const mismatches = enKeys.filter((key) => {
      const enIsPlural = typeof en[key as keyof typeof en] !== 'string';
      const frIsPlural = typeof fr[key as keyof typeof fr] !== 'string';
      return enIsPlural !== frIsPlural;
    });
    expect(mismatches).toEqual([]);
  });

  it('has no empty string values in either catalogue', () => {
    const isEmpty = (value: unknown): boolean =>
      typeof value === 'string' ? value.length === 0 : false;
    const emptyInEn = enKeys.filter((key) => isEmpty(en[key as keyof typeof en]));
    const emptyInFr = frKeys.filter((key) => isEmpty(fr[key as keyof typeof fr]));
    expect(emptyInEn).toEqual([]);
    expect(emptyInFr).toEqual([]);
  });

  // Regression: both the toolbar's view-mode button and the shortcuts modal's
  // matching row translated "Split" (the Markdown preview/split/editor view
  // toggle) as "Partagé" ("shared"), not "Fractionné" — the term Word/
  // LibreOffice actually use for a split view. Caught by the same
  // documentation-truth pass that found the two false claims in
  // README/docs/KNOWN_LIMITATIONS.md.
  it('translates the "Split" view mode as "Fractionné", not "Partagé"', () => {
    expect(fr['toolbar.viewMode.split']).toBe('Fractionné');
    expect(fr['shortcuts.split']).toBe('Fractionné');
  });
});
