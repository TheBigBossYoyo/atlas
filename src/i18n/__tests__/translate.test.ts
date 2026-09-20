import { describe, expect, it } from 'vitest';
import { formatMessage, interpolate, translate } from '../translate';
import type { PluralForms } from '../types';

describe('interpolate', () => {
  it('replaces a single placeholder', () => {
    expect(interpolate('Hello {name}', 'en', { name: 'World' })).toBe('Hello World');
  });

  it('replaces multiple distinct placeholders', () => {
    expect(interpolate('{a} and {b}', 'en', { a: 'X', b: 'Y' })).toBe('X and Y');
  });

  it('formats a numeric param with locale-aware thousands separators', () => {
    expect(interpolate('{count} items', 'en', { count: 1234 })).toBe('1,234 items');
    // French uses a narrow no-break space (U+202F) as the thousands separator.
    expect(interpolate('{count} items', 'fr', { count: 1234 })).toBe('1 234 items');
  });

  it('leaves an unmatched placeholder untouched instead of throwing', () => {
    expect(interpolate('Hello {name}', 'en')).toBe('Hello {name}');
    expect(interpolate('Hello {name}', 'en', {})).toBe('Hello {name}');
  });

  it('leaves a template with no placeholders unchanged', () => {
    expect(interpolate('Plain text', 'en', { unused: 'x' })).toBe('Plain text');
  });
});

describe('formatMessage — plural forms', () => {
  const forms: PluralForms = { one: '{count} item', other: '{count} items' };

  it('selects the singular form for count=1 in English', () => {
    expect(formatMessage(forms, 'en', { count: 1 })).toBe('1 item');
  });

  it('selects the plural form for count=0 and count>1 in English', () => {
    expect(formatMessage(forms, 'en', { count: 0 })).toBe('0 items');
    expect(formatMessage(forms, 'en', { count: 5 })).toBe('5 items');
  });

  it('selects the singular form for count=1 in French too (same cardinal rule)', () => {
    expect(formatMessage(forms, 'fr', { count: 1 })).toBe('1 item');
  });

  it('defaults the count to 0 when no params are given', () => {
    expect(formatMessage(forms, 'en')).toBe('0 items');
  });

  it('prefers an explicit `zero` form over the plural-rule result when count is 0', () => {
    const withZero: PluralForms = { zero: 'no items', one: '{count} item', other: '{count} items' };
    expect(formatMessage(withZero, 'en', { count: 0 })).toBe('no items');
  });
});

describe('formatMessage — plain strings', () => {
  it('interpolates a plain (non-plural) message', () => {
    expect(formatMessage('Hi {name}', 'en', { name: 'Ada' })).toBe('Hi Ada');
  });
});

describe('translate', () => {
  it('resolves a real key to its French text', () => {
    expect(translate('fr', 'common.dismiss')).toBe('Ignorer');
  });

  it('resolves the same key to its English text', () => {
    expect(translate('en', 'common.dismiss')).toBe('Dismiss');
  });

  it('returns the bare key for a key that exists in neither catalogue', () => {
    expect(translate('en', 'this.key.does.not.exist')).toBe('this.key.does.not.exist');
    expect(translate('fr', 'this.key.does.not.exist')).toBe('this.key.does.not.exist');
  });

  it('interpolates params through a full translate() call', () => {
    expect(translate('en', 'tabBar.closeAria', { name: 'report.md' })).toBe('Close report.md');
    expect(translate('fr', 'tabBar.closeAria', { name: 'report.md' })).toBe('Fermer report.md');
  });

  it('applies plural selection through a full translate() call', () => {
    expect(translate('en', 'statusBar.words', { count: 1 })).toBe('1 word');
    expect(translate('en', 'statusBar.words', { count: 2 })).toBe('2 words');
    expect(translate('fr', 'statusBar.words', { count: 1 })).toBe('1 mot');
    expect(translate('fr', 'statusBar.words', { count: 2 })).toBe('2 mots');
  });
});
