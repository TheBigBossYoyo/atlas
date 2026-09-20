import { fireEvent, render, screen } from '@testing-library/react';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { LocaleProvider } from '../LocaleProvider';
import { useLocale, useTranslate } from '../useLocale';
import { LOCALE_KEY } from '../locale';

function Probe() {
  const t = useTranslate();
  const { preference, setPreference } = useLocale();
  return (
    <div>
      <span data-testid="save-label">{t('toolbar.save')}</span>
      <span data-testid="preference">{preference}</span>
      <button onClick={() => setPreference('fr')}>go-french</button>
      <button onClick={() => setPreference('en')}>go-english</button>
    </div>
  );
}

describe('LocaleProvider', () => {
  beforeEach(() => {
    window.localStorage.clear();
  });

  afterEach(() => {
    window.localStorage.clear();
  });

  it('defaults to English on a fresh profile (no stored preference)', () => {
    render(
      <LocaleProvider>
        <Probe />
      </LocaleProvider>,
    );
    expect(screen.getByTestId('save-label')).toHaveTextContent('Save');
    expect(screen.getByTestId('preference')).toHaveTextContent('en');
  });

  it('re-renders every consumer with French text after switching the preference', () => {
    render(
      <LocaleProvider>
        <Probe />
      </LocaleProvider>,
    );
    expect(screen.getByTestId('save-label')).toHaveTextContent('Save');

    fireEvent.click(screen.getByText('go-french'));

    expect(screen.getByTestId('save-label')).toHaveTextContent('Enregistrer');
    expect(screen.getByTestId('preference')).toHaveTextContent('fr');
  });

  it('switching back to English re-renders back to English text', () => {
    render(
      <LocaleProvider>
        <Probe />
      </LocaleProvider>,
    );
    fireEvent.click(screen.getByText('go-french'));
    expect(screen.getByTestId('save-label')).toHaveTextContent('Enregistrer');

    fireEvent.click(screen.getByText('go-english'));
    expect(screen.getByTestId('save-label')).toHaveTextContent('Save');
  });

  it('persists the chosen preference to localStorage, like the theme picker', () => {
    render(
      <LocaleProvider>
        <Probe />
      </LocaleProvider>,
    );
    fireEvent.click(screen.getByText('go-french'));
    expect(window.localStorage.getItem(LOCALE_KEY)).toBe('fr');
  });

  it('a fresh mount picks up a previously-persisted French preference', () => {
    window.localStorage.setItem(LOCALE_KEY, 'fr');
    render(
      <LocaleProvider>
        <Probe />
      </LocaleProvider>,
    );
    expect(screen.getByTestId('save-label')).toHaveTextContent('Enregistrer');
  });
});
