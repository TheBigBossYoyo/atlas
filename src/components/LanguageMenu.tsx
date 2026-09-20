import { useState, useRef, useEffect, useCallback } from 'react';
import { Languages, Check } from 'lucide-react';
import type { LocalePreference } from '../i18n';
import { useLocale, useTranslate } from '../i18n';
import { useShellShortcut } from '../hooks/useShortcutManager';
import { useRestoreFocusOnClose } from '../hooks/useRestoreFocusOnClose';

const OPTIONS: readonly LocalePreference[] = ['en', 'fr', 'system'];

const OPTION_LABEL_KEY: Readonly<Record<LocalePreference, string>> = {
  en: 'language.english',
  fr: 'language.french',
  system: 'language.system',
};

/**
 * The language picker, styled and behaving exactly like `ThemeMenu` (same
 * dropdown, outside-click/Escape-to-close, plain labeled buttons rather than
 * the `role="menu"` ARIA pattern — see that component's UX-14 note). Placed
 * next to `ThemeMenu` in the toolbar per the i18n plan.
 */
export function LanguageMenu() {
  const { preference, setPreference } = useLocale();
  const t = useTranslate();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  useRestoreFocusOnClose(open, triggerRef);

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    if (open) {
      window.addEventListener('mousedown', handleOutsideClick);
    }
    return () => {
      window.removeEventListener('mousedown', handleOutsideClick);
    };
  }, [open]);

  useShellShortcut(
    useCallback(
      (e) => {
        if (e.key !== 'Escape') return false;
        setOpen(false);
        return true;
      },
      [],
    ),
    open,
  );

  return (
    <div className="dropdown" ref={ref}>
      <button
        ref={triggerRef}
        className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
        title={t('language.triggerTitle')}
        aria-label={t('language.triggerAria')}
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        <Languages size={18} />
      </button>
      {open && (
        <ul className="dropdown__menu" aria-label={t('language.menuLabel')}>
          {OPTIONS.map((option) => (
            <li key={option}>
              <button
                className={`dropdown__item ${preference === option ? 'dropdown__item--active' : ''}`}
                aria-current={preference === option ? 'true' : undefined}
                onClick={() => {
                  setPreference(option);
                  setOpen(false);
                }}
              >
                <span className="dropdown__item-label">{t(OPTION_LABEL_KEY[option])}</span>
                {preference === option && <Check size={14} />}
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
