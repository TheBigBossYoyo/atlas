import { useCallback, useEffect, useRef } from 'react';
import { X } from 'lucide-react';
import { useShellShortcut } from '../hooks/useShortcutManager';
import { useTranslate } from '../i18n';

interface ShortcutsModalProps {
  isOpen: boolean;
  onClose: () => void;
}

const FOCUSABLE_SELECTOR = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';

export function ShortcutsModal({ isOpen, onClose }: ShortcutsModalProps) {
  const t = useTranslate();
  const modalRef = useRef<HTMLDivElement>(null);
  const previouslyFocusedRef = useRef<HTMLElement | null>(null);

  // P2.1 — Escape-close registered with the shared dispatcher instead of its
  // own ad hoc `window` listener; only listens while the modal is open.
  useShellShortcut(
    useCallback(
      (e) => {
        if (e.key !== 'Escape') return false;
        onClose();
        return true;
      },
      [onClose],
    ),
    isOpen,
  );

  // UX-13 — focus trap: move focus into the modal on open, keep Tab/Shift+Tab
  // cycling within it while it's open, and restore focus to whatever
  // triggered it once it closes (a separate effect/listener from the Escape
  // handler above — this is Tab containment, not the close shortcut).
  useEffect(() => {
    if (!isOpen) return;

    previouslyFocusedRef.current = document.activeElement as HTMLElement | null;
    const modal = modalRef.current;
    modal?.querySelector<HTMLElement>(FOCUSABLE_SELECTOR)?.focus();

    const handleTab = (e: KeyboardEvent) => {
      if (e.key !== 'Tab' || !modal) return;
      const focusables = Array.from(modal.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR));
      if (focusables.length === 0) return;
      const first = focusables[0]!;
      const last = focusables[focusables.length - 1]!;
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    };

    window.addEventListener('keydown', handleTab);
    return () => {
      window.removeEventListener('keydown', handleTab);
      previouslyFocusedRef.current?.focus();
    };
  }, [isOpen]);

  if (!isOpen) return null;

  const renderKeys = (combo: string) => {
    const keys = combo.split('+');
    return keys.map((key, i) => (
      <span key={i}>
        <kbd>{key}</kbd>
        {i < keys.length - 1 && ' + '}
      </span>
    ));
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div ref={modalRef} className="modal" role="dialog" aria-modal="true" aria-labelledby="shortcuts-title" onClick={e => e.stopPropagation()}>
        <div className="modal__header">
          <h2 id="shortcuts-title" className="modal__title">{t('shortcuts.title')}</h2>
          <button className="modal__close" aria-label={t('shortcuts.closeAria')} onClick={onClose}>
            <X size={18} />
          </button>
        </div>

        <section className="modal__section">
          <h3 className="modal__section-title">{t('shortcuts.sectionFile')}</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+N')}</dt><dd>{t('shortcuts.newDocumentMenu')}</dd>
            <dt>{renderKeys('Ctrl+O')}</dt><dd>{t('shortcuts.openFile')}</dd>
            <dt>{renderKeys('Ctrl+S')}</dt><dd>{t('shortcuts.save')}</dd>
            <dt>{renderKeys('Ctrl+Shift+S')}</dt><dd>{t('shortcuts.saveAs')}</dd>
            <dt>{renderKeys('Ctrl+W')}</dt><dd>{t('shortcuts.closeFile')}</dd>
            <dt>{renderKeys('Ctrl+Tab')} / {renderKeys('Ctrl+Shift+Tab')}</dt><dd>{t('shortcuts.nextPrevTab')}</dd>
            <dt>{renderKeys('Ctrl+Shift+T')}</dt><dd>{t('shortcuts.reopenClosed')}</dd>
            <dt>{renderKeys('Ctrl+P')}</dt><dd>{t('shortcuts.printExport')}</dd>
            <dt>{renderKeys('Ctrl+E')}</dt><dd>{t('shortcuts.exportMenu')}</dd>
          </dl>
        </section>

        <section className="modal__section">
          <h3 className="modal__section-title">{t('shortcuts.sectionView')}</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+1')}</dt><dd>{t('shortcuts.preview')}</dd>
            <dt>{renderKeys('Ctrl+2')}</dt><dd>{t('shortcuts.split')}</dd>
            <dt>{renderKeys('Ctrl+3')}</dt><dd>{t('shortcuts.editor')}</dd>
            <dt>{renderKeys('Ctrl+B')}</dt><dd>{t('shortcuts.toggleSidebar')}</dd>
            <dt>{renderKeys('Ctrl+T')}</dt><dd>{t('shortcuts.cycleTheme')}</dd>
          </dl>
        </section>

        <section className="modal__section">
          <h3 className="modal__section-title">{t('shortcuts.sectionEdit')}</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+=')}</dt><dd>{t('shortcuts.increaseFont')}</dd>
            <dt>{renderKeys('Ctrl+-')}</dt><dd>{t('shortcuts.decreaseFont')}</dd>
            <dt>{renderKeys('Ctrl+0')}</dt><dd>{t('shortcuts.resetFont')}</dd>
          </dl>
        </section>

        <section className="modal__section">
          <h3 className="modal__section-title">{t('shortcuts.sectionSearchHelp')}</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+F')}</dt><dd>{t('shortcuts.find')}</dd>
            <dt>{renderKeys('Enter')} / {renderKeys('Shift+Enter')}</dt><dd>{t('shortcuts.nextPrev')}</dd>
            <dt>{renderKeys('Esc')}</dt><dd>{t('shortcuts.closeDialog')}</dd>
            <dt>{renderKeys('Ctrl+/')}</dt><dd>{t('shortcuts.toggleDialog')}</dd>
          </dl>
        </section>

        <p className="modal__note">
          {t('shortcuts.note')}
        </p>
      </div>
    </div>
  );
}
