import { useCallback, useEffect, useRef } from 'react';
import { X } from 'lucide-react';
import { useShellShortcut } from '../hooks/useShortcutManager';

interface ShortcutsModalProps {
  isOpen: boolean;
  onClose: () => void;
}

const FOCUSABLE_SELECTOR = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';

export function ShortcutsModal({ isOpen, onClose }: ShortcutsModalProps) {
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
          <h2 id="shortcuts-title" className="modal__title">Keyboard Shortcuts</h2>
          <button className="modal__close" aria-label="Close" onClick={onClose}>
            <X size={18} />
          </button>
        </div>
        
        <section className="modal__section">
          <h3 className="modal__section-title">File</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+N')}</dt><dd>New document menu</dd>
            <dt>{renderKeys('Ctrl+O')}</dt><dd>Open file</dd>
            <dt>{renderKeys('Ctrl+S')}</dt><dd>Save</dd>
            <dt>{renderKeys('Ctrl+Shift+S')}</dt><dd>Save As</dd>
            <dt>{renderKeys('Ctrl+W')}</dt><dd>Close file</dd>
            <dt>{renderKeys('Ctrl+P')}</dt><dd>Print / export</dd>
            <dt>{renderKeys('Ctrl+E')}</dt><dd>Export menu</dd>
          </dl>
        </section>

        <section className="modal__section">
          <h3 className="modal__section-title">View</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+1')}</dt><dd>Preview</dd>
            <dt>{renderKeys('Ctrl+2')}</dt><dd>Split</dd>
            <dt>{renderKeys('Ctrl+3')}</dt><dd>Editor</dd>
            <dt>{renderKeys('Ctrl+B')}</dt><dd>Toggle sidebar</dd>
            <dt>{renderKeys('Ctrl+T')}</dt><dd>Cycle theme</dd>
          </dl>
        </section>

        <section className="modal__section">
          <h3 className="modal__section-title">Edit</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+=')}</dt><dd>Increase font</dd>
            <dt>{renderKeys('Ctrl+-')}</dt><dd>Decrease font</dd>
            <dt>{renderKeys('Ctrl+0')}</dt><dd>Reset font</dd>
          </dl>
        </section>

        <section className="modal__section">
          <h3 className="modal__section-title">Search & Help</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+F')}</dt><dd>Find (Markdown, Text, Code, RTF, ODT)</dd>
            <dt>{renderKeys('Enter')} / {renderKeys('Shift+Enter')}</dt><dd>Next/prev</dd>
            <dt>{renderKeys('Esc')}</dt><dd>Close</dd>
            <dt>{renderKeys('Ctrl+/')}</dt><dd>Toggle this dialog</dd>
          </dl>
        </section>

        <p className="modal__note">
          While editing a DOCX, its own editor shortcuts (bold/italic/underline, alignment, find/replace, line spacing, undo/redo, save, print) take priority over the shortcuts above.
        </p>
      </div>
    </div>
  );
}
