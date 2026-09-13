import { useEffect } from 'react';
import { X } from 'lucide-react';

interface ShortcutsModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export function ShortcutsModal({ isOpen, onClose }: ShortcutsModalProps) {
  useEffect(() => {
    const handleEsc = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    if (isOpen) {
      window.addEventListener('keydown', handleEsc);
    }
    return () => {
      window.removeEventListener('keydown', handleEsc);
    };
  }, [isOpen, onClose]);

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
      <div className="modal" role="dialog" aria-modal="true" aria-labelledby="shortcuts-title" onClick={e => e.stopPropagation()}>
        <div className="modal__header">
          <h2 id="shortcuts-title" className="modal__title">Keyboard Shortcuts</h2>
          <button className="modal__close" aria-label="Close" onClick={onClose}>
            <X size={18} />
          </button>
        </div>
        
        <section className="modal__section">
          <h3 className="modal__section-title">File</h3>
          <dl className="modal__shortcuts">
            <dt>{renderKeys('Ctrl+O')}</dt><dd>Open file</dd>
            <dt>{renderKeys('Ctrl+S')}</dt><dd>Save</dd>
            <dt>{renderKeys('Ctrl+Shift+S')}</dt><dd>Save As</dd>
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
            <dt>{renderKeys('Ctrl+F')}</dt><dd>Find</dd>
            <dt>{renderKeys('Enter')} / {renderKeys('Shift+Enter')}</dt><dd>Next/prev</dd>
            <dt>{renderKeys('Esc')}</dt><dd>Close</dd>
            <dt>{renderKeys('Ctrl+/')}</dt><dd>Toggle this dialog</dd>
          </dl>
        </section>
      </div>
    </div>
  );
}
