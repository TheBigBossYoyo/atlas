import { useCallback, useEffect, useMemo, useRef } from 'react';
import { Download, FileCode, FileText, FileType, FileDown } from 'lucide-react';
import type { FormatId } from '../formats/types';
import type { ExportFormat } from '../types';
import { useShellShortcut } from '../hooks/useShortcutManager';

interface ExportMenuProps {
  onExport: (format: ExportFormat) => void;
  format: FormatId;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  disabled?: boolean;
}

interface ExportMenuItem {
  value: ExportFormat;
  label: string;
  Icon: typeof FileCode;
}

export function ExportMenu({ onExport, format, open, onOpenChange, disabled }: ExportMenuProps) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        onOpenChange(false);
      }
    };
    if (open) {
      window.addEventListener('mousedown', handleOutsideClick);
    }
    return () => {
      window.removeEventListener('mousedown', handleOutsideClick);
    };
  }, [onOpenChange, open]);

  // P2.1 — Escape-close registered with the shared dispatcher instead of its
  // own ad hoc `window` listener; only listens while the menu is open.
  useShellShortcut(
    useCallback(
      (e) => {
        if (e.key !== 'Escape') return false;
        onOpenChange(false);
        return true;
      },
      [onOpenChange],
    ),
    open,
  );

  const handleExport = useCallback((targetFormat: ExportFormat) => {
    onExport(targetFormat);
    onOpenChange(false);
  }, [onExport, onOpenChange]);

  const items = useMemo<readonly ExportMenuItem[]>(() => {
    if (format === 'markdown') {
      return [
        { value: 'html', label: 'HTML', Icon: FileCode },
        { value: 'pdf', label: 'PDF', Icon: FileText },
        { value: 'docx', label: 'DOCX', Icon: FileType },
        { value: 'md', label: 'Markdown', Icon: FileDown },
      ] as const;
    }

    if (format === 'pdf') {
      return [{ value: 'pdf', label: 'Save a copy', Icon: FileDown }] as const;
    }

    return [{ value: 'pdf', label: 'Export to PDF', Icon: FileText }] as const;
  }, [format]);

  return (
    <div className="dropdown" ref={ref}>
      <button
        className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
        title="Export"
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => onOpenChange(!open)}
        disabled={disabled}
      >
        <Download size={18} />
      </button>
      {open && (
        <ul className="dropdown__menu" role="menu">
          {items.map(({ value, label, Icon }) => (
            <li key={`${format}-${value}-${label}`}>
              <button className="dropdown__item" role="menuitem" onClick={() => handleExport(value)}>
                <Icon size={15} />
                <span className="dropdown__item-label">{label}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
