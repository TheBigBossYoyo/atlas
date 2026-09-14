import { useCallback, useEffect, useMemo, useRef } from 'react';
import { Download, FileCode, FileText, FileType, FileDown, FileSpreadsheet } from 'lucide-react';
import type { FormatId } from '../formats/types';
import type { ExportFormat } from '../types';
import { useShellShortcut } from '../hooks/useShortcutManager';

interface ExportMenuProps {
  onExport: (format: ExportFormat) => void;
  format: FormatId;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  disabled?: boolean;
  /** UX-12 — whether a real CSV export is available for the current format
   * (always true for csv/tsv; true for a spreadsheet format only once its
   * viewer has registered `getExportableContent` with `format: 'csv'`). */
  canExportCsv?: boolean;
}

interface ExportMenuItem {
  value: ExportFormat;
  label: string;
  Icon: typeof FileCode;
}

export function ExportMenu({ onExport, format, open, onOpenChange, disabled, canExportCsv }: ExportMenuProps) {
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

    // UX-12 — csv/tsv already carry real delimited-text content, so a
    // faithful "Export to CSV" is always safe to offer alongside the
    // generic full-panel PDF screenshot. A spreadsheet format only gets it
    // once its viewer registers real parsed-row content (`canExportCsv`).
    if (format === 'csv' || format === 'tsv' || canExportCsv) {
      return [
        { value: 'pdf', label: 'Export to PDF', Icon: FileText },
        { value: 'csv', label: 'Export to CSV', Icon: FileSpreadsheet },
      ] as const;
    }

    return [{ value: 'pdf', label: 'Export to PDF', Icon: FileText }] as const;
  }, [canExportCsv, format]);

  return (
    <div className="dropdown" ref={ref}>
      <button
        className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
        title="Export"
        aria-label="Export"
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => onOpenChange(!open)}
        disabled={disabled}
      >
        <Download size={18} />
      </button>
      {open && (
        // UX-14 — a plain labeled list of buttons, not the `role="menu"`
        // ARIA pattern (which requires full arrow-key navigation this
        // dropdown never implemented, so it was announcing menu semantics AT
        // users couldn't actually use — Tab/Shift+Tab works with plain
        // buttons out of the box).
        <ul className="dropdown__menu" aria-label="Export options">
          {items.map(({ value, label, Icon }) => (
            <li key={`${format}-${value}-${label}`}>
              <button className="dropdown__item" onClick={() => handleExport(value)}>
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
