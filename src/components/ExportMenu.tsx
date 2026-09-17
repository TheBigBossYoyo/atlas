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

    // X1 — a passthrough "Save a copy" of the original bytes, unrelated to
    // the real vector PDF export every other format now gets.
    if (format === 'pdf') {
      return [{ value: 'copy', label: 'Save a copy', Icon: FileDown }] as const;
    }

    // wave-4 legacy-office — LegacyDocViewer/LegacyPptViewer are read-only,
    // text-only previews with no render surface the PDF/HTML export helpers
    // know how to print; no export item is offered rather than one that
    // always fails (see App.tsx's `handleNonMarkdownExport`).
    if (format === 'doc' || format === 'ppt') {
      return [] as const;
    }

    // X1 — a real per-format PDF export now exists for every format (table
    // PDF for xlsx/ods, from the parsed workbook — see `spreadsheetPdf.ts`),
    // so xlsx/ods additionally get a native "Save a copy" of their own bytes
    // alongside the CSV/PDF exports every spreadsheet-shaped format shares.
    if (format === 'xlsx' || format === 'ods') {
      return [
        { value: 'pdf', label: 'Export to PDF', Icon: FileText },
        { value: 'csv', label: 'Export to CSV', Icon: FileSpreadsheet },
        { value: 'copy', label: 'Save a copy', Icon: FileDown },
      ] as const;
    }

    // UX-12 — csv/tsv already carry real delimited-text content, so a
    // faithful "Export to CSV" is always safe to offer alongside the real
    // table PDF export. A spreadsheet format only gets it once its viewer
    // registers real parsed-row content (`canExportCsv`).
    if (format === 'csv' || format === 'tsv' || canExportCsv) {
      return [
        { value: 'pdf', label: 'Export to PDF', Icon: FileText },
        { value: 'csv', label: 'Export to CSV', Icon: FileSpreadsheet },
      ] as const;
    }

    // X1 — text/code additionally get a standalone HTML export (both are
    // plain content with no format-specific PDF-vs-HTML distinction the way
    // markdown has).
    if (format === 'text' || format === 'code') {
      return [
        { value: 'pdf', label: 'Export to PDF', Icon: FileText },
        { value: 'html', label: 'Export to HTML', Icon: FileCode },
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
