import { useCallback, useEffect, useRef } from 'react';
import { FilePlus, FileText, FileSpreadsheet, FileType, Presentation } from 'lucide-react';
import type { NewDocumentFormat } from '../electron';
import { useShellShortcut } from '../hooks/useShortcutManager';
import { useRestoreFocusOnClose } from '../hooks/useRestoreFocusOnClose';
import { useTranslate } from '../i18n';

interface NewDocumentMenuProps {
  onCreate: (format: NewDocumentFormat) => void;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

interface NewDocumentMenuItem {
  value: NewDocumentFormat;
  labelKey: string;
  Icon: typeof FileText;
}

// NEW-01 — only the formats Atlas can both edit and save end-to-end (see
// `electron/lib/newDocumentTemplates.cjs`'s header for why plain text/code
// and CSV/TSV are deliberately left out).
const NEW_DOCUMENT_ITEMS: readonly NewDocumentMenuItem[] = [
  { value: 'markdown', labelKey: 'newDocument.markdown', Icon: FileText },
  { value: 'docx', labelKey: 'newDocument.docx', Icon: FileType },
  { value: 'xlsx', labelKey: 'newDocument.xlsx', Icon: FileSpreadsheet },
  { value: 'ods', labelKey: 'newDocument.ods', Icon: FileSpreadsheet },
  { value: 'pptx', labelKey: 'newDocument.pptx', Icon: Presentation },
  { value: 'odp', labelKey: 'newDocument.odp', Icon: Presentation },
];

/** NEW-01 — a toolbar dropdown offering every "New document" type, styled and behaving exactly like ExportMenu (outside-click/Escape-to-close, plain labeled buttons rather than the `role="menu"` pattern). */
export function NewDocumentMenu({ onCreate, open, onOpenChange }: NewDocumentMenuProps) {
  const t = useTranslate();
  const ref = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  useRestoreFocusOnClose(open, triggerRef);

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

  const handleCreate = useCallback(
    (format: NewDocumentFormat) => {
      onCreate(format);
      onOpenChange(false);
    },
    [onCreate, onOpenChange],
  );

  return (
    <div className="dropdown" ref={ref}>
      <button
        ref={triggerRef}
        className="toolbar__btn toolbar__nodrag"
        title={t('newDocument.triggerTitle')}
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => onOpenChange(!open)}
      >
        <FilePlus size={16} />
        <span>{t('newDocument.trigger')}</span>
      </button>
      {open && (
        <ul className="dropdown__menu" aria-label={t('newDocument.menuLabel')}>
          {NEW_DOCUMENT_ITEMS.map(({ value, labelKey, Icon }) => (
            <li key={value}>
              <button className="dropdown__item" onClick={() => handleCreate(value)}>
                <Icon size={15} />
                <span className="dropdown__item-label">{t(labelKey)}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
