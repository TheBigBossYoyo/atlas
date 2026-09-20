import { FileX2 } from 'lucide-react';
import { useTranslate } from '../i18n';

interface EmptyFileNoticeProps {
  readonly fileName: string;
}

/**
 * NEW-01 — a 0-byte file of a format Atlas can't treat as a blank editable
 * document (pdf/rtf/odt/doc/ppt/unknown — every BINARY_CLASS format
 * `electron/lib/newDocumentTemplates.cjs` has no blank template for) used to
 * reach that format's viewer/parser as a genuinely empty buffer and fail
 * with a low-level parse error ("Failed to unzip DOCX: End of data
 * reached…"-style message, just for a different format). Shown instead of
 * `ViewerRouter` for exactly that case — see App.tsx's `viewerFile` render
 * branch. Editable formats never reach here: main.cjs's read handlers
 * transparently substitute that format's blank template for a 0-byte read
 * before the renderer ever sees it (see that file's
 * `substituteBlankTemplateIfEmpty`).
 */
export function EmptyFileNotice({ fileName }: EmptyFileNoticeProps) {
  const t = useTranslate();
  return (
    <div className="viewer-fallback empty-file-notice">
      <FileX2 size={48} strokeWidth={1.5} className="viewer-fallback__icon" aria-hidden="true" />
      <h2 className="viewer-fallback__title">{t('emptyFile.title')}</h2>
      <p className="viewer-fallback__detail">
        {t('emptyFile.detail', { fileName })}
      </p>
    </div>
  );
}
