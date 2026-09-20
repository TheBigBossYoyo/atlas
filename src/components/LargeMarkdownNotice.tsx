import { FileWarning } from 'lucide-react';

import { useTranslate } from '../i18n';

/**
 * Markdown parsing (micromark, under react-markdown) degrades superlinearly:
 * measured on this project's own corpus, 0.5 MB takes ~2 s, 1 MB ~9 s and
 * 2 MB ~44 s, all of it on the main thread — a 5 MB document froze the app for
 * over two minutes and peaked at 2.6 GB. The editor pane stays fast whatever
 * the size, so past this threshold the document opens there with this notice
 * in place of the preview, and rendering stays one click away.
 */
export const LARGE_MARKDOWN_PREVIEW_BYTES = 750_000;

interface LargeMarkdownNoticeProps {
  readonly characters: number;
  readonly onRenderAnyway: () => void;
}

export function LargeMarkdownNotice({ characters, onRenderAnyway }: LargeMarkdownNoticeProps) {
  const t = useTranslate();
  const megabytes = (characters / (1024 * 1024)).toFixed(1);

  return (
    <div className="viewer-fallback large-markdown-notice">
      <FileWarning size={48} strokeWidth={1.5} className="viewer-fallback__icon" aria-hidden="true" />
      <h2 className="viewer-fallback__title">{t('largeMarkdown.title')}</h2>
      <p className="viewer-fallback__detail">{t('largeMarkdown.detail', { size: megabytes })}</p>
      <button type="button" className="viewer-fallback__btn" onClick={onRenderAnyway}>
        {t('largeMarkdown.renderAnyway')}
      </button>
    </div>
  );
}
