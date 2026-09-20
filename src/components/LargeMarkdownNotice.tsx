import { FileWarning } from 'lucide-react';

import { useTranslate } from '../i18n';

/**
 * Markdown parsing (micromark, under react-markdown) degrades superlinearly:
 * measured on this project's own corpus, 0.5 MB takes ~2 s, 1 MB ~9 s and
 * 2 MB ~44 s, and a 5 MB document takes minutes at a ~2.6 GB peak.
 *
 * That cost no longer freezes the app: past `MARKDOWN_WORKER_BYTE_THRESHOLD`
 * (`components/markdown/sizeThresholds.ts`) — well below this notice's own
 * threshold — `MarkdownRenderer` parses off the main thread in a Worker (see
 * `components/markdown/markdownParse.worker.ts`), so a 0.5/1/2 MB document
 * now renders on its own, with a spinner, no click required. This notice's
 * threshold exists for what a Worker alone doesn't fix: the wall-clock wait
 * itself (tens of seconds to minutes) and the risk of exhausting the
 * Worker's own heap on a truly huge document — both still worth an opt-in
 * warning rather than an automatic multi-minute render. The editor pane
 * stays fast whatever the size either way.
 */
export const LARGE_MARKDOWN_PREVIEW_BYTES = 2_000_000;

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
