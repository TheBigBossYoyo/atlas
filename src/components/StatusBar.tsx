import { assertNever, type ViewerStats } from '../formats/types';
import { useViewerStats } from '../viewers/shared/useViewerContext';
import { useTranslate } from '../i18n';
import type { TranslateFn } from '../i18n';

interface StatusBarProps {
  fileName: string | null;
  isDirty: boolean;
}

function renderStats(stats: ViewerStats, t: TranslateFn): string {
  switch (stats.kind) {
    case 'markdown':
      return t('statusBar.markdownStats', {
        words: t('statusBar.words', { count: stats.words }),
        headings: t('statusBar.headings', { count: stats.headings }),
      });
    case 'spreadsheet':
      return t('statusBar.spreadsheetStats', {
        sheet: stats.sheet,
        rows: t('statusBar.rows', { count: stats.rows }),
        cols: t('statusBar.cols', { count: stats.cols }),
      });
    case 'pdf':
      return t('statusBar.pdfStats', { page: stats.page, pageCount: stats.pageCount });
    case 'slides':
      return t('statusBar.slidesStats', { slide: stats.slide, slideCount: stats.slideCount });
    case 'code':
      return t('statusBar.codeStats', { language: stats.language, lines: t('statusBar.lines', { count: stats.lines }) });
    case 'text':
      return t('statusBar.textStats', {
        lines: t('statusBar.lines', { count: stats.lines }),
        chars: t('statusBar.chars', { count: stats.chars }),
      });
    case 'document':
      return t('statusBar.documentStats', {
        words: t('statusBar.words', { count: stats.words }),
        pages: t('statusBar.pages', { count: stats.pages }),
      });
  }

  const exhaustiveCheck: never = stats;
  return assertNever(exhaustiveCheck);
}

// A11Y pass 3 — only page/slide position changes (an explicit jump/next/prev,
// never a per-keystroke count) get announced; see this file's own doc
// comment on `announcedPositionLabel` for why the other `ViewerStats` kinds
// are deliberately excluded from the live region.
function announcedPositionLabel(stats: ViewerStats, t: TranslateFn): string | null {
  switch (stats.kind) {
    case 'pdf':
    case 'slides':
      return renderStats(stats, t);
    default:
      return null;
  }
}

export function StatusBar({ fileName, isDirty }: StatusBarProps) {
  const stats = useViewerStats();
  const t = useTranslate();
  const statsLabel = stats ? renderStats(stats, t) : null;
  // Deliberately a SEPARATE, narrower live region from the visible stats
  // item below rather than making `.statusbar__item` itself `aria-live`:
  // markdown's word count and the spreadsheet's row/col count are part of
  // the same `ViewerStats` union but change on every keystroke, and a live
  // region there would read exactly like the "wall of chatter" this pass is
  // supposed to avoid. Page/slide numbers only change on an explicit
  // navigation, so those alone are safe to announce.
  const announcedPosition = stats ? announcedPositionLabel(stats, t) : null;

  return (
    <footer className="statusbar">
      <div className="statusbar__left">
        <span className="statusbar__file" title={fileName ?? t('statusBar.untitled')}>
          {fileName ?? t('statusBar.untitled')}
        </span>
        {isDirty && (
          <span className="statusbar__dirty" title={t('statusBar.unsavedChangesTitle')} aria-label={t('common.unsavedChanges')}>●</span>
        )}
      </div>
      {statsLabel ? (
        <div className="statusbar__right">
          <span className="statusbar__item">{statsLabel}</span>
        </div>
      ) : null}
      <span className="visually-hidden" role="status" aria-live="polite">
        {announcedPosition}
      </span>
    </footer>
  );
}
