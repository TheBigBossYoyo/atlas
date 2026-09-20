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

export function StatusBar({ fileName, isDirty }: StatusBarProps) {
  const stats = useViewerStats();
  const t = useTranslate();
  const statsLabel = stats ? renderStats(stats, t) : null;

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
    </footer>
  );
}
