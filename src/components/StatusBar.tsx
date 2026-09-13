import { assertNever, type ViewerStats } from '../formats/types';
import { useViewerStats } from '../viewers/shared/useViewerContext';

interface StatusBarProps {
  fileName: string | null;
  isDirty: boolean;
}

function renderStats(stats: ViewerStats, formatter: Intl.NumberFormat): string {
  switch (stats.kind) {
    case 'markdown':
      return `${formatter.format(stats.words)} words · ${formatter.format(stats.headings)} headings`;
    case 'spreadsheet':
      return `Sheet: ${stats.sheet} · ${formatter.format(stats.rows)} rows × ${formatter.format(stats.cols)} cols`;
    case 'pdf':
      return `Page ${formatter.format(stats.page)} / ${formatter.format(stats.pageCount)}`;
    case 'slides':
      return `Slide ${formatter.format(stats.slide)} / ${formatter.format(stats.slideCount)}`;
    case 'code':
      return `${stats.language} · ${formatter.format(stats.lines)} lines`;
    case 'text':
      return `${formatter.format(stats.lines)} lines · ${formatter.format(stats.chars)} chars`;
    case 'document':
      return `${formatter.format(stats.words)} words · ${formatter.format(stats.pages)} pages`;
  }

  const exhaustiveCheck: never = stats;
  return assertNever(exhaustiveCheck);
}

export function StatusBar({ fileName, isDirty }: StatusBarProps) {
  const stats = useViewerStats();
  const formatter = new Intl.NumberFormat();
  const statsLabel = stats ? renderStats(stats, formatter) : null;

  return (
    <footer className="statusbar">
      <div className="statusbar__left">
        <span className="statusbar__file" title={fileName ?? 'Untitled'}>
          {fileName ?? 'Untitled'}
        </span>
        {isDirty && (
          <span className="statusbar__dirty" title="Unsaved changes" aria-label="Unsaved changes">●</span>
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
