import { useEffect } from 'react';
import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import type { ViewerStats } from '../../formats/types';
import { ViewerProvider } from '../../viewers/shared/ViewerContext';
import { useSetViewerStats } from '../../viewers/shared/useViewerContext';
import { StatusBar } from '../StatusBar';

// i18n — StatusBar's counts now go through `src/i18n`'s translate/interpolate
// pipeline, which formats numbers with an explicit `Intl.NumberFormat('en-US')`
// for the English locale (see `translate.ts`'s `numberFormatterFor`) rather
// than the environment's own unspecified default locale (previously
// `new Intl.NumberFormat()` with no locale argument, at either call site —
// this environment's own ICU default isn't `en-US`, e.g. it renders 1200 as
// "1 200"). Matching that explicit locale here keeps this test verifying the
// same thing regardless of the machine it runs on.
const formatNumber = (value: number) => new Intl.NumberFormat('en-US').format(value);
const expectStatusText = (value: string) => {
  expect(
    screen.getByText((_, element) => {
      if (!(element instanceof HTMLElement)) {
        return false;
      }

      return element.classList.contains('statusbar__item') && element.textContent === value;
    })
  ).toBeInTheDocument();
};

function StatusBarHarness({ stats }: { stats: ViewerStats | null }) {
  const setViewerStats = useSetViewerStats();

  useEffect(() => {
    setViewerStats(stats);
  }, [setViewerStats, stats]);

  return <StatusBar fileName="atlas.pdf" isDirty={false} />;
}

describe('StatusBar', () => {
  it('renders only filename when stats are unavailable', () => {
    render(
      <ViewerProvider filePath="/tmp/atlas.pdf">
        <StatusBarHarness stats={null} />
      </ViewerProvider>
    );

    expect(screen.getByText('atlas.pdf')).toBeInTheDocument();
    expect(screen.queryByText(/words|Page|Sheet|Slide|lines|chars/)).not.toBeInTheDocument();
  });

  it('renders markdown stats', () => {
    render(
      <ViewerProvider filePath="/tmp/readme.md">
        <StatusBarHarness stats={{ kind: 'markdown', words: 1200, headings: 8 }} />
      </ViewerProvider>
    );

    expectStatusText(`${formatNumber(1200)} words · ${formatNumber(8)} headings`);
  });

  it('renders spreadsheet stats', () => {
    render(
      <ViewerProvider filePath="/tmp/data.xlsx">
        <StatusBarHarness stats={{ kind: 'spreadsheet', sheet: 'Summary', rows: 1024, cols: 12 }} />
      </ViewerProvider>
    );

    expectStatusText(`Sheet: Summary · ${formatNumber(1024)} rows × ${formatNumber(12)} cols`);
  });

  it('renders pdf stats', () => {
    render(
      <ViewerProvider filePath="/tmp/file.pdf">
        <StatusBarHarness stats={{ kind: 'pdf', page: 4, pageCount: 24 }} />
      </ViewerProvider>
    );

    expectStatusText(`Page ${formatNumber(4)} / ${formatNumber(24)}`);
  });

  it('renders slides stats', () => {
    render(
      <ViewerProvider filePath="/tmp/deck.pptx">
        <StatusBarHarness stats={{ kind: 'slides', slide: 2, slideCount: 15 }} />
      </ViewerProvider>
    );

    expectStatusText(`Slide ${formatNumber(2)} / ${formatNumber(15)}`);
  });

  it('renders code stats', () => {
    render(
      <ViewerProvider filePath="/tmp/main.ts">
        <StatusBarHarness stats={{ kind: 'code', language: 'typescript', lines: 240 }} />
      </ViewerProvider>
    );

    expectStatusText(`typescript · ${formatNumber(240)} lines`);
  });

  it('renders text stats', () => {
    render(
      <ViewerProvider filePath="/tmp/notes.txt">
        <StatusBarHarness stats={{ kind: 'text', lines: 88, chars: 4096 }} />
      </ViewerProvider>
    );

    expectStatusText(`${formatNumber(88)} lines · ${formatNumber(4096)} chars`);
  });

  it('renders document stats', () => {
    render(
      <ViewerProvider filePath="/tmp/report.docx">
        <StatusBarHarness stats={{ kind: 'document', words: 5000, pages: 18 }} />
      </ViewerProvider>
    );

    expectStatusText(`${formatNumber(5000)} words · ${formatNumber(18)} pages`);
  });
});
