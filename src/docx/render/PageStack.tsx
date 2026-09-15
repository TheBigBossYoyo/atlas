import React from 'react';
import type { Document } from '../model/document';
import type { Page } from '../layout/pageTypes';
import type { Relationship } from '../parser/relationships';
import type { Theme } from '../parser/theme';
import { PageView } from './PageView';
import './__styles__/page-view.css';

export type PageStackProps = {
  pages: ReadonlyArray<Page>;
  zoom: number;
  document: Document;
  theme?: Theme;
  relationships?: ReadonlyArray<Relationship>;
  /** DXE-14 — column resize by dragging a table's column border. Omit to
   * render every page's tables with no resize handles at all (e.g. a
   * read-only preview) — see `PageView.tsx`'s own prop doc comment. */
  onResizeTableColumn?: (tablePath: ReadonlyArray<number>, columnIndex: number, widthTwips: number) => void;
};

export const PageStack: React.FC<PageStackProps> = ({
  pages,
  zoom,
  document,
  theme,
  relationships,
  onResizeTableColumn,
}) => {
  return (
    <div className="docx-page-stack">
      {pages.map((page, idx) => (
        <PageView
          key={`page-${page.sectionIndex}-${page.pageIndex}-${idx}`}
          page={page}
          zoom={zoom}
          document={document}
          theme={theme}
          relationships={relationships}
          onResizeTableColumn={onResizeTableColumn}
        />
      ))}
    </div>
  );
};
