import React, { memo, useMemo } from 'react';
import type { Document } from '../model/document';
import type { Page } from '../layout/pageTypes';
import type { Relationship } from '../parser/relationships';
import type { Theme } from '../parser/theme';
import { collectBookmarkNamesByParagraph, collectRunMetaByParagraph } from './pageMeta';
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

const EMPTY_RELATIONSHIPS: ReadonlyArray<Relationship> = [];

const PageStackComponent: React.FC<PageStackProps> = ({
  pages,
  zoom,
  document,
  theme,
  relationships,
  onResizeTableColumn,
}) => {
  const resolvedRelationships = relationships ?? EMPTY_RELATIONSHIPS;
  // D23-PERF — computed ONCE per document/relationships pair here, instead of
  // once per PAGE inside `PageView` (each walking the whole document): with
  // N pages that was an O(N × documentSize) cost paid on every render this
  // stack's `document` prop changed, i.e. every keystroke. See `PageView`'s
  // `runMetaByParagraph`/`bookmarkNamesByParagraph` prop doc comment.
  const runMetaByParagraph = useMemo(
    () => collectRunMetaByParagraph(document, resolvedRelationships),
    [document, resolvedRelationships],
  );
  const bookmarkNamesByParagraph = useMemo(() => collectBookmarkNamesByParagraph(document), [document]);

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
          runMetaByParagraph={runMetaByParagraph}
          bookmarkNamesByParagraph={bookmarkNamesByParagraph}
        />
      ))}
    </div>
  );
};

// D23-PERF — see `PageView`'s own `memo` doc comment: this stops a parent
// re-render (e.g. `DocxViewer` re-rendering for a reason unrelated to
// pagination, like toolbar/selection state) from re-running this component
// — and, transitively, every `PageView` and its O(documentSize) memoized
// maps above — when none of `PageStack`'s own props actually changed.
export const PageStack = memo(PageStackComponent);
