import React, { memo, useCallback, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { Document } from '../model/document';
import type { Page } from '../layout/pageTypes';
import type { Relationship } from '../parser/relationships';
import type { Theme } from '../parser/theme';
import { collectBookmarkNamesByParagraph, collectRunMetaByParagraph } from './pageMeta';
import { PageView } from './PageView';
import { computePageOffsets, findVisiblePageRange, PT_TO_PX } from './pageVirtualization';
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
  /**
   * D23-PERF-2 — the scrollable ancestor whose position decides which pages
   * actually mount (see `pageVirtualization.ts`'s module doc for why this
   * exists at all: `paginate()` hands back a brand-new `Page` object for
   * EVERY page on every call, so a single keystroke's re-render mounts a
   * full DOM subtree for every page regardless of what's on screen).
   *
   * Left `undefined` (the default), `PageStack` renders every page
   * unconditionally — the exact pre-virtualization behavior. This keeps
   * every existing standalone caller (`PageView.test.tsx`'s `PageStack`
   * tests, any future read-only preview with no scroll container to
   * measure) working with no change, and means virtualization only ever
   * activates where `DocxViewer.tsx` explicitly opts in by passing the
   * editor surface's own ref.
   */
  scrollContainerRef?: React.RefObject<HTMLElement | null>;
  /**
   * Page indices that must stay mounted no matter where the viewport is —
   * typically the page(s) holding the current caret/selection (native
   * `contentEditable` selection sync and typed input both need that page's
   * real DOM to exist) plus any page a caller is mid-way through revealing
   * (Find, heading navigation). Ignored when `scrollContainerRef` is
   * omitted (everything is already mounted in that case).
   */
  pinnedPageIndices?: ReadonlySet<number>;
  /**
   * Bypasses virtualization entirely and mounts every page — used
   * transiently by print and PDF export, both of which read pages straight
   * out of the live DOM (`window.print()`'s `@media print` pass, and
   * `exportDocxPdf`'s `.docx-page-stack` clone in `src/utils/export/pdf.ts`)
   * and would otherwise only capture whatever happened to be mounted.
   */
  forceRenderAll?: boolean;
};

const EMPTY_RELATIONSHIPS: ReadonlyArray<Relationship> = [];
const EMPTY_PINNED: ReadonlySet<number> = new Set();

/** How much extra space (in viewport heights) to render beyond what's
 * strictly visible, on each side — smooths out a normal scroll or a fast
 * flick without a bare placeholder ever flashing before the next
 * scroll-driven recompute catches up. */
const VIEWPORT_BUFFER_VH = 1;

type Viewport = {
  /** The scroll container's `scrollTop`, re-expressed in the SAME coordinate
   * space `computePageOffsets` uses (0 = the top of `.docx-page-stack`
   * itself) — see `measureViewport`'s own doc comment for why that
   * re-expressing is necessary at all. */
  readonly scrollTop: number;
  readonly height: number;
};

/**
 * `scrollContainerRef` is whichever ancestor actually scrolls
 * (`DocxViewer.tsx` passes its outer `.docx-viewer`, not the inner editor
 * surface — see that prop's own doc comment there), which normally has
 * OTHER content above `.docx-page-stack` inside it (the toolbar, an open
 * Find/Replace panel, an open header/footer editor panel — all siblings or
 * ancestors of the stack, all variable height). `container.scrollTop` alone
 * is measured from the top of the CONTAINER's own content, not the top of
 * the stack, so it isn't directly comparable to `computePageOffsets`'
 * offsets (which start at 0 for the stack's own top padding).
 * `stack.getBoundingClientRect().top - container.getBoundingClientRect().top
 * + container.scrollTop` recovers exactly that fixed offset regardless of
 * what's sitting above the stack or how tall it currently is, and
 * subtracting it converts `scrollTop` into the stack's own coordinate
 * space. Re-measured on every scroll/resize (not cached) since an open/
 * closed panel changes this offset without necessarily firing its own
 * resize on the scroll container.
 */
function measureViewport(container: HTMLElement | null, stack: HTMLElement | null): Viewport | null {
  if (container === null) {
    return null;
  }
  const stackOffset =
    stack === null ? 0 : stack.getBoundingClientRect().top - container.getBoundingClientRect().top + container.scrollTop;
  return { scrollTop: container.scrollTop - stackOffset, height: container.clientHeight };
}

/**
 * Stand-in for a page that isn't currently mounted: reserves the exact same
 * box (`.docx-page`'s own background/shadow/sizing — see page-view.css)
 * `PageView` would occupy, so scroll height, the scrollbar, and every other
 * page's position stay correct while this one is virtualized out. No props
 * beyond size: it never touches `document`/`theme`/relationships, so
 * toggling a page between placeholder and real `PageView` never runs any of
 * PageView's O(pageSize) rendering work until it's actually needed.
 */
const PagePlaceholder: React.FC<{ page: Page; zoom: number }> = memo(({ page, zoom }) => {
  const scale = zoom * PT_TO_PX;
  return (
    <div
      className="docx-page"
      data-page-index={page.pageIndex}
      data-virtualized="placeholder"
      style={{
        width: `${page.sizePt.width * scale}px`,
        height: `${page.sizePt.height * scale}px`,
      }}
    />
  );
});
PagePlaceholder.displayName = 'PagePlaceholder';

const PageStackComponent: React.FC<PageStackProps> = ({
  pages,
  zoom,
  document,
  theme,
  relationships,
  onResizeTableColumn,
  scrollContainerRef,
  pinnedPageIndices,
  forceRenderAll,
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

  // The `.docx-page-stack` div itself (see the returned JSX below) —
  // `measureViewport` needs it to translate the scroll container's
  // `scrollTop` into the stack's own coordinate space (see that function's
  // doc comment).
  const stackRef = useRef<HTMLDivElement | null>(null);

  // D23-PERF-2 — the container's own scroll position/size. `null` until
  // measured (or whenever no `scrollContainerRef` was given at all), which
  // `visibleRange` below treats as "render everything" — see that memo and
  // `scrollContainerRef`'s own prop doc comment for why that's the correct,
  // safe default rather than a transient placeholder flash. Starts `null`
  // unconditionally rather than reading `stackRef.current` here too (it
  // would always be `null` on a true first render anyway — refs aren't
  // populated until after commit, and reading one during render is exactly
  // what the `react-hooks/refs` lint rule exists to catch); the
  // `useLayoutEffect` right below measures for real immediately after that
  // first commit, before the browser ever paints.
  const [viewport, setViewport] = useState<Viewport | null>(null);

  useLayoutEffect(() => {
    const container = scrollContainerRef?.current ?? null;
    if (container === null) {
      return undefined;
    }

    setViewport(measureViewport(container, stackRef.current));

    // rAF-throttled: a scroll or resize can fire far more often than the
    // browser actually repaints, and re-deriving `visibleRange` (and the
    // resulting page mount/unmount churn) doesn't need to happen more than
    // once per frame.
    let rafId: number | null = null;
    const scheduleUpdate = (): void => {
      if (rafId !== null) {
        return;
      }
      rafId = requestAnimationFrame(() => {
        rafId = null;
        setViewport(measureViewport(container, stackRef.current));
      });
    };

    container.addEventListener('scroll', scheduleUpdate, { passive: true });
    // A sidebar toggle, window resize, or zoom change can all resize the
    // container without necessarily firing a `scroll` event on it.
    const resizeObserver = typeof ResizeObserver === 'function' ? new ResizeObserver(scheduleUpdate) : null;
    resizeObserver?.observe(container);

    return () => {
      container.removeEventListener('scroll', scheduleUpdate);
      resizeObserver?.disconnect();
      if (rafId !== null) {
        cancelAnimationFrame(rafId);
      }
    };
  }, [scrollContainerRef]);

  const { offsets } = useMemo(() => computePageOffsets(pages, zoom), [pages, zoom]);

  const visibleRange = useMemo(() => {
    if (scrollContainerRef === undefined || viewport === null) {
      // No container to measure against (a standalone/test caller — see the
      // prop doc comment), or not measured yet: render every page rather
      // than guess.
      return { start: 0, end: pages.length - 1 };
    }
    const bufferPx = viewport.height * VIEWPORT_BUFFER_VH;
    return findVisiblePageRange(offsets, viewport.scrollTop, viewport.height, bufferPx);
  }, [scrollContainerRef, viewport, offsets, pages.length]);

  const pinned = pinnedPageIndices ?? EMPTY_PINNED;

  const shouldRenderPage = useCallback(
    (index: number): boolean => {
      if (forceRenderAll === true || scrollContainerRef === undefined) {
        return true;
      }
      if (index >= visibleRange.start && index <= visibleRange.end) {
        return true;
      }
      return pinned.has(index);
    },
    [forceRenderAll, scrollContainerRef, visibleRange, pinned],
  );

  return (
    <div className="docx-page-stack" data-page-count={pages.length} ref={stackRef}>
      {pages.map((page, idx) =>
        shouldRenderPage(idx) ? (
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
        ) : (
          <PagePlaceholder key={`page-${page.sectionIndex}-${page.pageIndex}-${idx}`} page={page} zoom={zoom} />
        ),
      )}
    </div>
  );
};

// D23-PERF — see `PageView`'s own `memo` doc comment: this stops a parent
// re-render (e.g. `DocxViewer` re-rendering for a reason unrelated to
// pagination, like toolbar/selection state) from re-running this component
// — and, transitively, every `PageView` and its O(documentSize) memoized
// maps above — when none of `PageStack`'s own props actually changed.
export const PageStack = memo(PageStackComponent);
