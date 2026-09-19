/**
 * D23-PERF-2 — pure geometry/lookup helpers behind page virtualization.
 * Kept dependency-free from React so they're trivial to unit-test and so
 * `PageStack`'s scroll-driven state management stays a thin wrapper around
 * plain functions instead of re-deriving this math inline.
 *
 * Why virtualize at all: `paginate()` returns a brand-new `Page` object for
 * every page on every call (see `paginate.ts`), so even a single-character
 * edit invalidates every `PageView`'s `page` prop by reference — `React.memo`
 * (D23-PERF, c0d969e) cannot bail any of them out, and each one mounts a
 * full run/table/footnote DOM subtree. Measured on the 35-page typing perf
 * fixture (tests/e2e/perf.spec.ts), the React commit + browser layout/paint
 * for all ~24 on-screen pages dominated a keystroke's cost far more than
 * pagination itself (which the D23 line cache already makes cheap — see
 * this task's before/after numbers in the D23-PERF-2 commit). Rendering only
 * the pages actually near the viewport (plus a small buffer, plus any page
 * that must stay mounted for correctness — see `PageStack.tsx`) cuts that
 * DOM work roughly in proportion to (visible pages / total pages).
 */
import type { Page } from '../layout/pageTypes'

// CSS px is 96dpi, pt is 72dpi (matches PageView.tsx's own PT_TO_PX).
export const PT_TO_PX = 4 / 3

// `.docx-page-stack`'s own `padding`/`gap` (page-view.css) — duplicated here
// (rather than read from the stylesheet, which JS can't do cheaply) so the
// offsets below land exactly where the real flex layout will place each
// page. If that CSS ever changes, this must change with it — both live in
// `src/docx/render/`, a few lines apart, specifically to keep that easy to
// notice.
export const PAGE_STACK_PADDING_PX = 24
export const PAGE_STACK_GAP_PX = 24

export type PageOffset = {
  readonly top: number
  readonly height: number
}

export type PageOffsetsResult = {
  readonly offsets: ReadonlyArray<PageOffset>
  readonly totalHeight: number
}

/**
 * Precomputes each page's top offset and height (in the same px units
 * `PageView` renders its `.docx-page` box at) from pagination's own output,
 * without touching the DOM. `.docx-page-stack` is a `flex-direction: column`
 * container with a fixed `padding`/`gap` (see the constants above), so this
 * is exactly the box position the browser will lay each page out at,
 * independent of which pages are actually mounted.
 */
export function computePageOffsets(pages: ReadonlyArray<Page>, zoom: number): PageOffsetsResult {
  const scale = zoom * PT_TO_PX
  const offsets: PageOffset[] = []
  let cursor = PAGE_STACK_PADDING_PX

  for (const page of pages) {
    const height = page.sizePt.height * scale
    offsets.push({ top: cursor, height })
    cursor += height + PAGE_STACK_GAP_PX
  }

  const totalHeight = pages.length === 0 ? 0 : cursor - PAGE_STACK_GAP_PX + PAGE_STACK_PADDING_PX
  return { offsets, totalHeight }
}

export type PageRange = {
  readonly start: number
  readonly end: number
}

/** An empty range — `end < start` — meaning "render nothing from the viewport window". */
const EMPTY_RANGE: PageRange = { start: 0, end: -1 }

/**
 * Binary-searches `offsets` (sorted, non-overlapping, ascending — exactly
 * what `computePageOffsets` produces) for the inclusive index range whose
 * boxes intersect `[scrollTop - bufferPx, scrollTop + viewportHeight +
 * bufferPx]`. `bufferPx` renders a little extra above/below the visible
 * area so a small scroll or a fast flick doesn't show a bare placeholder
 * flash before the next scroll-driven recompute catches up.
 */
export function findVisiblePageRange(
  offsets: ReadonlyArray<PageOffset>,
  scrollTop: number,
  viewportHeight: number,
  bufferPx: number,
): PageRange {
  if (offsets.length === 0) {
    return EMPTY_RANGE
  }

  const lo = scrollTop - bufferPx
  const hi = scrollTop + Math.max(viewportHeight, 0) + bufferPx

  // First page whose bottom edge is >= lo. `result` defaults to
  // `offsets.length` (a past-the-end sentinel, not the last index) so that
  // when NO page's bottom reaches `lo` — the viewport has been scrolled
  // past every page — `start > end` below correctly yields an empty range
  // instead of falling back to "the last page".
  let start = 0
  {
    let low = 0
    let high = offsets.length - 1
    let result = offsets.length
    while (low <= high) {
      const mid = (low + high) >> 1
      const bottom = offsets[mid].top + offsets[mid].height
      if (bottom >= lo) {
        result = mid
        high = mid - 1
      } else {
        low = mid + 1
      }
    }
    start = result
  }

  // Last page whose top edge is <= hi.
  let end = offsets.length - 1
  {
    let low = 0
    let high = offsets.length - 1
    let result = -1
    while (low <= high) {
      const mid = (low + high) >> 1
      if (offsets[mid].top <= hi) {
        result = mid
        low = mid + 1
      } else {
        high = mid - 1
      }
    }
    end = result
  }

  if (end < start) {
    return EMPTY_RANGE
  }

  return { start, end }
}

/**
 * Which page index (if any) contains a paragraph/table addressed by
 * `paragraphPath`, matching a `Position.paragraphPath` from the editor's
 * Selection model. Mirrors the same addressing `DocxViewer.tsx`'s
 * `buildPageOfParagraph` already relies on for PAGE/NUMPAGES field
 * evaluation: for a single-section document `paragraphPath` is just
 * `[blockIndex, ...]` (section 0 implied); for a multi-section document it's
 * `[sectionIndex, blockIndex, ...]`. Only the first (block-index) segment is
 * matched — a caret inside a deeply-nested table cell still resolves to the
 * page holding that cell's owning top-level table/paragraph, which is all
 * `PageLineRef`/`PageTableRef` addressing exposes.
 *
 * Returns every matching page index (not just the first): a paragraph or
 * table split across a page break can genuinely have lines on more than one
 * page, and the caller (`PageStack`'s pinning logic) needs every one of them
 * kept mounted, not just wherever the paragraph starts — the caret could be
 * on a line that only exists on the later page.
 */
export function findPagesForParagraphPath(
  pages: ReadonlyArray<Page>,
  paragraphPath: ReadonlyArray<number> | undefined,
): ReadonlyArray<number> {
  if (paragraphPath === undefined || paragraphPath.length === 0) {
    return []
  }

  const [sectionIndex, blockIndex] =
    paragraphPath.length === 1 ? [0, paragraphPath[0]] : [paragraphPath[0], paragraphPath[1]]

  const matches: number[] = []

  for (const [pageIndex, page] of pages.entries()) {
    if (page.sectionIndex !== sectionIndex) {
      continue
    }

    const hasMatch = page.columns.some(
      (column) =>
        column.lines.some((line) => line.paragraphPath[0] === blockIndex) ||
        column.tables.some((table) => table.blockPath[0] === blockIndex),
    )

    if (hasMatch) {
      matches.push(pageIndex)
    }
  }

  return matches
}
