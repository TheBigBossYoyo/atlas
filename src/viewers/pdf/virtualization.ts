/**
 * Pure virtualization math (PDF-01/P1, PDF-07/PDF-08/P7).
 *
 * Deliberately decoupled from `IntersectionObserver` itself (which can't run
 * meaningfully in a unit test without heavy DOM mocking) — the component
 * feeds this module plain data derived from observer callbacks, and this
 * module answers "which pages should have a live canvas right now" and
 * "which page is most visible right now", both of which are unit-testable
 * in isolation.
 */

export const DEFAULT_OVERSCAN = 2

export type PageRatioEntry = {
  readonly pageNumber: number
  readonly ratio: number
  readonly isIntersecting: boolean
}

/**
 * Incrementally merges one IntersectionObserver callback batch into a
 * running page→visibility-ratio map (PDF-08). The previous implementation
 * derived "the most visible page" from only the entries included in the
 * current callback batch, silently discarding every page NOT included in
 * that batch — during continuous scrolling this made the reported current
 * page flicker/regress because a page that scrolled out of view earlier
 * wasn't in the new batch to be "reset". Merging into a persistent map fixes
 * that: a page's last-known ratio sticks until the observer reports again
 * for that same page.
 */
export function updateVisibilityRatios(
  previous: ReadonlyMap<number, number>,
  entries: ReadonlyArray<PageRatioEntry>,
): Map<number, number> {
  const next = new Map(previous)

  for (const entry of entries) {
    if (entry.isIntersecting && entry.ratio > 0) {
      next.set(entry.pageNumber, entry.ratio)
    } else {
      next.delete(entry.pageNumber)
    }
  }

  return next
}

/** Picks the page with the highest current visibility ratio, or `fallback`
 * when nothing is currently visible (e.g. mid-programmatic-scroll). */
export function pickMostVisiblePage(
  ratios: ReadonlyMap<number, number>,
  fallback: number,
): number {
  let bestPage = fallback
  let bestRatio = 0

  for (const [pageNumber, ratio] of ratios) {
    if (ratio > bestRatio) {
      bestRatio = ratio
      bestPage = pageNumber
    }
  }

  return bestRatio > 0 ? bestPage : fallback
}

/**
 * Expands a set of currently-intersecting page numbers by `overscan` pages
 * in either direction and clamps to `[1, pageCount]` (PDF-01). Pages inside
 * the returned window should have a live rendered canvas; pages outside it
 * should have their canvas released to keep memory bounded on long
 * documents.
 */
export function computeActivePageWindow(
  intersectingPages: ReadonlyArray<number>,
  pageCount: number,
  overscan: number = DEFAULT_OVERSCAN,
): ReadonlySet<number> {
  const active = new Set<number>()
  if (pageCount <= 0) {
    return active
  }

  for (const page of intersectingPages) {
    const start = Math.max(1, page - overscan)
    const end = Math.min(pageCount, page + overscan)
    for (let candidate = start; candidate <= end; candidate += 1) {
      active.add(candidate)
    }
  }

  return active
}

/**
 * Ensures a page-jump target (bookmark click, page-number entry, Home/End)
 * is always part of the active window, even if it isn't intersecting yet —
 * otherwise `scrollIntoView` has nothing to scroll to on a page that has
 * never been rendered (PDF-07).
 */
export function includePendingJumpTarget(
  active: ReadonlySet<number>,
  pendingPage: number | null,
  pageCount: number,
): ReadonlySet<number> {
  if (pendingPage === null || pendingPage < 1 || pendingPage > pageCount) {
    return active
  }
  if (active.has(pendingPage)) {
    return active
  }

  const next = new Set(active)
  next.add(pendingPage)
  return next
}

/** Clamps a requested page number into `[1, pageCount]` (or `1` if the
 * document has no pages yet). */
export function clampPageNumber(pageNumber: number, pageCount: number): number {
  if (pageCount <= 0) return 1
  if (!Number.isFinite(pageNumber)) return 1
  return Math.min(pageCount, Math.max(1, Math.trunc(pageNumber)))
}
