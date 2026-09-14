/**
 * Pure outline/destination-resolution helpers (PDF-16/P12).
 *
 * Extracted unchanged from the pre-Wave-3-P `PdfViewer.tsx` so the existing
 * behavior (outline-based nav items, falling back to a flat page list when a
 * document has no outline) stays covered by unit tests instead of only ever
 * being exercised indirectly through the full component.
 */

import type { NavItem } from '../../formats/types'
import type { PdfDestination, PdfOutlineNode, PdfRef } from './types'

export function isPdfRef(value: unknown): value is PdfRef {
  if (typeof value !== 'object' || value === null) {
    return false
  }

  const maybeRef = value as Record<string, unknown>
  return typeof maybeRef.num === 'number' && typeof maybeRef.gen === 'number'
}

/**
 * Resolves a pdf.js outline-entry `dest` (a named destination string, an
 * explicit destination array, or null) to a 1-based page number.
 */
export async function resolveDestinationPage(
  getDestination: (id: string) => Promise<readonly unknown[] | null>,
  getPageIndex: (ref: PdfRef) => Promise<number>,
  dest: PdfDestination,
): Promise<number | null> {
  const resolvedDest = typeof dest === 'string' ? await getDestination(dest) : dest

  if (!resolvedDest || resolvedDest.length === 0) {
    return null
  }

  const target = resolvedDest[0]

  if (typeof target === 'number' && Number.isFinite(target)) {
    return target + 1
  }

  if (isPdfRef(target)) {
    return (await getPageIndex(target)) + 1
  }

  return null
}

export async function buildOutlineNavItems(
  outline: ReadonlyArray<PdfOutlineNode>,
  scrollToPage: (pageNumber: number) => void,
  getDestination: (id: string) => Promise<readonly unknown[] | null>,
  getPageIndex: (ref: PdfRef) => Promise<number>,
  level = 1,
): Promise<NavItem[]> {
  const items: NavItem[] = []

  for (const node of outline) {
    let pageNumber: number | null
    try {
      pageNumber = await resolveDestinationPage(getDestination, getPageIndex, node.dest)
    } catch {
      // A single malformed/dangling bookmark (a named destination that
      // throws, or a ref pointing at a page that no longer exists — both
      // seen in real-world, slightly-corrupted PDFs) must not abort
      // building the rest of the outline, and must not bubble up into the
      // document-load effect's catch block and brick loading the WHOLE
      // document over one broken link. Treat it the same as a destination
      // that legitimately resolved to nothing.
      pageNumber = null
    }
    const label = node.title.trim() || 'Untitled'

    items.push({
      id: pageNumber === null ? `${level}-${label}` : `page-${pageNumber}-${label}`,
      label,
      level,
      onSelect: () => {
        if (pageNumber !== null) {
          scrollToPage(pageNumber)
        }
      },
    })

    if (node.items.length > 0) {
      items.push(
        ...(await buildOutlineNavItems(
          node.items,
          scrollToPage,
          getDestination,
          getPageIndex,
          level + 1,
        )),
      )
    }
  }

  return items
}

export function buildFallbackNavItems(
  pageCount: number,
  scrollToPage: (pageNumber: number) => void,
): NavItem[] {
  return Array.from({ length: pageCount }, (_, index) => {
    const pageNumber = index + 1

    return {
      id: `page-${pageNumber}`,
      label: `Page ${pageNumber}`,
      level: 1,
      onSelect: () => scrollToPage(pageNumber),
    }
  })
}
