/**
 * Pure PDF find/search helpers (PDF-06/P6).
 *
 * The index itself is built from `page.getTextContent()`, independent of
 * whether a page currently has a rendered canvas/text-layer — a document's
 * text can be searched before (or without) any page ever being scrolled
 * into view. Actual on-screen highlighting is applied separately by the
 * component once a matched page's text layer exists in the DOM (see
 * `PdfPage.tsx`); this module only answers "where are the matches".
 */

export type PdfTextItem = { readonly str?: unknown }

export type PdfPageText = {
  readonly pageNumber: number
  readonly text: string
}

export type PdfMatch = {
  readonly pageNumber: number
  /** Character offset of the match within that page's concatenated text. */
  readonly index: number
  readonly length: number
}

/** Joins a `getTextContent()` result's items into one search-able string. */
export function extractPageText(items: ReadonlyArray<PdfTextItem>): string {
  return items
    .map((item) => (typeof item.str === 'string' ? item.str : ''))
    .join(' ')
}

/**
 * Case-insensitive substring search across all indexed pages, in page order.
 * Returns every occurrence (not just the first per page) so the find bar can
 * report an accurate total match count and step through them one at a time.
 */
export function findMatches(
  pages: ReadonlyArray<PdfPageText>,
  query: string,
): PdfMatch[] {
  const trimmed = query.trim()
  if (trimmed.length === 0) {
    return []
  }

  const lowerQuery = trimmed.toLowerCase()
  const matches: PdfMatch[] = []

  for (const page of pages) {
    const lowerText = page.text.toLowerCase()
    let fromIndex = 0

    for (;;) {
      const index = lowerText.indexOf(lowerQuery, fromIndex)
      if (index === -1) break
      matches.push({ pageNumber: page.pageNumber, index, length: trimmed.length })
      fromIndex = index + 1
    }
  }

  return matches
}

/** How many of `matches` belong to `pageNumber`, in document order, before
 * the match at `globalIndex` — i.e. that match's 0-based ordinal within its
 * own page. Used to map a "Nth overall match" to "Nth highlight mark on this
 * specific page" for scroll-into-view + active-highlight styling. */
export function localMatchIndexOnPage(
  matches: ReadonlyArray<PdfMatch>,
  globalIndex: number,
): number {
  if (globalIndex < 0 || globalIndex >= matches.length) {
    return -1
  }
  const target = matches[globalIndex]
  let local = 0
  for (let i = 0; i < globalIndex; i += 1) {
    if (matches[i].pageNumber === target.pageNumber) {
      local += 1
    }
  }
  return local
}

export function countMatchesOnPage(
  matches: ReadonlyArray<PdfMatch>,
  pageNumber: number,
): number {
  let count = 0
  for (const match of matches) {
    if (match.pageNumber === pageNumber) count += 1
  }
  return count
}
