import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

import { extractPageText, findMatches, type PdfMatch, type PdfPageText } from './search'
import type { PdfDocument } from './types'

const INDEXING_BATCH_CONCURRENCY = 4
/** How many newly-indexed pages accumulate before a state update — bounds
 * re-renders on very long documents while still updating progressively. */
const INDEXING_UPDATE_BATCH_SIZE = 15

export type UsePdfFindResult = {
  readonly isOpen: boolean
  readonly query: string
  readonly matches: ReadonlyArray<PdfMatch>
  readonly currentMatchIndex: number
  readonly isIndexing: boolean
  readonly indexedPageCount: number
  readonly setQuery: (query: string) => void
  readonly open: () => void
  readonly close: () => void
  readonly next: () => void
  readonly prev: () => void
}

/**
 * PDF find/search state (PDF-06/P6). Indexes `getTextContent()` across pages
 * lazily in the background (only once find is actually opened, and without
 * blocking rendering), independent of which pages currently have a rendered
 * canvas — see `search.ts` for the pure match-finding logic this wraps.
 */
export function usePdfFind(
  pdfDoc: PdfDocument | null,
  pageCount: number,
  onNavigateToPage: (pageNumber: number) => void,
): UsePdfFindResult {
  const [isOpen, setIsOpen] = useState(false)
  const [query, setQueryState] = useState('')
  const [currentMatchIndex, setCurrentMatchIndex] = useState(0)
  const [pageTexts, setPageTexts] = useState<ReadonlyArray<PdfPageText>>([])
  const [isIndexing, setIsIndexing] = useState(false)

  const indexingStartedRef = useRef(false)

  // Lazily index every page's text content once find is opened. Runs once
  // per document (indexingStartedRef), with bounded concurrency so a
  // 500-page document doesn't fire 500 simultaneous getTextContent() calls.
  useEffect(() => {
    if (!isOpen || !pdfDoc || pageCount === 0 || indexingStartedRef.current) {
      return
    }
    indexingStartedRef.current = true
    // Narrow once, outside the closures below — TS narrowing doesn't cross
    // function boundaries, so `pdfDoc` itself would still read as nullable
    // inside `worker()`.
    const doc = pdfDoc

    let cancelled = false
    setIsIndexing(true)

    void (async () => {
      const results: PdfPageText[] = []
      let nextPage = 1

      async function worker(): Promise<void> {
        for (;;) {
          if (cancelled) return
          const pageNumber = nextPage
          nextPage += 1
          if (pageNumber > pageCount) return

          try {
            const page = await doc.getPage(pageNumber)
            if (cancelled) return
            const textContent = await page.getTextContent()
            if (cancelled) return
            results.push({
              pageNumber,
              text: extractPageText(textContent.items as ReadonlyArray<{ str?: unknown }>),
            })
          } catch {
            // A single unreadable page shouldn't abort indexing the rest.
            results.push({ pageNumber, text: '' })
          }

          if (results.length % INDEXING_UPDATE_BATCH_SIZE === 0) {
            setPageTexts([...results].sort((a, b) => a.pageNumber - b.pageNumber))
          }
        }
      }

      const workers = Array.from(
        { length: Math.min(INDEXING_BATCH_CONCURRENCY, pageCount) },
        () => worker(),
      )
      await Promise.all(workers)

      if (!cancelled) {
        setPageTexts([...results].sort((a, b) => a.pageNumber - b.pageNumber))
        setIsIndexing(false)
      }
    })()

    return () => {
      cancelled = true
    }
  }, [isOpen, pdfDoc, pageCount])

  // A new document invalidates any previous index.
  useEffect(() => {
    indexingStartedRef.current = false
    setPageTexts([])
    setIsIndexing(false)
  }, [pdfDoc])

  const matches = useMemo(() => findMatches(pageTexts, query), [pageTexts, query])

  useEffect(() => {
    setCurrentMatchIndex(0)
  }, [query])

  useEffect(() => {
    const active = matches[currentMatchIndex]
    if (active) {
      onNavigateToPage(active.pageNumber)
    }
    // Only re-navigate when the SELECTED match changes, not on every
    // onNavigateToPage identity change.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matches, currentMatchIndex])

  const setQuery = useCallback((next: string) => {
    setQueryState(next)
  }, [])

  const open = useCallback(() => {
    setIsOpen(true)
  }, [])

  const close = useCallback(() => {
    setIsOpen(false)
    setQueryState('')
  }, [])

  const next = useCallback(() => {
    setCurrentMatchIndex((prev) => (matches.length === 0 ? 0 : (prev + 1) % matches.length))
  }, [matches.length])

  const prev = useCallback(() => {
    setCurrentMatchIndex((prev) =>
      matches.length === 0 ? 0 : (prev - 1 + matches.length) % matches.length,
    )
  }, [matches.length])

  return {
    isOpen,
    query,
    matches,
    currentMatchIndex,
    isIndexing,
    indexedPageCount: pageTexts.length,
    setQuery,
    open,
    close,
    next,
    prev,
  }
}
