import { useCallback, useEffect, useRef, useState } from 'react'

import type { PageGeometry, PdfDocument } from './types'

const PREFETCH_CONCURRENCY = 6
/** Batch size before a state update, so a 500-page document doesn't cause
 * 500 separate re-renders while background geometry discovery runs. */
const PREFETCH_UPDATE_BATCH_SIZE = 20

/**
 * Tracks each page's unscaled (scale 1, rotation 0) size (PDF-09/P8 — needed
 * for correct per-page fit scale — and PDF-01/P1 — needed to size an
 * inactive page's placeholder so the scrollbar/scroll position stay stable
 * as pages render in and out of view).
 *
 * Any page that becomes active reports its own geometry immediately (see
 * `PdfPage.tsx`) — this hook's background prefetch only needs to catch up
 * for pages that never became active, so the document's overall layout
 * converges to fully-accurate quickly without blocking initial paint.
 */
export function usePdfPageGeometry(pdfDoc: PdfDocument | null, pageCount: number) {
  const [geometry, setGeometry] = useState<ReadonlyMap<number, PageGeometry>>(new Map())
  const geometryRef = useRef(geometry)
  geometryRef.current = geometry

  useEffect(() => {
    setGeometry(new Map())
  }, [pdfDoc])

  useEffect(() => {
    if (!pdfDoc || pageCount === 0) return undefined
    // Narrow once outside the closure below — see usePdfFind.ts for why.
    const doc = pdfDoc

    let cancelled = false
    let nextPage = 1
    let sinceLastUpdate = 0
    const collected = new Map<number, PageGeometry>()

    async function worker(): Promise<void> {
      for (;;) {
        if (cancelled) return
        const pageNumber = nextPage
        nextPage += 1
        if (pageNumber > pageCount) return
        if (geometryRef.current.has(pageNumber)) continue

        try {
          const page = await doc.getPage(pageNumber)
          if (cancelled) return
          const viewport = page.getViewport({ scale: 1, rotation: 0 })
          collected.set(pageNumber, { width: viewport.width, height: viewport.height })
        } catch {
          // Leave ungeometried; the fallback placeholder size is used until
          // this page becomes active and reports its own geometry.
        }

        sinceLastUpdate += 1
        if (sinceLastUpdate >= PREFETCH_UPDATE_BATCH_SIZE) {
          sinceLastUpdate = 0
          const merged = new Map(geometryRef.current)
          collected.forEach((value, key) => merged.set(key, value))
          collected.clear()
          setGeometry(merged)
        }
      }
    }

    void Promise.all(
      Array.from({ length: Math.min(PREFETCH_CONCURRENCY, pageCount) }, () => worker()),
    ).then(() => {
      if (!cancelled && collected.size > 0) {
        const merged = new Map(geometryRef.current)
        collected.forEach((value, key) => merged.set(key, value))
        setGeometry(merged)
      }
    })

    return () => {
      cancelled = true
    }
  }, [pdfDoc, pageCount])

  const reportGeometry = useCallback((pageNumber: number, pageGeometry: PageGeometry) => {
    setGeometry((prev) => {
      if (prev.has(pageNumber)) return prev
      const next = new Map(prev)
      next.set(pageNumber, pageGeometry)
      return next
    })
  }, [])

  return { geometry, reportGeometry }
}
