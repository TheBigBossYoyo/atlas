/**
 * Forcing every paginated page into the real DOM, for printing and PDF export.
 *
 * Extracted from `src/viewers/DocxViewer.tsx`. It depends on nothing else in that
 * component, which is what makes it a clean unit: `PageStack` virtualizes pages,
 * and both of the operations here read the live DOM, so both need the
 * virtualization suspended first.
 *
 * `flushSync` is load-bearing in both paths, not a precaution. A plain
 * `setForceRenderAll(true)` only SCHEDULES the render that mounts every page,
 * which would race the synchronous `window.print()` / `dispatchEvent` call right
 * after it and print (or export) placeholders.
 *
 * The two `window` events exist because `utils/export/pdf.ts`'s `exportDocxPdf`
 * clones `.docx-page-stack` from outside this component — it is invoked
 * generically by element id from `App.tsx`'s export menu and holds no reference
 * into this viewer's state, so it cannot call the setter. It dispatches instead,
 * and since `dispatchEvent` runs listeners synchronously and `flushSync` commits
 * inside one, every page is real by the time `exportDocxPdf` resumes.
 */
import { useCallback, useEffect, useState } from 'react'
import { flushSync } from 'react-dom'

export type DocxFullRender = {
  /** Passed to `PageStack` to bypass page virtualization while true. */
  readonly forceRenderAll: boolean
  readonly handlePrint: () => void
  /**
   * Mounts every page SYNCHRONOUSLY — for a caller that is about to read the DOM
   * in the same tick (outline navigation scrolling to a heading that may sit on a
   * page virtualization has not mounted). Pair it with `releaseAllPages`.
   */
  readonly forceAllPagesNow: () => void
  /**
   * Re-enables virtualization. Outline navigation deliberately calls this on the
   * NEXT animation frame rather than immediately, so a smooth scroll still has
   * its target mounted while it animates.
   */
  readonly releaseAllPages: () => void
}

export function useDocxFullRender(): DocxFullRender {
  const [forceRenderAll, setForceRenderAll] = useState(false)

  const handlePrint = useCallback(() => {
    if (typeof window === 'undefined') {
      return
    }
    // D23-PERF-2 — `window.print()`'s `@media print` pass reads whatever is
    // ACTUALLY in the DOM at the moment it's called; a virtualized page that is
    // currently a placeholder would print blank. See this module's header on why
    // `flushSync` and not a plain setState.
    flushSync(() => setForceRenderAll(true))
    // The print stylesheet in viewer-docx.css hides toolbar/comments/save UI and
    // forces a page break after each .docx-page so output matches on-screen
    // pagination.
    document.body.classList.add('atlas-printing')
    try {
      window.print()
    } finally {
      document.body.classList.remove('atlas-printing')
      setForceRenderAll(false)
    }
  }, [])

  useEffect(() => {
    const handleForceFullRender = (): void => {
      flushSync(() => setForceRenderAll(true))
    }
    const handleReleaseFullRender = (): void => {
      setForceRenderAll(false)
    }
    window.addEventListener('atlas:docx-force-full-render', handleForceFullRender)
    window.addEventListener('atlas:docx-release-full-render', handleReleaseFullRender)
    return () => {
      window.removeEventListener('atlas:docx-force-full-render', handleForceFullRender)
      window.removeEventListener('atlas:docx-release-full-render', handleReleaseFullRender)
    }
  }, [])

  const forceAllPagesNow = useCallback(() => {
    flushSync(() => setForceRenderAll(true))
  }, [])

  const releaseAllPages = useCallback(() => {
    setForceRenderAll(false)
  }, [])

  return { forceRenderAll, handlePrint, forceAllPagesNow, releaseAllPages }
}
