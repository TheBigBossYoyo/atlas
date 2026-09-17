import { memo, useEffect, useMemo, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { LegacyFormatBanner } from './shared/LegacyFormatBanner'
import { estimatePageCount } from './shared/pageEstimate'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { VirtualizedPlainText } from './shared/VirtualizedPlainText'
import './__styles__/viewer-legacy-doc.css'

/** Matches `VirtualizedPlainText`'s own fixed row height — used only to turn a line count into an approximate rendered height for the page-count estimate below. */
const ROW_HEIGHT_PX = 20

function countWordsInParagraphs(paragraphs: ReadonlyArray<string>): number {
  let count = 0
  for (const paragraph of paragraphs) {
    if (paragraph.trim().length === 0) continue
    count += paragraph.trim().split(/\s+/).filter(Boolean).length
  }
  return count
}

/** Flattens paragraphs (which may embed `\n` soft line breaks) into display rows, with a blank separator row after each paragraph for readability. */
function toDisplayLines(paragraphs: ReadonlyArray<string>): ReadonlyArray<string> {
  const lines: string[] = []
  for (const paragraph of paragraphs) {
    lines.push(...paragraph.split('\n'))
    lines.push('')
  }
  return lines
}

/**
 * Read-only viewer for legacy Word 97-2003 (.doc) documents (wave-4
 * legacy-office). Renders the plain-paragraph structure `legacy/doc`
 * extracts — no formatting, images, or exact layout, by design (see that
 * module's header for exactly what is and isn't covered).
 */
function LegacyDocViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const [paragraphs, setParagraphs] = useState<ReadonlyArray<string>>([])
  const [error, setError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  useEffect(() => {
    setNavItems([])
  }, [file.path, setNavItems])

  useEffect(() => {
    setParagraphs([])
    setError(null)

    if (file.kind !== 'binary') {
      setIsLoading(false)
      return
    }

    let cancelled = false
    setIsLoading(true)

    // The actual extraction is synchronous (CPU-bound, no I/O) — dynamically
    // importing the parser (rather than a static top-level import) both
    // code-splits it into its own chunk, only ever fetched once a .doc is
    // actually opened, and gives this effect a genuine async boundary
    // before it touches state, so "Reading…" actually paints first instead
    // of the parse blocking the same tick that requested it.
    void (async () => {
      try {
        const { extractLegacyDocText } = await import('../legacy/doc')
        if (cancelled) return

        const result = extractLegacyDocText(new Uint8Array(file.content))
        if (cancelled) return

        setParagraphs(result.paragraphs)
        setIsLoading(false)
      } catch (err) {
        if (cancelled) return
        setError(err instanceof Error ? err.message : String(err))
        setIsLoading(false)
      }
    })()

    return () => {
      cancelled = true
    }
  }, [file])

  const lines = useMemo(() => toDisplayLines(paragraphs), [paragraphs])

  useEffect(() => {
    if (paragraphs.length === 0) {
      setStats(null)
      return
    }
    setStats({
      kind: 'document',
      words: countWordsInParagraphs(paragraphs),
      pages: estimatePageCount(lines.length * ROW_HEIGHT_PX),
    })
  }, [lines.length, paragraphs, setStats])

  if (file.kind !== 'binary') {
    return (
      <div className="legacy-doc-viewer legacy-doc-viewer--error">
        LegacyDocViewer received a text file; expected binary.
      </div>
    )
  }

  if (error !== null) {
    return (
      <div className="legacy-doc-viewer legacy-doc-viewer--error">
        Couldn't read this Word 97-2003 document: {error}
      </div>
    )
  }

  if (isLoading) {
    return <div className="legacy-doc-viewer">Reading legacy Word document…</div>
  }

  return (
    <div className="legacy-doc-viewer">
      <LegacyFormatBanner formatLabel="Word 97-2003 document (.doc)" modernExtension=".docx" />
      {paragraphs.length === 0 ? (
        <div className="legacy-doc-viewer__empty">This document has no readable text.</div>
      ) : (
        <VirtualizedPlainText lines={lines} className="legacy-doc-viewer__body" lineClassName="legacy-doc-viewer__line" />
      )}
    </div>
  )
}

export const LegacyDocViewer = memo(LegacyDocViewerBase)
