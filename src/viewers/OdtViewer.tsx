import { memo, useCallback, useEffect, useRef, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { DOCUMENT_RENDER_TIMEOUT_MS, DOCUMENT_SIZE_CAP_BYTES, formatSizeCapMessage, withRenderTimeout } from './shared/documentGuard'
import { estimatePageCount, parseLengthToPx } from './shared/pageEstimate'
import { countWords } from './shared/textStats'
import './__styles__/viewer-odt.css'

type OdtRenderResult = {
  readonly html: string
  readonly pageHeightPx: number | undefined
  /** Whether the document contains at least one tracked insertion/deletion/format change (T7/DAT-18). */
  readonly hasTrackedChanges: boolean
}

async function renderOdt(bytes: Uint8Array): Promise<OdtRenderResult> {
  const { readOdt } = await import('odf-kit/reader')

  // T7/DAT-18 — parse with tracked changes exposed (not silently
  // auto-accepted) so pending edits are visible to the reader instead of
  // indistinguishable from already-accepted content.
  const doc = readOdt(bytes, { trackedChanges: 'changes' })
  const html = doc.toHtml({ fragment: true, trackedChanges: 'changes' })

  return {
    html,
    pageHeightPx: parseLengthToPx(doc.pageLayout?.height),
    hasTrackedChanges: /<ins[\s>]|<del[\s>]|class="odf-format-change"/.test(html),
  }
}

function OdtViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [hasTrackedChanges, setHasTrackedChanges] = useState(false)

  useEffect(() => {
    setNavItems([])
  }, [file.path, setNavItems])

  const renderInto = useCallback(async (
    buffer: ArrayBuffer,
    container: HTMLDivElement,
    signal: { cancelled: boolean },
  ): Promise<void> => {
    const bytes = new Uint8Array(buffer)

    const [{ html, pageHeightPx, hasTrackedChanges: hasChanges }, { default: DOMPurify }] = await withRenderTimeout(
      Promise.all([renderOdt(bytes), import('dompurify')]),
      DOCUMENT_RENDER_TIMEOUT_MS,
      'Rendering this ODT file took too long and was stopped.',
    )

    if (signal.cancelled) return

    const safeHtml = DOMPurify.sanitize(html, {
      USE_PROFILES: { html: true },
      FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed'],
      FORBID_ATTR: ['onerror', 'onload', 'onclick'],
    })

    if (signal.cancelled) return

    container.innerHTML = safeHtml
    setHasTrackedChanges(hasChanges)

    const words = countWords(container)
    const pages = estimatePageCount(container.scrollHeight, pageHeightPx)
    setStats({ kind: 'document', words, pages })
  }, [setStats])

  useEffect(() => {
    const container = containerRef.current
    if (!container) return

    if (file.kind === 'text') {
      setError('OdtViewer received a text file; expected binary.')
      return
    }

    // T9/DAT-20 — refuse an oversized file before ever attempting
    // conversion; see documentGuard.ts for why the timeout below can't
    // substitute for this check.
    if (file.content.byteLength > DOCUMENT_SIZE_CAP_BYTES) {
      setError(formatSizeCapMessage(file.content.byteLength, 'odt'))
      return
    }

    const signal = { cancelled: false }
    setError(null)
    setHasTrackedChanges(false)

    renderInto(file.content, container, signal).catch((err: unknown) => {
      if (signal.cancelled) return
      const message = err instanceof Error ? err.message : String(err)
      setError(message)
    })

    return () => {
      signal.cancelled = true
      container.innerHTML = ''
    }
  }, [file, renderInto])

  if (error !== null) {
    return (
      <div className="odt-viewer odt-viewer--error">
        Failed to render ODT: {error}
      </div>
    )
  }

  return (
    <div className="odt-viewer">
      {hasTrackedChanges && (
        <div className="odt-viewer__tracked-changes-banner" role="status">
          This document has tracked changes — insertions and deletions are shown inline.
        </div>
      )}
      <div ref={containerRef} className="odt-viewer__body" />
    </div>
  )
}

export const OdtViewer = memo(OdtViewerBase)
