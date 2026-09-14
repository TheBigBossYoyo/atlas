import { memo, useEffect, useRef, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { DOCUMENT_RENDER_TIMEOUT_MS, DOCUMENT_SIZE_CAP_BYTES, formatSizeCapMessage, withRenderTimeout } from './shared/documentGuard'
import { estimatePageCount } from './shared/pageEstimate'
import { countWords } from './shared/textStats'
import './__styles__/viewer-rtf.css'

function toArrayBuffer(file: ViewerProps['file']): ArrayBuffer {
  return file.kind === 'binary'
    ? file.content
    : new TextEncoder().encode(file.kind === 'text' ? file.content : '').buffer
}

async function renderRtf(arrayBuffer: ArrayBuffer): Promise<Node[]> {
  const { RTFJS, WMFJS, EMFJS } = await import('rtf.js')

  try {
    RTFJS.loggingEnabled(false)
    WMFJS.loggingEnabled(false)
    EMFJS.loggingEnabled(false)
  } catch {
    // logging API may differ — ignore
  }

  const doc = new RTFJS.Document(arrayBuffer, {})
  return doc.render()
}

function RtfViewerBase({ file }: ViewerProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [error, setError] = useState<string | null>(null)
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()

  useEffect(() => {
    setNavItems([])
  }, [file.path, setNavItems])

  useEffect(() => {
    const container = containerRef.current
    if (!container) return

    let cancelled = false

    // Nested in a callback (rather than at the effect's top level) so the
    // "refuse an oversized file" bail-out synchronizes with the same
    // external render/parse call the rest of this effect makes, instead of
    // reading as state derivable during render (mirrors CsvViewer's load
    // effect and useSpreadsheetWorkbook.ts).
    void (async () => {
      const arrayBuffer = toArrayBuffer(file)

      // T9/DAT-20 — refuse an oversized file before ever attempting
      // conversion; see documentGuard.ts for why the timeout below can't
      // substitute for this check.
      if (arrayBuffer.byteLength > DOCUMENT_SIZE_CAP_BYTES) {
        setError(formatSizeCapMessage(arrayBuffer.byteLength, 'rtf'))
        return
      }

      setError(null)

      try {
        const nodes = await withRenderTimeout(
          renderRtf(arrayBuffer),
          DOCUMENT_RENDER_TIMEOUT_MS,
          'Rendering this RTF file took too long and was stopped.',
        )

        if (cancelled) return

        container.innerHTML = ''
        for (const node of nodes) {
          container.appendChild(node)
        }

        const words = countWords(container)
        const pages = estimatePageCount(container.scrollHeight)

        setStats({ kind: 'document', words, pages })
      } catch (err) {
        if (cancelled) return
        setError(err instanceof Error ? err.message : String(err))
      }
    })()

    return () => {
      cancelled = true
    }
  }, [file, setStats])

  if (error !== null) {
    return (
      <div className="rtf-viewer rtf-viewer--error">
        Failed to render RTF: {error}
      </div>
    )
  }

  return <div ref={containerRef} className="rtf-viewer" />
}

export const RtfViewer = memo(RtfViewerBase)
