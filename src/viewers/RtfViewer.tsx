import { memo, useEffect, useRef, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { DOCUMENT_RENDER_TIMEOUT_MS, DOCUMENT_SIZE_CAP_BYTES, formatSizeCapMessage, withRenderTimeout } from './shared/documentGuard'
import { estimatePageCount } from './shared/pageEstimate'
import { sanitizeDocumentHtml } from './shared/sanitizeDocumentHtml'
import { countWords } from './shared/textStats'
import './__styles__/viewer-rtf.css'

function toArrayBuffer(file: ViewerProps['file']): ArrayBuffer {
  return file.kind === 'binary'
    ? file.content
    : new TextEncoder().encode(file.kind === 'text' ? file.content : '').buffer
}

/**
 * Renders RTF to sanitized HTML (security fix, found during this review).
 *
 * rtf.js hands back live DOM `Node`s it builds itself from the file's own
 * (fully untrusted) content — including, for a `{\field{\*\fldinst
 * HYPERLINK "..."}}` field, a real `<a>` whose `href` is assigned directly
 * from the RTF's URL string with zero scheme validation (see rtf.js's
 * `Renderer.buildHyperlinkElement`: `link.href = url`). A crafted RTF file
 * can therefore render a clickable `javascript:`-URI link that runs
 * arbitrary script in this renderer when clicked — the same class of risk
 * OdtViewer already guards against by sanitizing odf-kit's output with
 * DOMPurify. Serializing rtf.js's node tree and sanitizing it the same way
 * closes that gap. rtf.js never produces `<canvas>` (images arrive as
 * `<img src="data:...">`), so nothing visual is lost by round-tripping
 * through markup instead of appending the live nodes directly.
 */
async function renderRtf(arrayBuffer: ArrayBuffer): Promise<string> {
  const [{ RTFJS, WMFJS, EMFJS }, { default: DOMPurify }] = await Promise.all([
    import('rtf.js'),
    import('dompurify'),
  ])

  try {
    RTFJS.loggingEnabled(false)
    WMFJS.loggingEnabled(false)
    EMFJS.loggingEnabled(false)
  } catch {
    // logging API may differ — ignore
  }

  const doc = new RTFJS.Document(arrayBuffer, {})
  const nodes = await doc.render()

  const wrapper = document.createElement('div')
  for (const node of nodes) {
    wrapper.appendChild(node)
  }

  return sanitizeDocumentHtml(wrapper.innerHTML, DOMPurify)
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
        const html = await withRenderTimeout(
          renderRtf(arrayBuffer),
          DOCUMENT_RENDER_TIMEOUT_MS,
          'Rendering this RTF file took too long and was stopped.',
        )

        if (cancelled) return

        container.innerHTML = html

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
