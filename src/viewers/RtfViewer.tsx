import { memo, useEffect, useRef, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'

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

    const arrayBuffer: ArrayBuffer =
      file.kind === 'binary'
        ? file.content
        : new TextEncoder().encode(
            file.kind === 'text' ? file.content : '',
          ).buffer

    void (async () => {
      try {
        const { RTFJS, WMFJS, EMFJS } = await import('rtf.js')

        try {
          RTFJS.loggingEnabled(false)
          WMFJS.loggingEnabled(false)
          EMFJS.loggingEnabled(false)
        } catch {
          // logging API may differ — ignore
        }

        const doc = new RTFJS.Document(arrayBuffer, {})
        const nodes = await doc.render()

        if (cancelled) return

        container.innerHTML = ''
        for (const node of nodes) {
          container.appendChild(node)
        }

        const words = container.innerText
          .trim()
          .split(/\s+/)
          .filter(Boolean).length

        setStats({ kind: 'document', words, pages: 1 })
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
