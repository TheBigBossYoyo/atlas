import { memo, useCallback, useEffect, useRef, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'

function OdtViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    setNavItems([])
  }, [file.path, setNavItems])

  const renderOdt = useCallback(async (
    buffer: ArrayBuffer,
    container: HTMLDivElement,
    signal: { cancelled: boolean },
  ): Promise<void> => {
    const [{ odtToHtml }, { default: DOMPurify }] = await Promise.all([
      import('odf-kit/reader'),
      import('dompurify'),
    ])

    if (signal.cancelled) return

    const bytes = new Uint8Array(buffer)
    const rawHtml = odtToHtml(bytes, { fragment: true })

    if (signal.cancelled) return

    const safeHtml = DOMPurify.sanitize(rawHtml, {
      USE_PROFILES: { html: true },
      FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed'],
      FORBID_ATTR: ['onerror', 'onload', 'onclick'],
    })

    if (signal.cancelled) return

    container.innerHTML = safeHtml

    const words = container.innerText.trim().split(/\s+/).filter(Boolean).length
    setStats({ kind: 'document', words, pages: 1 })
  }, [setStats])

  useEffect(() => {
    const container = containerRef.current
    if (!container) return

    if (file.kind === 'text') {
      setError('OdtViewer received a text file; expected binary.')
      return
    }

    const signal = { cancelled: false }
    setError(null)

    renderOdt(file.content, container, signal).catch((err: unknown) => {
      if (signal.cancelled) return
      const message = err instanceof Error ? err.message : String(err)
      setError(message)
    })

    return () => {
      signal.cancelled = true
      container.innerHTML = ''
    }
  }, [file, renderOdt])

  if (error !== null) {
    return (
      <div className="odt-viewer odt-viewer--error">
        Failed to render ODT: {error}
      </div>
    )
  }

  return <div ref={containerRef} className="odt-viewer" />
}

export const OdtViewer = memo(OdtViewerBase)
