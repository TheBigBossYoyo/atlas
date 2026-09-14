import { memo, useEffect, useMemo } from 'react'

import type { ViewerProps, ViewerStats } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { VirtualizedPlainText } from './shared/VirtualizedPlainText'

function TextViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()

  const text = useMemo(() => {
    if (file.kind === 'text') {
      return file.content
    }

    return new TextDecoder('utf-8').decode(file.content)
  }, [file.content, file.kind])

  const lines = useMemo(() => text.split('\n'), [text])

  const stats = useMemo<ViewerStats>(
    () => ({ kind: 'text', lines: lines.length, chars: text.length }),
    [lines.length, text.length],
  )

  useEffect(() => {
    setStats(stats)
  }, [setStats, lines.length, stats, text.length])

  useEffect(() => {
    setNavItems([])
  }, [file.path, setNavItems])

  return <VirtualizedPlainText lines={lines} className="text-viewer" lineClassName="text-viewer__line" />
}

export const TextViewer = memo(TextViewerBase)
