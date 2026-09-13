import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { List as FixedSizeList, type RowComponentProps } from 'react-window'

import type { ViewerProps, ViewerStats } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'

const ROW_HEIGHT = 20

function TextViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [height, setHeight] = useState(600)

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

  useEffect(() => {
    const element = containerRef.current
    if (!element) {
      return
    }

    const updateHeight = () => {
      setHeight(element.clientHeight || 600)
    }

    updateHeight()

    if (typeof ResizeObserver === 'undefined') {
      return
    }

    const observer = new ResizeObserver(() => {
      updateHeight()
    })

    observer.observe(element)

    return () => {
      observer.disconnect()
    }
  }, [])

  type TextRowProps = { lines: readonly string[] }

  const Row = useCallback(
    ({ index, style, ariaAttributes, lines }: RowComponentProps<TextRowProps>) => (
      <div
        {...ariaAttributes}
        className="text-viewer__line"
        style={{ ...style, whiteSpace: 'pre' }}
      >
        {lines[index]}
      </div>
    ),
    [],
  )

  return (
    <div
      ref={containerRef}
      className="text-viewer"
      style={{ height: '100%', width: '100%' }}
    >
      <FixedSizeList
        defaultHeight={600}
        rowComponent={Row}
        rowCount={lines.length}
        rowHeight={ROW_HEIGHT}
        rowProps={{ lines }}
        style={{ height, width: '100%' }}
      >
      </FixedSizeList>
    </div>
  )
}

export const TextViewer = memo(TextViewerBase)
