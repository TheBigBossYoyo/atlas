/**
 * Shared virtualized plain-text renderer (react-window), extracted from
 * TextViewer so CodeViewer can fall back to the same rendering path for very
 * large files instead of running shiki's synchronous WASM tokenizer over the
 * whole file (T2/DAT-13).
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import { List as FixedSizeList, type RowComponentProps } from 'react-window'

const ROW_HEIGHT = 20
const DEFAULT_HEIGHT = 600

type TextRowProps = { lines: readonly string[]; lineClassName: string }

export type VirtualizedPlainTextProps = {
  readonly lines: ReadonlyArray<string>
  /** Class on the scrolling container (e.g. `text-viewer`, `code-viewer`). */
  readonly className?: string
  /** Class on each rendered line (e.g. `text-viewer__line`). */
  readonly lineClassName: string
}

export function VirtualizedPlainText({ lines, className, lineClassName }: VirtualizedPlainTextProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [height, setHeight] = useState(DEFAULT_HEIGHT)

  useEffect(() => {
    const element = containerRef.current
    if (!element) {
      return
    }

    const updateHeight = () => {
      setHeight(element.clientHeight || DEFAULT_HEIGHT)
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

  const Row = useCallback(
    ({ index, style, ariaAttributes, lines, lineClassName }: RowComponentProps<TextRowProps>) => (
      <div {...ariaAttributes} className={lineClassName} style={{ ...style, whiteSpace: 'pre' }}>
        {lines[index]}
      </div>
    ),
    [],
  )

  return (
    <div ref={containerRef} className={className} style={{ height: '100%', width: '100%' }}>
      <FixedSizeList
        defaultHeight={DEFAULT_HEIGHT}
        rowComponent={Row}
        rowCount={lines.length}
        rowHeight={ROW_HEIGHT}
        rowProps={{ lines, lineClassName }}
        style={{ height, width: '100%' }}
      >
      </FixedSizeList>
    </div>
  )
}
