import { useCallback, useState } from 'react'

import type { Position, Range } from './Selection'

export type UseSelectionResult = Readonly<{
  range: Range | null
  setRange: (range: Range | null) => void
  collapseTo: (position: Position) => void
  extendTo: (position: Position) => void
}>

export function useSelection(): UseSelectionResult {
  const [range, setRangeState] = useState<Range | null>(null)

  const setRange = useCallback((nextRange: Range | null) => {
    setRangeState(nextRange)
  }, [])

  const collapseTo = useCallback((position: Position) => {
    setRangeState({
      anchor: position,
      focus: position,
    })
  }, [])

  const extendTo = useCallback((position: Position) => {
    setRangeState((currentRange) => {
      if (!currentRange) {
        return {
          anchor: position,
          focus: position,
        }
      }

      return {
        anchor: currentRange.anchor,
        focus: position,
      }
    })
  }, [])

  return {
    range,
    setRange,
    collapseTo,
    extendTo,
  }
}
