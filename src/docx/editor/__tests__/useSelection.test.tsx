import { act, renderHook } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import type { Position, Range } from '../Selection'
import { useSelection } from '../useSelection'

function makePosition(
  paragraphPath: ReadonlyArray<number>,
  runIndex: number,
  charOffset: number,
): Position {
  return {
    paragraphPath,
    runIndex,
    charOffset,
  }
}

function makeRange(anchor: Position, focus: Position): Range {
  return {
    anchor,
    focus,
  }
}

describe('useSelection', () => {
  it('starts with a null range', () => {
    const { result } = renderHook(() => useSelection())

    expect(result.current.range).toBeNull()
  })

  it('setRange stores an explicit range', () => {
    const { result } = renderHook(() => useSelection())
    const nextRange = makeRange(makePosition([0], 0, 1), makePosition([0], 0, 3))

    act(() => {
      result.current.setRange(nextRange)
    })

    expect(result.current.range).toEqual(nextRange)
  })

  it('collapseTo creates a collapsed selection at the given position', () => {
    const { result } = renderHook(() => useSelection())
    const position = makePosition([2], 1, 4)

    act(() => {
      result.current.collapseTo(position)
    })

    expect(result.current.range).toEqual({
      anchor: position,
      focus: position,
    })
  })

  it('extendTo preserves the anchor and updates the focus', () => {
    const { result } = renderHook(() => useSelection())
    const anchor = makePosition([1], 0, 2)
    const focus = makePosition([1], 0, 6)

    act(() => {
      result.current.collapseTo(anchor)
      result.current.extendTo(focus)
    })

    expect(result.current.range).toEqual({
      anchor,
      focus,
    })
  })

  it('extendTo collapses when no range exists yet', () => {
    const { result } = renderHook(() => useSelection())
    const position = makePosition([4], 2, 1)

    act(() => {
      result.current.extendTo(position)
    })

    expect(result.current.range).toEqual({
      anchor: position,
      focus: position,
    })
  })

  it('setRange can clear the selection back to null', () => {
    const { result } = renderHook(() => useSelection())

    act(() => {
      result.current.setRange(makeRange(makePosition([0], 0, 0), makePosition([0], 0, 2)))
    })

    act(() => {
      result.current.setRange(null)
    })

    expect(result.current.range).toBeNull()
  })
})
