/**
 * In-grid find/jump (spreadsheet + CSV "openFind()" capability).
 */
import { act, renderHook } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import { useGridFind } from '../useGridFind'

const ROWS: ReadonlyArray<ReadonlyArray<string>> = [
  ['Name', 'City'],
  ['Alice', 'Boston'],
  ['Bob', 'Austin'],
  ['Carol', 'Boston'],
]

describe('useGridFind', () => {
  it('starts closed with no query and no matches', () => {
    const { result } = renderHook(() => useGridFind(ROWS))

    expect(result.current.isOpen).toBe(false)
    expect(result.current.matchCount).toBe(0)
    expect(result.current.gridSelection).toBeUndefined()
  })

  it('open() sets isOpen; close() clears the query and selection', () => {
    const { result } = renderHook(() => useGridFind(ROWS))

    act(() => result.current.open())
    expect(result.current.isOpen).toBe(true)

    act(() => result.current.setQuery('boston'))
    expect(result.current.matchCount).toBe(2)

    act(() => result.current.close())
    expect(result.current.isOpen).toBe(false)
    expect(result.current.query).toBe('')
    expect(result.current.gridSelection).toBeUndefined()
  })

  it('finds every case-insensitive match across all cells and selects the first one', () => {
    const { result } = renderHook(() => useGridFind(ROWS))

    act(() => result.current.setQuery('boston'))

    expect(result.current.matchCount).toBe(2)
    expect(result.current.currentMatch).toBe(0)
    expect(result.current.gridSelection?.current?.cell).toEqual([1, 1])
  })

  it('goToMatch cycles forward and wraps around', () => {
    const { result } = renderHook(() => useGridFind(ROWS))

    act(() => result.current.setQuery('boston'))
    act(() => result.current.goToMatch('next'))

    expect(result.current.currentMatch).toBe(1)
    expect(result.current.gridSelection?.current?.cell).toEqual([1, 3])

    act(() => result.current.goToMatch('next'))
    expect(result.current.currentMatch).toBe(0)
  })

  it('goToMatch cycles backward and wraps around', () => {
    const { result } = renderHook(() => useGridFind(ROWS))

    act(() => result.current.setQuery('boston'))
    act(() => result.current.goToMatch('prev'))

    expect(result.current.currentMatch).toBe(1)
  })

  it('reports zero matches and no selection for a query that matches nothing', () => {
    const { result } = renderHook(() => useGridFind(ROWS))

    act(() => result.current.setQuery('nowhere'))

    expect(result.current.matchCount).toBe(0)
    expect(result.current.gridSelection).toBeUndefined()
  })
})
