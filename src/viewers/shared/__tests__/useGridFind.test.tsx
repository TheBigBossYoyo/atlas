/**
 * In-grid find/jump (spreadsheet + CSV "openFind()" capability).
 */
import type { ReactNode } from 'react'
import { act, renderHook } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import { ShortcutManagerProvider } from '../../../hooks/ShortcutManagerProvider'
import { useGridFind } from '../useGridFind'

function withShortcutManager({ children }: { children: ReactNode }) {
  return <ShortcutManagerProvider>{children}</ShortcutManagerProvider>
}

function dispatchKeydown(init: KeyboardEventInit) {
  window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }))
}

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

  // wave-3 shell-polish follow-up: Ctrl+F/Escape now register through the
  // shell's centralized shortcut dispatcher instead of doing nothing (the
  // ViewerContext `registerFind` capability alone was never wired to a key).
  describe('Ctrl+F / Escape via the shell shortcut dispatcher', () => {
    it('opens on Ctrl+F and closes on a second Ctrl+F', () => {
      const { result } = renderHook(() => useGridFind(ROWS), { wrapper: withShortcutManager })

      act(() => dispatchKeydown({ key: 'f', ctrlKey: true }))
      expect(result.current.isOpen).toBe(true)

      act(() => dispatchKeydown({ key: 'f', ctrlKey: true }))
      expect(result.current.isOpen).toBe(false)
    })

    it('closes on Escape while open, clearing the query', () => {
      const { result } = renderHook(() => useGridFind(ROWS), { wrapper: withShortcutManager })

      act(() => dispatchKeydown({ key: 'f', ctrlKey: true }))
      act(() => result.current.setQuery('boston'))
      expect(result.current.isOpen).toBe(true)

      act(() => dispatchKeydown({ key: 'Escape' }))
      expect(result.current.isOpen).toBe(false)
      expect(result.current.query).toBe('')
    })

    it('does nothing without a ShortcutManagerProvider ancestor (silently no-ops)', () => {
      const { result } = renderHook(() => useGridFind(ROWS))

      act(() => dispatchKeydown({ key: 'f', ctrlKey: true }))
      expect(result.current.isOpen).toBe(false)
    })
  })
})
