import { act, renderHook } from '@testing-library/react'
import React from 'react'
import { describe, expect, it } from 'vitest'
import type { NavItem, ViewerStats } from '../../../formats/types'
import { ViewerProvider } from '../ViewerContext'
import {
  useNavItems,
  useSetNavItems,
  useSetViewerStats,
  useViewerStats,
} from '../useViewerContext'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeWrapper(filePath: string | null) {
  return function Wrapper({ children }: { children: React.ReactNode }) {
    return <ViewerProvider filePath={filePath}>{children}</ViewerProvider>
  }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('ViewerContext', () => {
  it('set + get nav items', () => {
    const wrapper = makeWrapper('/a.md')
    const { result } = renderHook(
      () => ({ items: useNavItems(), set: useSetNavItems() }),
      { wrapper },
    )

    const item: NavItem = { id: 'h1', label: 'Heading', level: 0, onSelect: () => {} }

    act(() => {
      result.current.set([item])
    })

    expect(result.current.items).toEqual([item])
  })

  it('set + get stats', () => {
    const wrapper = makeWrapper('/a.md')
    const { result } = renderHook(
      () => ({ stats: useViewerStats(), set: useSetViewerStats() }),
      { wrapper },
    )

    const stats: ViewerStats = { kind: 'markdown', words: 10, headings: 2 }

    act(() => {
      result.current.set(stats)
    })

    expect(result.current.stats).toEqual(stats)
  })

  it('resets nav items and stats when filePath changes', () => {
    let filePath = '/a.md'

    const { result, rerender } = renderHook(
      () => ({
        items: useNavItems(),
        setItems: useSetNavItems(),
        stats: useViewerStats(),
        setStats: useSetViewerStats(),
      }),
      {
        wrapper: ({ children }: { children: React.ReactNode }) => (
          <ViewerProvider filePath={filePath}>{children}</ViewerProvider>
        ),
      },
    )

    const item: NavItem = { id: 'h1', label: 'Heading', level: 0, onSelect: () => {} }
    const stats: ViewerStats = { kind: 'markdown', words: 5, headings: 1 }

    act(() => {
      result.current.setItems([item])
      result.current.setStats(stats)
    })

    expect(result.current.items).toHaveLength(1)
    expect(result.current.stats).not.toBeNull()

    // Change filePath and rerender
    filePath = '/b.md'
    rerender()

    expect(result.current.items).toEqual([])
    expect(result.current.stats).toBeNull()
  })
})
