import { act, renderHook, waitFor } from '@testing-library/react'
import React from 'react'
import { describe, expect, it, vi } from 'vitest'
import type { NavItem, ViewerStats } from '../../../formats/types'
import { ViewerProvider } from '../ViewerContext'
import {
  useCanFindViewer,
  useGetExportableContent,
  useNavItems,
  useOpenViewerFind,
  useRegisterViewerFind,
  useRegisterViewerSave,
  useSetNavItems,
  useSetViewerDirty,
  useSetViewerStats,
  useViewerIsDirty,
  useViewerSave,
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

  // ---------------------------------------------------------------------------
  // P1.1 — document-session capability contract
  // ---------------------------------------------------------------------------

  it('P1.1: setDirty/isDirty round-trip', () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(
      () => ({ isDirty: useViewerIsDirty(), setDirty: useSetViewerDirty() }),
      { wrapper },
    )

    expect(result.current.isDirty).toBe(false)

    act(() => {
      result.current.setDirty(true)
    })

    expect(result.current.isDirty).toBe(true)
  })

  it('P1.1: save() resolves false when no viewer has registered a save implementation', async () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(() => useViewerSave(), { wrapper })

    await expect(result.current()).resolves.toBe(false)
  })

  it('P1.1: save() calls whatever the active viewer registered via registerSave', async () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(
      () => ({ save: useViewerSave(), registerSave: useRegisterViewerSave() }),
      { wrapper },
    )

    const viewerSave = vi.fn().mockResolvedValue(true)

    act(() => {
      result.current.registerSave(viewerSave)
    })

    await expect(result.current.save()).resolves.toBe(true)
    expect(viewerSave).toHaveBeenCalledTimes(1)
  })

  it('P1.1: registerSave(null) unregisters — save() falls back to resolving false', async () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(
      () => ({ save: useViewerSave(), registerSave: useRegisterViewerSave() }),
      { wrapper },
    )

    const viewerSave = vi.fn().mockResolvedValue(true)

    act(() => {
      result.current.registerSave(viewerSave)
    })
    act(() => {
      result.current.registerSave(null)
    })

    await expect(result.current.save()).resolves.toBe(false)
    expect(viewerSave).not.toHaveBeenCalled()
  })

  it('P1.1: isDirty resets to false when filePath changes (new file, fresh session)', () => {
    let filePath = '/a.docx'

    const { result, rerender } = renderHook(
      () => ({ isDirty: useViewerIsDirty(), setDirty: useSetViewerDirty() }),
      {
        wrapper: ({ children }: { children: React.ReactNode }) => (
          <ViewerProvider filePath={filePath}>{children}</ViewerProvider>
        ),
      },
    )

    act(() => {
      result.current.setDirty(true)
    })
    expect(result.current.isDirty).toBe(true)

    filePath = '/b.docx'
    rerender()

    expect(result.current.isDirty).toBe(false)
  })

  it('P1.1: getExportableContent is a typed placeholder that always resolves null today', () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(() => useGetExportableContent(), { wrapper })

    expect(result.current()).toBeNull()
  })

  it('P1.1: a viewer unmounting (e.g. switching files) clears its own save registration via cleanup', async () => {
    const registeredSave = vi.fn().mockResolvedValue(true)

    function ViewerThatRegistersSave() {
      const registerSave = useRegisterViewerSave()
      React.useEffect(() => {
        registerSave(registeredSave)
        return () => registerSave(null)
      }, [registerSave])
      return null
    }

    // Mirrors ViewerRouter: the viewer subtree is keyed on the open file, so
    // switching files unmounts the old viewer (and its cleanup effect) before
    // any new one mounts — exactly like DocxEditor registering/unregistering
    // its own save() as the user switches documents.
    let filePath: string | null = '/a.docx'
    const { result, rerender } = renderHook(() => useViewerSave(), {
      wrapper: ({ children }: { children: React.ReactNode }) => (
        <ViewerProvider filePath={filePath}>
          {filePath === '/a.docx' ? <ViewerThatRegistersSave /> : null}
          {children}
        </ViewerProvider>
      ),
    })

    await expect(result.current()).resolves.toBe(true)
    expect(registeredSave).toHaveBeenCalledTimes(1)

    // Simulate switching files: the registering viewer unmounts (its own
    // cleanup calls registerSave(null)) — save() must not keep calling the
    // stale handler from the file that's no longer open.
    filePath = '/b.docx'
    rerender()

    await waitFor(async () => {
      await expect(result.current()).resolves.toBe(false)
    })
    expect(registeredSave).toHaveBeenCalledTimes(1)
  })

  // ---------------------------------------------------------------------------
  // Wave 3-T — in-viewer find capability contract
  // ---------------------------------------------------------------------------

  it('canFind starts false and openFind() is a no-op when nothing is registered', () => {
    const wrapper = makeWrapper('/a.csv')
    const { result } = renderHook(
      () => ({ canFind: useCanFindViewer(), openFind: useOpenViewerFind() }),
      { wrapper },
    )

    expect(result.current.canFind).toBe(false)
    expect(() => result.current.openFind()).not.toThrow()
  })

  it('registerFind flips canFind to true and openFind() calls the registered implementation', () => {
    const wrapper = makeWrapper('/a.csv')
    const { result } = renderHook(
      () => ({
        canFind: useCanFindViewer(),
        registerFind: useRegisterViewerFind(),
        openFind: useOpenViewerFind(),
      }),
      { wrapper },
    )

    const viewerOpenFind = vi.fn()

    act(() => {
      result.current.registerFind(viewerOpenFind)
    })

    expect(result.current.canFind).toBe(true)

    act(() => {
      result.current.openFind()
    })

    expect(viewerOpenFind).toHaveBeenCalledTimes(1)
  })

  it('registerFind(null) unregisters — canFind returns to false and openFind() no-ops again', () => {
    const wrapper = makeWrapper('/a.csv')
    const { result } = renderHook(
      () => ({
        canFind: useCanFindViewer(),
        registerFind: useRegisterViewerFind(),
        openFind: useOpenViewerFind(),
      }),
      { wrapper },
    )

    const viewerOpenFind = vi.fn()

    act(() => {
      result.current.registerFind(viewerOpenFind)
    })
    act(() => {
      result.current.registerFind(null)
    })

    expect(result.current.canFind).toBe(false)

    act(() => {
      result.current.openFind()
    })
    expect(viewerOpenFind).not.toHaveBeenCalled()
  })

  it('canFind resets to false when filePath changes (new file, fresh session)', () => {
    let filePath = '/a.csv'

    const { result, rerender } = renderHook(
      () => ({ canFind: useCanFindViewer(), registerFind: useRegisterViewerFind() }),
      {
        wrapper: ({ children }: { children: React.ReactNode }) => (
          <ViewerProvider filePath={filePath}>{children}</ViewerProvider>
        ),
      },
    )

    act(() => {
      result.current.registerFind(() => {})
    })
    expect(result.current.canFind).toBe(true)

    filePath = '/b.csv'
    rerender()

    expect(result.current.canFind).toBe(false)
  })
})
