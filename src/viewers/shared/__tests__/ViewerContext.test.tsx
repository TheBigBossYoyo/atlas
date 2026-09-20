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
  useRegisterViewerSaveAs,
  useReportSavedPath,
  useSetNavItems,
  useSetViewerDirty,
  useSetViewerStats,
  useViewerIsDirty,
  useViewerSave,
  useViewerSaveAs,
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
  // "Save As" side of the P1.1 document-session contract — added after
  // discovering (real-app testing) that the global Ctrl+Shift+S shortcut
  // silently did nothing for every non-markdown document: App.tsx's
  // saveFileAs() had no equivalent of save()'s viewerSaveRef fallback, so
  // there was nothing for a viewer to register into and nothing for the
  // shortcut to call. Mirrors the `save`/`registerSave` tests above.
  // ---------------------------------------------------------------------------

  it('saveAs() resolves false when no viewer has registered a saveAs implementation', async () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(() => useViewerSaveAs(), { wrapper })

    await expect(result.current()).resolves.toBe(false)
  })

  it('saveAs() calls whatever the active viewer registered via registerSaveAs', async () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(
      () => ({ saveAs: useViewerSaveAs(), registerSaveAs: useRegisterViewerSaveAs() }),
      { wrapper },
    )

    const viewerSaveAs = vi.fn().mockResolvedValue(true)

    act(() => {
      result.current.registerSaveAs(viewerSaveAs)
    })

    await expect(result.current.saveAs()).resolves.toBe(true)
    expect(viewerSaveAs).toHaveBeenCalledTimes(1)
  })

  it('registerSaveAs(null) unregisters — saveAs() falls back to resolving false', async () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(
      () => ({ saveAs: useViewerSaveAs(), registerSaveAs: useRegisterViewerSaveAs() }),
      { wrapper },
    )

    const viewerSaveAs = vi.fn().mockResolvedValue(true)

    act(() => {
      result.current.registerSaveAs(viewerSaveAs)
    })
    act(() => {
      result.current.registerSaveAs(null)
    })

    await expect(result.current.saveAs()).resolves.toBe(false)
    expect(viewerSaveAs).not.toHaveBeenCalled()
  })

  it('save() and saveAs() are independent registrations', async () => {
    const wrapper = makeWrapper('/a.docx')
    const { result } = renderHook(
      () => ({
        save: useViewerSave(),
        registerSave: useRegisterViewerSave(),
        saveAs: useViewerSaveAs(),
        registerSaveAs: useRegisterViewerSaveAs(),
      }),
      { wrapper },
    )

    const viewerSave = vi.fn().mockResolvedValue(true)
    const viewerSaveAs = vi.fn().mockResolvedValue(true)

    act(() => {
      result.current.registerSave(viewerSave)
      result.current.registerSaveAs(viewerSaveAs)
    })

    await result.current.saveAs()
    expect(viewerSaveAs).toHaveBeenCalledTimes(1)
    expect(viewerSave).not.toHaveBeenCalled()

    await result.current.save()
    expect(viewerSave).toHaveBeenCalledTimes(1)
    expect(viewerSaveAs).toHaveBeenCalledTimes(1)
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

  it('reports a Save As path up to the shell, so the tab can follow the document', () => {
    const onSavedPath = vi.fn()
    const { result } = renderHook(() => useReportSavedPath(), {
      wrapper: ({ children }: { children: React.ReactNode }) => (
        <ViewerProvider filePath="/old.docx" onSavedPath={onSavedPath}>
          {children}
        </ViewerProvider>
      ),
    })

    act(() => result.current('/old.docx', '/new.docx'))

    expect(onSavedPath).toHaveBeenCalledWith('/old.docx', '/new.docx')
  })

  it('stays callable (and harmless) when the shell registers no handler', () => {
    const wrapper = makeWrapper('/old.docx')
    const { result } = renderHook(() => useReportSavedPath(), { wrapper })

    expect(() => act(() => result.current('/old.docx', '/new.docx'))).not.toThrow()
  })

  // SAVE-1 — a Save As can resolve after the user has already switched to a
  // different tab: `filePath` (and thus `onSavedPath`, which App.tsx rebinds
  // to a fresh closure over whichever tab is now active) changes *before*
  // `reportSavedPath` is ever called. The test above only proves the
  // callback fires while `filePath` is still the one that started the save —
  // exactly the case that was never broken. This proves the two things the
  // fix actually depends on: (1) `reportSavedPath`'s identity is stable
  // across the filePath change (a viewer's in-flight save closes over it
  // once, before any tab switch, and must still be able to call it), and (2)
  // it forwards to whichever `onSavedPath` is current *along with* the path
  // the save started from — so the shell (App.tsx's `handleViewerSavedPath`)
  // can tell this report is about the OLD document, not the new one, even
  // though its own closure now belongs to the new tab.
  it('still reports the correct startedFromPath after filePath (and onSavedPath) has since changed', () => {
    const onSavedPathForOldTab = vi.fn()
    const onSavedPathForNewTab = vi.fn()

    let filePath = '/old.docx'
    let onSavedPath = onSavedPathForOldTab

    const { result, rerender } = renderHook(() => useReportSavedPath(), {
      wrapper: ({ children }: { children: React.ReactNode }) => (
        <ViewerProvider filePath={filePath} onSavedPath={onSavedPath}>
          {children}
        </ViewerProvider>
      ),
    })

    // The old tab's viewer captured this stable function before the switch.
    const reportSavedPathFromOldTab = result.current

    // The user switches tabs — App.tsx remounts the provider's subtree onto
    // the new file and rebinds onSavedPath to a closure over the new tab.
    filePath = '/new.pptx'
    onSavedPath = onSavedPathForNewTab
    rerender()

    // The old tab's Save As now resolves, reporting the document it actually
    // saved — NOT the new one that happens to be showing.
    act(() => reportSavedPathFromOldTab('/old.docx', '/old-renamed.docx'))

    // It reaches the CURRENT handler (App.tsx's own design: the ref always
    // points at the latest committed onSavedPath, since only App.tsx knows
    // how to look a path up across tab switches) carrying the old document's
    // own identity, so App.tsx can route it correctly instead of assuming it
    // is about whatever is showing now.
    expect(onSavedPathForNewTab).toHaveBeenCalledWith('/old.docx', '/old-renamed.docx')
    expect(onSavedPathForOldTab).not.toHaveBeenCalled()
  })
})
