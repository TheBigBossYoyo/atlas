/**
 * USR-16 — the editing session shared by the PPTX and ODP editors.
 *
 * Owns everything that has nothing to do with a particular file format: the
 * in-memory package, undo/redo, a queue that serializes edits (each one must
 * see the previous one's result even while its re-parse is still running),
 * the dirty/save contract the shell drives with Ctrl+S, and the editor's own
 * undo/redo shortcuts. Each format supplies only how to parse its package
 * into slides and how to rewrite its parts.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'

import type { SlideData } from '../../shared/SlideDeck.types'
import { useRegisterViewerSave, useRegisterViewerSaveAs, useSetViewerDirty } from '../../shared/useViewerContext'
import { useViewerShortcuts } from '../../../hooks/useShortcutManager'
import { useUndoableState } from '../../spreadsheet/useUndoableState'
import { loadOfficePackage, writeOfficePackage, type OfficePackage } from '../../../office/officePackage'

export type DeckState = { readonly pkg: OfficePackage; readonly slides: ReadonlyArray<SlideData> }

const EMPTY_STATE: DeckState = { pkg: { parts: new Map() }, slides: [] }

export type SlideEditorCoreOptions = {
  readonly buffer: ArrayBuffer | null
  readonly filePath: string
  /** Parses the package into slides; `previous` lets a format reuse unchanged slides. */
  readonly parseDeck: (pkg: OfficePackage, previous: DeckState | null) => Promise<ReadonlyArray<SlideData>>
  readonly saveFilter: { readonly name: string; readonly extensions: ReadonlyArray<string> }
}

export type SlideEditorCore = {
  readonly status: 'loading' | 'ready' | 'error'
  readonly error: string | null
  readonly saveError: string | null
  readonly slides: ReadonlyArray<SlideData>
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly undo: () => void
  readonly redo: () => void
  readonly save: () => Promise<boolean>
  readonly saveAs: () => Promise<boolean>
  /** Applies one edit to the package; returns once the deck has been re-parsed. */
  readonly apply: (edit: (pkg: OfficePackage) => OfficePackage) => Promise<void>
  /** The live state, for queries that must not wait for a render (e.g. "can this slide take notes?"). */
  readonly current: () => DeckState
}

function baseName(filePath: string): string {
  return filePath.replace(/\\/g, '/').split('/').pop() || 'presentation'
}

export function useSlideEditorCore({ buffer, filePath, parseDeck, saveFilter }: SlideEditorCoreOptions): SlideEditorCore {
  const history = useUndoableState<DeckState>(EMPTY_STATE)
  const [status, setStatus] = useState<'loading' | 'ready' | 'error'>('loading')
  const [error, setError] = useState<string | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [savedPkg, setSavedPkg] = useState<OfficePackage | null>(null)
  const [savePath, setSavePath] = useState<string>(filePath)
  const presentRef = useRef(history.present)
  useLayoutEffect(() => {
    presentRef.current = history.present
  }, [history.present])
  const queueRef = useRef<Promise<void>>(Promise.resolve())
  const parseRef = useRef(parseDeck)
  useEffect(() => {
    parseRef.current = parseDeck
  }, [parseDeck])

  useEffect(() => {
    let cancelled = false
    void (async () => {
      setStatus('loading')
      if (!buffer) return
      try {
        const pkg = await loadOfficePackage(buffer)
        const slides = await parseRef.current(pkg, null)
        if (cancelled) return
        history.reset({ pkg, slides })
        setSavedPkg(pkg)
        setError(slides.length > 0 ? null : 'No slides found in this presentation.')
        setStatus(slides.length > 0 ? 'ready' : 'error')
      } catch (err: unknown) {
        if (cancelled) return
        setError(err instanceof Error ? err.message : String(err))
        setStatus('error')
      }
    })()
    return () => {
      cancelled = true
    }
    // history.reset is stable; reload only when the file bytes change.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [buffer])

  const apply = useCallback(
    (edit: (pkg: OfficePackage) => OfficePackage): Promise<void> => {
      queueRef.current = queueRef.current.then(async () => {
        const current = presentRef.current
        const pkg = edit(current.pkg)
        if (pkg === current.pkg) return
        const next = { pkg, slides: await parseRef.current(pkg, current) }
        presentRef.current = next
        history.set(next)
      })
      return queueRef.current
    },
    [history],
  )

  // Undo/redo join the same queue: an undo landing between an edit and its
  // re-parse would otherwise be silently undone by that in-flight edit.
  // They also move presentRef at once, so a Save (or edit) queued right
  // behind an undo works on the undone deck, not the one before it.
  const undo = useCallback((): void => {
    queueRef.current = queueRef.current.then(() => {
      presentRef.current = history.undo()
    })
  }, [history])

  const redo = useCallback((): void => {
    queueRef.current = queueRef.current.then(() => {
      presentRef.current = history.redo()
    })
  }, [history])

  const setDirty = useSetViewerDirty()
  const registerSave = useRegisterViewerSave()
  const registerSaveAs = useRegisterViewerSaveAs()
  useEffect(() => {
    setDirty(savedPkg !== null && history.present.pkg !== savedPkg)
  }, [history.present.pkg, savedPkg, setDirty])

  const write = useCallback(
    async (forceDialog: boolean): Promise<boolean> => {
      await queueRef.current
      const { pkg } = presentRef.current
      setSaveError(null)
      try {
        const bytes = await writeOfficePackage(pkg)
        const result = await window.electronAPI?.saveBinaryFile?.({
          content: bytes,
          suggestedName: baseName(savePath),
          filters: [{ name: saveFilter.name, extensions: [...saveFilter.extensions] }],
          ...(forceDialog ? {} : { existingPath: savePath }),
        })
        if (!result?.saved) {
          setSaveError(result?.error ?? 'Save was cancelled or unavailable.')
          return false
        }
        if (result.path) setSavePath(result.path)
        setSavedPkg(pkg)
        return true
      } catch (err: unknown) {
        setSaveError(err instanceof Error ? err.message : String(err))
        return false
      }
    },
    [saveFilter.extensions, saveFilter.name, savePath],
  )

  const save = useCallback(() => write(false), [write])
  const saveAs = useCallback(() => write(true), [write])

  useEffect(() => {
    registerSave(save)
    return () => registerSave(null)
  }, [registerSave, save])

  // Same registration, for the global Ctrl+Shift+S / App.tsx `saveFileAs()`
  // — previously unregistered, so that shortcut silently did nothing for a
  // slide deck (pptx/odp) even though the toolbar's own "Save As" button
  // worked.
  useEffect(() => {
    registerSaveAs(saveAs)
    return () => registerSaveAs(null)
  }, [registerSaveAs, saveAs])

  useViewerShortcuts(
    useCallback(
      (event: KeyboardEvent): boolean => {
        const target = event.target
        if (target instanceof HTMLElement && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName))) {
          return false
        }
        if (!(event.ctrlKey || event.metaKey)) return false
        const key = event.key.toLowerCase()
        if (key === 'z' && !event.shiftKey) {
          event.preventDefault()
          undo()
          return true
        }
        if (key === 'y' || (key === 'z' && event.shiftKey)) {
          event.preventDefault()
          redo()
          return true
        }
        return false
      },
      [redo, undo],
    ),
  )

  return {
    status,
    error,
    saveError,
    slides: history.present.slides,
    canUndo: history.canUndo,
    canRedo: history.canRedo,
    undo,
    redo,
    save,
    saveAs,
    apply,
    current: () => presentRef.current,
  }
}
