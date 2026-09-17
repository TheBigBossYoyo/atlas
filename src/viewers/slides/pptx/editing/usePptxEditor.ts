/**
 * USR-16 — PPTX editing session: loads the package, applies edits with
 * undo/redo, re-parses only the slides whose parts changed, and plugs into
 * the shared save/dirty contract (Ctrl+S, unsaved-changes prompt) like the
 * DOCX and spreadsheet editors.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'

import type { SlideData } from '../../../shared/SlideDeck.types'
import { useRegisterViewerSave, useSetViewerDirty } from '../../../shared/useViewerContext'
import { useViewerShortcuts } from '../../../../hooks/useShortcutManager'
import { useUndoableState } from '../../../spreadsheet/useUndoableState'
import { parseOneSlide, resolveSlidePaths } from '../parser'
import { REL_TYPE, readRels, relsPathFor } from './opcXml'
import * as edits from './pptxEdits'
import { loadPptxPackage, packageArchive, readPart, writePptxPackage, type PptxPackage } from './pptxPackage'
import * as slideOps from './pptxSlideOps'

type DeckState = { readonly pkg: PptxPackage; readonly slides: ReadonlyArray<SlideData> }

const EMPTY_STATE: DeckState = { pkg: { parts: new Map() }, slides: [] }
const NEW_TEXT_BOX_TEXT = 'Text'

/** The parts a parsed slide depends on; unchanged references mean the cached parse is still valid. */
function slideInputs(pkg: PptxPackage, path: string): ReadonlyArray<unknown> {
  const relsPath = relsPathFor(path)
  const notesPath = readRels(readPart(pkg, relsPath), path).find((rel) => rel.type === REL_TYPE.notesSlide)?.target
  return [pkg.parts.get(path), pkg.parts.get(relsPath), notesPath ? pkg.parts.get(notesPath) : undefined]
}

async function parseDeck(pkg: PptxPackage, previous: DeckState | null): Promise<ReadonlyArray<SlideData>> {
  const archive = packageArchive(pkg)
  const signal = { cancelled: false }
  const { slidePaths, size } = await resolveSlidePaths(archive, signal)
  return Promise.all(
    slidePaths.map(async (path, index) => {
      const cached = previous?.slides.find((slide) => slide.partPath === path)
      const unchanged =
        cached !== undefined &&
        previous !== null &&
        slideInputs(pkg, path).every((part, i) => part === slideInputs(previous.pkg, path)[i])
      if (unchanged && cached) return { ...cached, index, id: `slide-${index + 1}` }
      return parseOneSlide(archive, path, index, size, signal)
    }),
  )
}

function baseName(filePath: string): string {
  return filePath.replace(/\\/g, '/').split('/').pop() || 'presentation.pptx'
}

export type SlideBox = edits.ShapeBox

export function usePptxEditor(buffer: ArrayBuffer | null, filePath: string) {
  const history = useUndoableState<DeckState>(EMPTY_STATE)
  const [status, setStatus] = useState<'loading' | 'ready' | 'error'>('loading')
  const [error, setError] = useState<string | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [savedPkg, setSavedPkg] = useState<PptxPackage | null>(null)
  const [savePath, setSavePath] = useState<string>(filePath)
  const presentRef = useRef(history.present)
  useLayoutEffect(() => {
    presentRef.current = history.present
  }, [history.present])
  const queueRef = useRef<Promise<void>>(Promise.resolve())

  useEffect(() => {
    let cancelled = false
    void (async () => {
      setStatus('loading')
      if (!buffer) return
      try {
        const pkg = await loadPptxPackage(buffer)
        const slides = await parseDeck(pkg, null)
        if (cancelled) return
        history.reset({ pkg, slides })
        setSavedPkg(pkg)
        setError(slides.length > 0 ? null : 'No slides found in PPTX.')
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

  /** Serializes edits: each one sees the result of the previous one, even while its re-parse is in flight. */
  const apply = useCallback(
    (edit: (pkg: PptxPackage) => PptxPackage): Promise<void> => {
      queueRef.current = queueRef.current.then(async () => {
        const current = presentRef.current
        const pkg = edit(current.pkg)
        if (pkg === current.pkg) return
        const next = { pkg, slides: await parseDeck(pkg, current) }
        presentRef.current = next
        history.set(next)
      })
      return queueRef.current
    },
    [history],
  )

  const pathOf = (index: number): string | null => presentRef.current.slides[index]?.partPath ?? null
  const onSlide = (index: number, edit: (pkg: PptxPackage, path: string) => PptxPackage) =>
    apply((pkg) => {
      const path = pathOf(index)
      return path ? edit(pkg, path) : pkg
    })

  // Undo/redo run through the SAME queue as edits: an edit job reads the
  // present state when it starts and writes its result when its re-parse
  // finishes, so an undo landing in between would be silently undone again by
  // that in-flight job (and its redo entry dropped by the next history.set).
  const undo = useCallback((): void => {
    queueRef.current = queueRef.current.then(() => history.undo())
  }, [history])

  const redo = useCallback((): void => {
    queueRef.current = queueRef.current.then(() => history.redo())
  }, [history])

  const setDirty = useSetViewerDirty()
  const registerSave = useRegisterViewerSave()
  useEffect(() => {
    setDirty(savedPkg !== null && history.present.pkg !== savedPkg)
  }, [history.present.pkg, savedPkg, setDirty])

  const save = useCallback(
    async (forceDialog: boolean): Promise<boolean> => {
      await queueRef.current
      const { pkg } = presentRef.current
      setSaveError(null)
      try {
        const bytes = await writePptxPackage(pkg)
        const result = await window.electronAPI?.saveBinaryFile?.({
          content: bytes,
          suggestedName: baseName(savePath),
          filters: [{ name: 'PowerPoint Presentation', extensions: ['pptx'] }],
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
    [savePath],
  )

  const handleSave = useCallback(() => save(false), [save])
  useEffect(() => {
    registerSave(handleSave)
    return () => registerSave(null)
  }, [registerSave, handleSave])

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
      [undo, redo],
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
    save: handleSave,
    saveAs: useCallback(() => save(true), [save]),
    canEditNotes: (index: number): boolean => {
      const path = history.present.slides[index]?.partPath
      return path ? edits.canEditNotes(history.present.pkg, path) : false
    },
    setShapeText: (index: number, sourceId: string, text: string) =>
      onSlide(index, (pkg, path) => edits.setShapeText(pkg, path, sourceId, text)),
    setShapeBox: (index: number, sourceId: string, box: SlideBox) =>
      onSlide(index, (pkg, path) => edits.setShapeBox(pkg, path, sourceId, box)),
    deleteShape: (index: number, sourceId: string) =>
      onSlide(index, (pkg, path) => edits.deleteShape(pkg, path, sourceId)),
    setNotes: (index: number, text: string) => onSlide(index, (pkg, path) => edits.setSlideNotes(pkg, path, text)),
    /** Resolves the new text box's shape id once it is on the slide. */
    insertTextBox: async (index: number, box: SlideBox): Promise<string | null> => {
      let sourceId: string | null = null
      await onSlide(index, (pkg, path) => {
        const result = edits.insertTextBox(pkg, path, box, NEW_TEXT_BOX_TEXT)
        sourceId = result.sourceId
        return result.pkg
      })
      return sourceId
    },
    /** Resolves the new slide's index. */
    addSlide: async (afterIndex: number): Promise<number> => {
      let index = afterIndex
      await apply((pkg) => {
        const result = slideOps.addSlide(pkg, afterIndex)
        index = result.index
        return result.pkg
      })
      return index
    },
    duplicateSlide: (index: number) => apply((pkg) => slideOps.duplicateSlide(pkg, index)),
    deleteSlide: (index: number) => apply((pkg) => slideOps.deleteSlide(pkg, index)),
    moveSlide: (from: number, to: number) => apply((pkg) => slideOps.moveSlide(pkg, from, to)),
  }
}

export type PptxEditor = ReturnType<typeof usePptxEditor>
