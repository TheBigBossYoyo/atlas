/**
 * The PDF editing session (PDF-EDIT-1): the edit model and its undo history,
 * the active tool and its style, form-field change tracking, and saving —
 * Save, Save As and the version-history capture — through the shell's
 * document-session contract.
 *
 * One session per loaded pdf.js document. It is keyed on the document object
 * itself and reset with the "store the prop in state" idiom (see
 * ViewerContext.tsx), so a new file, or the same file re-read from disk when
 * its tab is shown again, starts clean without a cascading effect.
 *
 * What a save writes is always rebuilt from the bytes the file was opened
 * with: form values through pdf.js, everything else through pdf-lib (see
 * `buildEditedPdf`). Nothing is ever applied to an already-edited copy, so
 * saving twice cannot apply an edit twice.
 */

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'

import { useViewerShortcuts } from '../../../hooks/useShortcutManager'
import type { ShortcutHandler } from '../../../hooks/shortcutManagerContext'
import { useTranslate } from '../../../i18n'
import { toFriendlyError } from '../../../utils/friendlyLibraryError'
import {
  useRegisterViewerCapture,
  useRegisterViewerSave,
  useRegisterViewerSaveAs,
  useReportSavedPath,
  useSetViewerDirty,
} from '../../shared/useViewerContext'
import type { PdfDocument } from '../types'
import {
  addAnnotation,
  createHistory,
  createInitialEditState,
  deletePage,
  isEditStateDirty,
  movePage,
  pushHistory,
  redoHistory,
  removeEdit,
  rotatePage,
  undoHistory,
  updateAnnotation,
  setTextEdit,
  type EditHistory,
} from './editState'
import { standardFontName } from './textLayout'
import type { FontStyle, PdfAnnotationDraft, PdfEditState, PdfEditTool, RgbColor, TextEdit } from './types'
import { StaleTextEditError, UnsupportedCharacterError } from './errors'
import type { PdfLib } from './writePdf'
import type { TextAnalysis } from './content/textAnalysis'
import { loadFallbackFonts } from './fallbackFonts'

/** Above this the periodic version capture is skipped: rebuilding the file would stall typing. */
const CAPTURE_SIZE_LIMIT_BYTES = 25 * 1024 * 1024

export const EDIT_COLORS: ReadonlyArray<{ readonly key: string; readonly color: RgbColor }> = [
  { key: 'black', color: { r: 0, g: 0, b: 0 } },
  { key: 'red', color: { r: 0.86, g: 0.15, b: 0.15 } },
  { key: 'blue', color: { r: 0.15, g: 0.35, b: 0.85 } },
  { key: 'green', color: { r: 0.1, g: 0.6, b: 0.25 } },
  { key: 'yellow', color: { r: 1, g: 0.85, b: 0.1 } },
  { key: 'pink', color: { r: 1, g: 0.45, b: 0.75 } },
]

const DEFAULT_TOOL_COLORS: Readonly<Record<'text' | 'highlight' | 'ink', RgbColor>> = {
  text: EDIT_COLORS[0].color,
  highlight: EDIT_COLORS[4].color,
  ink: EDIT_COLORS[1].color,
}

/** Measures and validates text in the 14 standard fonts, synchronously (UI needs both per keystroke). */
export type FontTools = {
  readonly widthOf: (font: FontStyle, text: string, size: number) => number
  /**
   * Whether every character of `text` can be drawn: by `font` itself, or —
   * once it has loaded — by the bundled sans fallback that text boxes switch
   * to for characters the standard fonts lack.
   */
  readonly canEncode: (font: FontStyle, text: string) => boolean
}

type Coverage = (codePoint: number) => boolean

function createFontTools(lib: PdfLib, fallbackCoverage: Coverage | null): FontTools {
  const embedders = new Map<string, ReturnType<PdfLib['StandardFontEmbedder']['for']>>()
  const embedderFor = (font: FontStyle) => {
    const name = standardFontName(font)
    let embedder = embedders.get(name)
    if (!embedder) {
      embedder = lib.StandardFontEmbedder.for(name as Parameters<PdfLib['StandardFontEmbedder']['for']>[0])
      embedders.set(name, embedder)
    }
    return embedder
  }
  return {
    widthOf: (font, text, size) => {
      try {
        return embedderFor(font).widthOfTextAtSize(text, size)
      } catch {
        return text.length * size * 0.5
      }
    },
    canEncode: (font, text) => {
      try {
        embedderFor(font).encodeText(text.replace(/[\r\n]/g, ''))
        return true
      } catch {
        if (!fallbackCoverage) return false
        return Array.from(text).every((ch) => /\s/.test(ch) || fallbackCoverage(ch.codePointAt(0) ?? 0))
      }
    },
  }
}

type Engine = typeof import('./engine')

let enginePromise: Promise<Engine> | null = null
/** The editing engine, loaded on first use (see engine.ts). */
function loadEngine(): Promise<Engine> {
  enginePromise ??= import('./engine').catch((error: unknown) => {
    enginePromise = null
    throw error
  })
  return enginePromise
}

let pdfLibPromise: Promise<PdfLib> | null = null
function loadPdfLib(): Promise<PdfLib> {
  pdfLibPromise ??= import('pdf-lib').catch((error: unknown) => {
    pdfLibPromise = null
    throw error
  })
  return pdfLibPromise
}

type EditorSession = {
  readonly doc: PdfDocument | null
  readonly history: EditHistory
  /** The edit state the file on disk matches. */
  readonly savedState: PdfEditState
  /** Bumped on every form-field change; non-zero means pdf.js must write the form. */
  readonly formRevision: number
  readonly savedFormRevision: number
}

function newSession(doc: PdfDocument | null): EditorSession {
  const initial = createInitialEditState(doc?.numPages ?? 0)
  return { doc, history: createHistory(initial), savedState: initial, formRevision: 0, savedFormRevision: 0 }
}

function suggestedFileName(path: string): string {
  const parts = path.replace(/\\/g, '/').split('/')
  return parts[parts.length - 1] || 'document.pdf'
}

export type DocumentProtection = {
  /** pdf-lib cannot rewrite an encrypted file, so only form filling is offered. */
  readonly encrypted: boolean
  /** Rewriting the file invalidates its digital signatures; the user is told before editing. */
  readonly signed: boolean
}

export type UsePdfEditorOptions = {
  readonly pdfDoc: PdfDocument | null
  /** The path this viewer was opened from; identifies the document to the shell. */
  readonly filePath: string
  /** The file's bytes as opened — every save starts from these. */
  readonly originalBytes: ArrayBuffer | null
}

export function usePdfEditor({ pdfDoc, filePath, originalBytes }: UsePdfEditorOptions) {
  const t = useTranslate()
  const setDirty = useSetViewerDirty()
  const registerSave = useRegisterViewerSave()
  const registerSaveAs = useRegisterViewerSaveAs()
  const registerCapture = useRegisterViewerCapture()
  const reportSavedPath = useReportSavedPath()

  const [session, setSession] = useState<EditorSession>(() => newSession(pdfDoc))
  if (session.doc !== pdfDoc) {
    setSession(newSession(pdfDoc))
  }
  // A layout effect, not a passive one: a save flushes the open text editor
  // and then reads this ref on the next task, by which point a layout effect
  // has certainly run and a passive one may not have.
  const sessionRef = useRef(session)
  useLayoutEffect(() => {
    sessionRef.current = session
  }, [session])

  /**
   * Commits for each page's open text editor (registered by PdfEditLayer). A
   * draft only reaches the edit state when its editor commits — on blur,
   * Ctrl+Enter or a click outside — and Ctrl+S does none of those, so a save
   * runs these first (the same job `flushHeaderFooterEdits` does for DOCX).
   */
  const flushersRef = useRef(new Set<() => void | Promise<void>>())
  const registerFlush = useCallback((flush: () => void | Promise<void>) => {
    flushersRef.current.add(flush)
    return () => {
      flushersRef.current.delete(flush)
    }
  }, [])
  const flushOpenEditors = useCallback(async () => {
    if (flushersRef.current.size === 0) return
    // A paragraph commit checks its characters first, so it can take a moment.
    await Promise.all(Array.from(flushersRef.current, (flush) => flush()))
    // Let the commit render, so `sessionRef` holds it.
    await new Promise((resolve) => setTimeout(resolve, 0))
  }, [])

  /**
   * Editing is locked while a Save As is in flight and, once it has moved the
   * document, until the viewer reloads from the new path: that reload starts a
   * fresh session, so an edit made in between would silently vanish. The lock
   * names the document it was taken on, so whatever the reload hands this
   * hook — a remount or just a new document — starts unlocked.
   */
  const [lockedDoc, setLockedDoc] = useState<PdfDocument | null>(null)
  const lockedDocRef = useRef<PdfDocument | null>(null)
  const isLocked = lockedDoc !== null && lockedDoc === session.doc
  const isLockedNow = useCallback(
    () => lockedDocRef.current !== null && lockedDocRef.current === sessionRef.current.doc,
    [],
  )
  const setLocked = useCallback((locked: boolean) => {
    const doc = locked ? sessionRef.current.doc : null
    lockedDocRef.current = doc
    setLockedDoc(doc)
  }, [])

  const [isEditing, setIsEditing] = useState(false)
  const [tool, setTool] = useState<PdfEditTool>('select')
  const [toolColors, setToolColors] = useState(DEFAULT_TOOL_COLORS)
  const [fontSize, setFontSize] = useState(12)
  const [inkWidth, setInkWidth] = useState(2)
  /** Typeface for text added with the edit-text tool. */
  const [textFamily, setTextFamily] = useState<FontStyle['family']>('helvetica')
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [pdfLib, setPdfLib] = useState<PdfLib | null>(null)
  const [engine, setEngine] = useState<Engine | null>(null)
  const [message, setMessage] = useState<string | null>(null)
  const [isSaving, setIsSaving] = useState(false)
  const [savePath, setSavePath] = useState(filePath)
  const [protection, setProtection] = useState<DocumentProtection>({ encrypted: false, signed: false })

  const edits = session.history.present
  const isDirty = isEditStateDirty(session.history, session.savedState) || session.formRevision !== session.savedFormRevision
  const [fallbackCoverage, setFallbackCoverage] = useState<{ readonly covers: Coverage } | null>(null)
  const fontTools = useMemo(
    () => (pdfLib ? createFontTools(pdfLib, fallbackCoverage?.covers ?? null) : null),
    [pdfLib, fallbackCoverage],
  )
  /** The original file's paragraphs, for the edit-text tool; available once pdf-lib has loaded. */
  const textAnalysis: TextAnalysis | null = useMemo(
    () =>
      pdfLib && engine && originalBytes
        ? engine.createTextAnalysis(pdfLib, new Uint8Array(originalBytes), loadFallbackFonts)
        : null,
    [engine, pdfLib, originalBytes],
  )

  // ---- Document facts ----------------------------------------------------------

  useEffect(() => {
    if (!pdfDoc || typeof pdfDoc.getMetadata !== 'function') return undefined
    let cancelled = false
    pdfDoc
      .getMetadata()
      .then(({ info }) => {
        if (cancelled) return
        const facts = (info ?? {}) as { EncryptFilterName?: unknown; IsSignaturesPresent?: unknown }
        setProtection({
          encrypted: facts.EncryptFilterName !== null && facts.EncryptFilterName !== undefined,
          signed: facts.IsSignaturesPresent === true,
        })
      })
      .catch(() => {
        // Unknown is treated as unprotected; a save that then fails says why.
      })
    return () => {
      cancelled = true
    }
  }, [pdfDoc])

  const ensurePdfLib = useCallback(async (): Promise<PdfLib> => {
    const [lib, loadedEngine] = await Promise.all([loadPdfLib(), loadEngine()])
    setPdfLib(lib)
    setEngine(loadedEngine)
    return lib
  }, [])

  const toggleEditing = useCallback(() => {
    if (isEditing) {
      setTool('select')
      setSelectedId(null)
    } else {
      void ensurePdfLib().catch(() => setMessage(t('pdf.edit.loadFailed')))
      // The sans fallback's character coverage, for text boxes in other scripts.
      void loadFallbackFonts()
        .then(async (source) => {
          const bytes = await source.load({ family: 'helvetica', bold: false, italic: false })
          const font = (source.fontkit as { create(b: Uint8Array): { hasGlyphForCodePoint(cp: number): boolean } }).create(bytes)
          setFallbackCoverage({ covers: (codePoint) => font.hasGlyphForCodePoint(codePoint) })
        })
        .catch(() => {
          // Without it, text boxes stay limited to the standard fonts' characters.
        })
    }
    setIsEditing(!isEditing)
  }, [ensurePdfLib, isEditing, t])

  // ---- Edits --------------------------------------------------------------------

  const commit = useCallback((change: (state: PdfEditState) => PdfEditState) => {
    if (isLockedNow()) return
    setSession((prev) => {
      const next = change(prev.history.present)
      return next === prev.history.present ? prev : { ...prev, history: pushHistory(prev.history, next) }
    })
  }, [isLockedNow])

  const undo = useCallback(() => {
    if (isLockedNow()) return
    setSelectedId(null)
    setSession((prev) => ({ ...prev, history: undoHistory(prev.history) }))
  }, [isLockedNow])

  const redo = useCallback(() => {
    if (isLockedNow()) return
    setSelectedId(null)
    setSession((prev) => ({ ...prev, history: redoHistory(prev.history) }))
  }, [isLockedNow])

  const addDraft = useCallback(
    (annotation: PdfAnnotationDraft) => {
      commit((state) => addAnnotation(state, annotation))
    },
    [commit],
  )

  const addDrafts = useCallback(
    (annotations: ReadonlyArray<PdfAnnotationDraft>) => {
      if (annotations.length === 0) return
      commit((state) => annotations.reduce(addAnnotation, state))
    },
    [commit],
  )

  const updateDraft = useCallback(
    (annotation: PdfAnnotationDraft) => {
      commit((state) => updateAnnotation(state, annotation))
    },
    [commit],
  )

  const removeSelected = useCallback(() => {
    if (selectedId === null) return
    commit((state) => removeEdit(state, selectedId))
    setSelectedId(null)
  }, [commit, selectedId])

  const removeById = useCallback(
    (id: string) => {
      commit((state) => removeEdit(state, id))
      setSelectedId((current) => (current === id ? null : current))
    },
    [commit],
  )

  /** Sets (or, when `unchanged`, drops) the edit of one paragraph or piece of new text. */
  const editText = useCallback(
    (edit: TextEdit, unchanged: boolean) => {
      commit((state) => setTextEdit(state, edit, unchanged))
    },
    [commit],
  )

  /**
   * The bytes the on-screen preview of text edits renders from: the original
   * file with only the text edits applied (annotations and page operations
   * are drawn by the viewer itself).
   */
  const buildTextPreview = useCallback(
    async (textEdits: ReadonlyArray<TextEdit>): Promise<Uint8Array | null> => {
      const doc = sessionRef.current.doc
      if (!doc || !originalBytes || textEdits.length === 0) return null
      const lib = await ensurePdfLib()
      const state: PdfEditState = { ...createInitialEditState(doc.numPages), textEdits }
      const { applyEditsToPdf } = await loadEngine()
      return applyEditsToPdf(lib, new Uint8Array(originalBytes), state, doc.numPages, await loadFallbackFonts())
    },
    [ensurePdfLib, originalBytes],
  )

  const rotatePageAt = useCallback(
    (displayIndex: number, delta: 90 | -90) => commit((state) => rotatePage(state, displayIndex, delta)),
    [commit],
  )
  const movePageAt = useCallback(
    (displayIndex: number, offset: number) => commit((state) => movePage(state, displayIndex, offset)),
    [commit],
  )
  const deletePageAt = useCallback(
    (displayIndex: number) => commit((state) => deletePage(state, displayIndex)),
    [commit],
  )

  const markFormChanged = useCallback(() => {
    setSession((prev) => ({ ...prev, formRevision: prev.formRevision + 1 }))
  }, [])

  const selectedAnnotation = edits.annotations.find((annotation) => annotation.id === selectedId) ?? null

  /**
   * Text boxes share the text colour; highlight and ink keep their own. With
   * the select tool and an annotation selected, a swatch recolours that
   * annotation instead of changing a tool's default.
   */
  const colorKey =
    tool === 'highlight' || tool === 'ink' ? tool : tool === 'select' && selectedAnnotation ? selectedAnnotation.kind : 'text'
  const setToolColor = useCallback(
    (color: RgbColor) => {
      if (tool === 'select' && selectedAnnotation) {
        commit((state) => updateAnnotation(state, { ...selectedAnnotation, color }))
      }
      setToolColors((colors) => ({ ...colors, [colorKey]: color }))
    },
    [colorKey, commit, selectedAnnotation, tool],
  )

  const setFontSizeAndSelection = useCallback(
    (size: number) => {
      if (selectedAnnotation?.kind === 'text') {
        commit((state) => updateAnnotation(state, { ...selectedAnnotation, fontSize: size }))
      }
      setFontSize(size)
    },
    [commit, selectedAnnotation],
  )

  // ---- Saving ---------------------------------------------------------------------

  const buildBytes = useCallback(
    async (snapshot: EditorSession): Promise<Uint8Array> => {
      const doc = snapshot.doc
      if (!doc || !originalBytes) throw new Error('The PDF has not finished loading.')
      const { buildEditedPdf } = await loadEngine()
      return buildEditedPdf({
        original: new Uint8Array(originalBytes),
        saveForms: snapshot.formRevision > 0 ? () => doc.saveDocument() : null,
        edits: snapshot.history.present,
        sourcePageCount: doc.numPages,
        loadPdfLib: ensurePdfLib,
        loadFallbackFonts,
      })
    },
    [ensurePdfLib, originalBytes],
  )

  const describeSaveError = useCallback(
    (error: unknown): string => {
      if (error instanceof UnsupportedCharacterError) return t('pdf.edit.unsupportedCharacters')
      if (error instanceof StaleTextEditError) return t('pdf.edit.staleTextEdit')
      if (error instanceof Error && error.name === 'EncryptedPDFError') return t('pdf.edit.encryptedSaveFailed')
      return toFriendlyError(error, t('pdf.edit.saveFailed')).message
    },
    [t],
  )

  const saveInternal = useCallback(
    async (forceDialog: boolean): Promise<boolean> => {
      if (!sessionRef.current.doc || isLockedNow()) return false
      await flushOpenEditors()
      const snapshot = sessionRef.current
      if (!snapshot.doc) return false
      setMessage(null)
      setIsSaving(true)
      if (forceDialog) setLocked(true)
      let movedToNewPath = false
      try {
        const bytes = await buildBytes(snapshot)
        const result = await window.electronAPI?.saveBinaryFile?.({
          content: bytes,
          suggestedName: suggestedFileName(savePath),
          ...(forceDialog ? {} : { existingPath: savePath }),
          filters: [{ name: t('pdf.edit.pdfFilter'), extensions: ['pdf'] }],
        })
        if (!result?.saved) {
          // A cancelled dialog is not an error; a failed write is.
          if (result?.error) setMessage(result.error)
          return false
        }
        if (result.path) {
          setSavePath(result.path)
          if (result.path !== savePath) {
            // The shell now reloads this document from its new path; stay
            // locked until that replaces this session.
            movedToNewPath = true
            reportSavedPath(filePath, result.path)
          }
        }
        // The snapshot that was written becomes the baseline — never the
        // current state, which may already hold edits made while writing.
        setSession((prev) =>
          prev.doc === snapshot.doc
            ? { ...prev, savedState: snapshot.history.present, savedFormRevision: snapshot.formRevision }
            : prev,
        )
        return true
      } catch (error) {
        console.error('Failed to save PDF', error)
        setMessage(describeSaveError(error))
        return false
      } finally {
        setIsSaving(false)
        if (forceDialog && !movedToNewPath) setLocked(false)
      }
    },
    [buildBytes, describeSaveError, filePath, flushOpenEditors, isLockedNow, reportSavedPath, savePath, setLocked, t],
  )

  /** Shows an error from outside a save (the preview rebuild) the way a save error shows. */
  const reportError = useCallback((error: unknown) => {
    console.error('PDF edit failed', error)
    setMessage(describeSaveError(error))
  }, [describeSaveError])

  const save = useCallback(() => saveInternal(false), [saveInternal])
  const saveAs = useCallback(() => saveInternal(true), [saveInternal])

  useEffect(() => {
    setDirty(isDirty)
  }, [isDirty, setDirty])

  useEffect(() => {
    registerSave(save)
    return () => registerSave(null)
  }, [registerSave, save])

  useEffect(() => {
    registerSaveAs(saveAs)
    return () => registerSaveAs(null)
  }, [registerSaveAs, saveAs])

  // VERSIONS-2 — a capture serialises without saving: no baseline move, no
  // dirty change. Read through a ref so the registered function never changes
  // identity (registering moves provider state; see useDocxSave.ts).
  const captureRef = useRef<() => Promise<Uint8Array | null>>(async () => null)
  useEffect(() => {
    captureRef.current = async () => {
      if (!originalBytes || originalBytes.byteLength > CAPTURE_SIZE_LIMIT_BYTES) return null
      return buildBytes(sessionRef.current)
    }
  }, [buildBytes, originalBytes])
  useEffect(() => {
    const capture = () => captureRef.current()
    registerCapture(capture)
    return () => registerCapture(null)
  }, [registerCapture])

  // ---- Shortcuts --------------------------------------------------------------------

  const shortcutHandler = useCallback<ShortcutHandler>(
    (event, ctx) => {
      if (ctx.inPlainField) return false
      const key = event.key.toLowerCase()
      if (event.ctrlKey || event.metaKey) {
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
      }
      if (!isEditing) return false
      if ((event.key === 'Delete' || event.key === 'Backspace') && selectedId !== null) {
        event.preventDefault()
        removeSelected()
        return true
      }
      if (event.key === 'Escape' && (selectedId !== null || tool !== 'select')) {
        event.preventDefault()
        setSelectedId(null)
        setTool('select')
        return true
      }
      return false
    },
    [isEditing, redo, removeSelected, selectedId, tool, undo],
  )
  useViewerShortcuts(shortcutHandler)

  return {
    edits,
    isDirty,
    canUndo: session.history.past.length > 0,
    canRedo: session.history.future.length > 0,
    isEditing,
    toggleEditing,
    tool,
    setTool,
    toolColors,
    /** The colour the swatches show as active: the selection's, or the tool's default. */
    activeColor: selectedAnnotation && tool === 'select' ? selectedAnnotation.color : toolColors[colorKey],
    setToolColor,
    fontSize: selectedAnnotation?.kind === 'text' ? selectedAnnotation.fontSize : fontSize,
    setFontSize: setFontSizeAndSelection,
    inkWidth,
    setInkWidth,
    textFamily,
    setTextFamily,
    selectedId,
    setSelectedId,
    fontTools,
    protection,
    message,
    setMessage,
    isSaving,
    /** True while a Save As is moving the document; the edit layer takes no input. */
    isLocked,
    registerFlush,
    undo,
    redo,
    addDraft,
    addDrafts,
    updateDraft,
    removeSelected,
    removeById,
    editText,
    reportError,
    textAnalysis,
    buildTextPreview,
    rotatePageAt,
    movePageAt,
    deletePageAt,
    markFormChanged,
    save,
    saveAs,
    buildBytes: useCallback(() => buildBytes(sessionRef.current), [buildBytes]),
  }
}

export type PdfEditor = ReturnType<typeof usePdfEditor>
