/**
 * Editing + save session for SpreadsheetViewer and CsvViewer (wave 3).
 *
 * Wraps `spreadsheetDocument.ts`'s pure operations with undo/redo history
 * (`useUndoableState`) and wires the result into the shared document-session
 * capability contract (`useSetViewerDirty`/`useRegisterViewerSave` — the
 * same contract DocxViewer's own P1.1 save/dirty plumbing uses), so a global
 * Ctrl+S/Save and the unsaved-changes-on-close confirmation both work here
 * exactly like they do for DOCX and markdown.
 *
 * `target` decides HOW the document is serialized on save:
 *  - `workbook`: `spreadsheetWrite.ts`'s `writeWorkbookBytes` +
 *    `window.electronAPI.saveBinaryFile` (SpreadsheetViewer's xlsx/ods
 *    family).
 *  - `delimited`: `documentToDelimitedText` + `window.electronAPI.saveFile`
 *    (CsvViewer's csv/tsv).
 *
 * "Save As with format choice" (the plan's own wording) is a `target`
 * override passed to `handleSaveAs` — the caller's own toolbar UI is
 * responsible for asking the user which format, this hook just knows how to
 * write whichever one it's given.
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import type * as XLSXNamespace from 'xlsx'

import {
  addSheet as addSheetOp,
  deleteColumnAt as deleteColumnOp,
  deleteRowAt as deleteRowOp,
  deleteSheet as deleteSheetOp,
  insertColumnAt as insertColumnOp,
  insertRowAt as insertRowOp,
  pasteRange as pasteRangeOp,
  renameSheet as renameSheetOp,
  setCellValue as setCellValueOp,
  type SpreadsheetDocument,
} from './spreadsheetDocument'
import { documentToDelimitedText, writeWorkbookBytesWithTables } from './spreadsheetWrite'
import { writeWorkbookThroughOriginal } from './xlsxPassthrough'
import { useUndoableState } from './useUndoableState'
import {
  useRegisterViewerSave,
  useRegisterViewerSaveAs,
  useReportSavedPath,
  useSetViewerDirty,
} from '../shared/useViewerContext'
import { useViewerShortcuts } from '../../hooks/useShortcutManager'

export type SpreadsheetSaveTarget =
  | {
      readonly kind: 'workbook'
      readonly bookType: XLSXNamespace.BookType
      readonly extension: string
      readonly filterName: string
    }
  | { readonly kind: 'delimited'; readonly delimiter: string; readonly extension: string; readonly filterName: string }

const EMPTY_DOCUMENT: SpreadsheetDocument = { sheets: [] }

function baseName(filePath: string): string {
  const normalized = filePath.replace(/\\/g, '/')
  const parts = normalized.split('/')
  return parts[parts.length - 1] || 'spreadsheet'
}

function withExtension(name: string, extension: string): string {
  const dot = name.lastIndexOf('.')
  const stem = dot > 0 ? name.slice(0, dot) : name
  return `${stem}.${extension}`
}

/** Formats whose package Atlas can patch in place, keeping styles and everything it does not model. */
const PASSTHROUGH_BOOK_TYPES: ReadonlySet<string> = new Set(['xlsx', 'xlsm'])

async function writeToDisk(
  doc: SpreadsheetDocument,
  target: SpreadsheetSaveTarget,
  suggestedName: string,
  existingPath: string | undefined,
  originalBuffer: ArrayBuffer | null,
): Promise<{ readonly saved: boolean; readonly path?: string; readonly error?: string }> {
  const filters = [{ name: target.filterName, extensions: [target.extension] }]

  if (target.kind === 'workbook') {
    // USR-17 — rewrite the file the user opened (keeping every style, chart
    // and filter Atlas does not model) whenever that is possible; the
    // fresh-workbook writer is the fallback for other formats and for
    // structural changes it cannot express.
    const throughOriginal =
      originalBuffer && PASSTHROUGH_BOOK_TYPES.has(target.bookType)
        ? await writeWorkbookThroughOriginal(originalBuffer, doc)
        : null
    const bytes = throughOriginal ?? (await writeWorkbookBytesWithTables(doc, target.bookType))
    const result = await window.electronAPI?.saveBinaryFile?.({
      content: bytes,
      suggestedName,
      filters,
      ...(existingPath ? { existingPath } : {}),
    })
    return { saved: result?.saved ?? false, path: result?.path, error: result?.error }
  }

  const text = documentToDelimitedText(doc, target.delimiter)
  const result = await window.electronAPI?.saveFile?.({
    content: text,
    suggestedName,
    filters,
    ...(existingPath ? { existingPath } : {}),
  })
  return { saved: result?.saved ?? false, path: result?.path, error: result?.error }
}

export type UseSpreadsheetEditorResult = {
  readonly document: SpreadsheetDocument
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly undo: () => void
  readonly redo: () => void
  readonly setCellValue: (sheetIndex: number, row: number, col: number, rawInput: string) => void
  readonly insertRowAt: (sheetIndex: number, atIndex: number) => void
  readonly deleteRowAt: (sheetIndex: number, atIndex: number) => void
  readonly insertColumnAt: (sheetIndex: number, atIndex: number) => void
  readonly deleteColumnAt: (sheetIndex: number, atIndex: number) => void
  readonly pasteRange: (sheetIndex: number, row: number, col: number, values: ReadonlyArray<ReadonlyArray<string>>) => void
  readonly addSheet: (name?: string) => void
  readonly renameSheet: (sheetIndex: number, name: string) => void
  readonly deleteSheet: (sheetIndex: number) => void
  readonly saveError: string | null
  readonly handleSave: () => Promise<boolean>
  readonly handleSaveAs: (overrideTarget?: SpreadsheetSaveTarget) => Promise<boolean>
}

export function useSpreadsheetEditor(
  initialDocument: SpreadsheetDocument | null,
  filePath: string,
  target: SpreadsheetSaveTarget,
  // .xls/.xlsb/.fods: "view + save-as-xlsx only" (plan scope) — SheetJS CAN
  // technically round-trip these in place (verified directly against the
  // library), but the plan deliberately doesn't offer that: pass `false` so
  // the very first Save always goes through the format-choice dialog
  // (effectively a forced Save As) instead of silently overwriting the
  // original legacy/flat file with re-encoded bytes. Once the user has
  // saved once, the *new* path they chose becomes the seed for subsequent
  // plain saves, same as any other format.
  seedExistingPath: boolean = true,
  /** The bytes the workbook was opened from, for save-through-original (USR-17). */
  originalBuffer: ArrayBuffer | null = null,
): UseSpreadsheetEditorResult {
  const history = useUndoableState<SpreadsheetDocument>(EMPTY_DOCUMENT)
  const hydratedRef = useRef(false)
  // Seeded from the loaded file's own path (mirrors DocxViewer's identical
  // `useState(file.path)`) so the very first Save silently overwrites the
  // file in place instead of behaving like an unwanted Save As — except for
  // the legacy-format case above, where that's the opposite of what's wanted.
  const [savePath, setSavePath] = useState<string | undefined>(seedExistingPath ? filePath : undefined)
  const [saveError, setSaveError] = useState<string | null>(null)
  // The document reference at the last successful save (or at load) — a
  // plain reference compare against the current `history.present` is
  // "has this document changed since load/save", mirroring DocxViewer's own
  // documentModel/lastSavedDocument dirty-tracking pattern exactly.
  const [lastSavedDocument, setLastSavedDocument] = useState<SpreadsheetDocument>(EMPTY_DOCUMENT)

  useEffect(() => {
    if (initialDocument && !hydratedRef.current) {
      hydratedRef.current = true
      history.reset(initialDocument)
      setLastSavedDocument(initialDocument)
    }
    // history.reset is stable (see useUndoableState); only re-run when a
    // real initialDocument shows up.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialDocument])

  const setDirty = useSetViewerDirty()
  const registerSave = useRegisterViewerSave()
  const registerSaveAs = useRegisterViewerSaveAs()
  const reportSavedPath = useReportSavedPath()

  useEffect(() => {
    setDirty(history.present !== lastSavedDocument)
  }, [history.present, lastSavedDocument, setDirty])

  const mutate = useCallback(
    (updater: (doc: SpreadsheetDocument) => SpreadsheetDocument) => {
      const current = history.present
      const next = updater(current)
      // Every guarded no-op in spreadsheetDocument.ts (deleting a sheet's
      // last remaining row/column, renaming to an already-used name,
      // deleting the only sheet, an out-of-bounds cell edit, ...)
      // deliberately returns the SAME document reference, unchanged, rather
      // than a structurally-equal copy — specifically so callers can detect
      // "nothing happened" this cheaply. Without this check, a blocked
      // action (e.g. clicking "Delete row" on a one-row sheet, which the
      // toolbar only disables when nothing is selected, not when deleting
      // would be a no-op) still pushed a duplicate entry onto the undo
      // stack, breaking useUndoableState's own "one call per discrete,
      // already-committed edit" contract: the very next Undo would silently
      // restore the exact same state instead of reverting the last REAL
      // edit, costing the user an extra, invisible Undo press.
      if (next !== current) history.set(next)
    },
    [history],
  )

  const setCellValue = useCallback(
    (sheetIndex: number, row: number, col: number, rawInput: string) =>
      mutate((doc) => setCellValueOp(doc, sheetIndex, row, col, rawInput)),
    [mutate],
  )
  const insertRowAt = useCallback(
    (sheetIndex: number, atIndex: number) => mutate((doc) => insertRowOp(doc, sheetIndex, atIndex)),
    [mutate],
  )
  const deleteRowAt = useCallback(
    (sheetIndex: number, atIndex: number) => mutate((doc) => deleteRowOp(doc, sheetIndex, atIndex)),
    [mutate],
  )
  const insertColumnAt = useCallback(
    (sheetIndex: number, atIndex: number) => mutate((doc) => insertColumnOp(doc, sheetIndex, atIndex)),
    [mutate],
  )
  const deleteColumnAt = useCallback(
    (sheetIndex: number, atIndex: number) => mutate((doc) => deleteColumnOp(doc, sheetIndex, atIndex)),
    [mutate],
  )
  const pasteRange = useCallback(
    (sheetIndex: number, row: number, col: number, values: ReadonlyArray<ReadonlyArray<string>>) =>
      mutate((doc) => pasteRangeOp(doc, sheetIndex, row, col, values)),
    [mutate],
  )
  const addSheet = useCallback((name?: string) => mutate((doc) => addSheetOp(doc, name)), [mutate])
  const renameSheet = useCallback(
    (sheetIndex: number, name: string) => mutate((doc) => renameSheetOp(doc, sheetIndex, name)),
    [mutate],
  )
  const deleteSheet = useCallback(
    (sheetIndex: number) => mutate((doc) => deleteSheetOp(doc, sheetIndex)),
    [mutate],
  )

  const handleSaveWith = useCallback(
    async (saveTarget: SpreadsheetSaveTarget, forceDialog: boolean): Promise<boolean> => {
      setSaveError(null)
      const docAtSaveStart = history.present
      try {
        const suggestedName = withExtension(baseName(filePath), saveTarget.extension)
        const result = await writeToDisk(
          docAtSaveStart,
          saveTarget,
          suggestedName,
          forceDialog ? undefined : savePath,
          originalBuffer,
        )

        if (!result.saved) {
          setSaveError(result.error ?? 'Save was cancelled or unavailable.')
          return false
        }

        if (result.path) {
          setSavePath(result.path)
          // Save As: move this document's tab to the file it was written to.
          // `filePath` (this hook instance's own identity — the caller
          // remounts on path change, so it never changes across this
          // instance's lifetime) is passed through as the path this save
          // started from, so the shell can route the update to the RIGHT tab
          // even if the user has since switched away (SAVE-1). Unlike
          // `savePath`, `filePath` is never `undefined` — `savePath` starts
          // out that way for legacy formats (`seedExistingPath: false`)
          // until their first save ever succeeds.
          if (result.path !== savePath) reportSavedPath(filePath, result.path)
        }
        // See DocxViewer's identical comment: deliberately does NOT also
        // call setDirty(false) here — the dirty-tracking effect above
        // recomputes from whatever the *current* history.present is once
        // this state update lands, so an edit made while the save's awaits
        // were in flight is never silently marked clean.
        setLastSavedDocument(docAtSaveStart)
        return true
      } catch (error) {
        setSaveError(error instanceof Error ? error.message : String(error))
        return false
      }
    },
    [filePath, history.present, savePath, originalBuffer, reportSavedPath],
  )

  const handleSave = useCallback(() => handleSaveWith(target, false), [handleSaveWith, target])
  const handleSaveAs = useCallback(
    (overrideTarget?: SpreadsheetSaveTarget) => handleSaveWith(overrideTarget ?? target, true),
    [handleSaveWith, target],
  )

  useEffect(() => {
    registerSave(handleSave)
    return () => registerSave(null)
  }, [registerSave, handleSave])

  // Same registration, for the global Ctrl+Shift+S / App.tsx `saveFileAs()`
  // — previously unregistered, so that shortcut silently did nothing for a
  // spreadsheet even though the toolbar's own "Save As…" button worked.
  // `handleSaveAs`'s optional override argument is never passed here, so it
  // always saves to the format the toolbar's own "Save As" button defaults
  // to (`target`/`defaultSaveTarget`), same as clicking that button plainly.
  useEffect(() => {
    registerSaveAs(handleSaveAs)
    return () => registerSaveAs(null)
  }, [registerSaveAs, handleSaveAs])

  // Ctrl+Z/Ctrl+Y (and Ctrl+Shift+Z as the common redo alternative) at the
  // active-viewer shortcut precedence tier (see DocxViewer's identical use
  // of `useViewerShortcuts` for its own undo/redo) — glide-data-grid's
  // canvas grid has no native contentEditable undo of its own to fall back
  // on, so without this, undo/redo only worked by clicking the toolbar
  // buttons despite being the very first thing a user reaches for. Bails out
  // for a real `<input>`/`<textarea>` target (the sheet-rename box, the
  // "Search rows..." field, FrozenRowsStrip's per-cell inputs) so their own
  // native text-editing undo keeps working instead of being hijacked.
  useViewerShortcuts(
    useCallback(
      (event: KeyboardEvent): boolean => {
        if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement) {
          return false
        }
        const ctrl = event.ctrlKey || event.metaKey
        if (!ctrl) return false
        const key = event.key.toLowerCase()
        if (key === 'z' && !event.shiftKey) {
          event.preventDefault()
          history.undo()
          return true
        }
        if (key === 'y' || (key === 'z' && event.shiftKey)) {
          event.preventDefault()
          history.redo()
          return true
        }
        return false
      },
      [history],
    ),
  )

  return {
    document: history.present,
    canUndo: history.canUndo,
    canRedo: history.canRedo,
    undo: history.undo,
    redo: history.redo,
    setCellValue,
    insertRowAt,
    deleteRowAt,
    insertColumnAt,
    deleteColumnAt,
    pasteRange,
    addSheet,
    renameSheet,
    deleteSheet,
    saveError,
    handleSave,
    handleSaveAs,
  }
}
