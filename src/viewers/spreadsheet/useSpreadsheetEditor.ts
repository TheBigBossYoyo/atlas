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
import { documentToDelimitedText, writeWorkbookBytes } from './spreadsheetWrite'
import { useUndoableState } from './useUndoableState'
import { useRegisterViewerSave, useSetViewerDirty } from '../shared/useViewerContext'

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

async function writeToDisk(
  doc: SpreadsheetDocument,
  target: SpreadsheetSaveTarget,
  suggestedName: string,
  existingPath: string | undefined,
): Promise<{ readonly saved: boolean; readonly path?: string; readonly error?: string }> {
  const filters = [{ name: target.filterName, extensions: [target.extension] }]

  if (target.kind === 'workbook') {
    const bytes = writeWorkbookBytes(doc, target.bookType)
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

  useEffect(() => {
    setDirty(history.present !== lastSavedDocument)
  }, [history.present, lastSavedDocument, setDirty])

  const mutate = useCallback(
    (updater: (doc: SpreadsheetDocument) => SpreadsheetDocument) => {
      history.set(updater(history.present))
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
        )

        if (!result.saved) {
          setSaveError(result.error ?? 'Save was cancelled or unavailable.')
          return false
        }

        if (result.path) setSavePath(result.path)
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
    [filePath, history.present, savePath],
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
