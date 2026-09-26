/**
 * Saving a DOCX: the write itself, the path it goes to, the dirty state the shell
 * reads, and the one-per-document round-trip fidelity notice.
 *
 * Extracted from `src/viewers/DocxViewer.tsx`, the last large logic block in that
 * file. It takes more inputs than the other hooks pulled out of there, but they
 * are almost all read-only (the bundle, the current revision, the selection) —
 * the only thing it writes besides its own state is `setSaveError`, which stayed
 * in the viewer because it is that component's general error banner (image
 * insert, hyperlink, lists and comments all use it), not a save-specific channel.
 *
 * It reaches the shared document-session contract itself rather than having it
 * threaded in: `useReportSavedPath`, `useSetViewerDirty`, `useRegisterViewerSave`
 * and `useRegisterViewerSaveAs` are context consumers, so consuming them here
 * keeps four parameters off the signature and puts the registration next to the
 * handlers being registered.
 *
 * Three pieces of reasoning worth keeping in view, all of which cost real bugs
 * to learn:
 *   - **Dirty is an integer compare, not a reference compare.** `documentRevision
 *     !== lastSavedRevision`, because undo/redo rebuild the document from a
 *     stored inverse and never hand back the same object twice even when the
 *     content is identical to what is on disk. The revision counter is "how many
 *     forward edits from the last save/load", and undo winds it back.
 *   - **The revision written is captured, the dirty flag is not.** `setDirty` is
 *     deliberately left to an effect that recomputes from the CURRENT render, so
 *     an edit landing while the write is still in flight is not lost; calling
 *     `setDirty(false)` inside the save would clear it against a stale snapshot.
 *   - **Header/footer fields are flushed first.** Ctrl+S never blurs the focused
 *     field, and the field commits on blur, so without `flushHeaderFooterEdits()`
 *     the save writes the pre-edit text.
 */
import { useCallback, useEffect, useRef, useState, type RefObject } from 'react'

import {
  categorizeLossySaveWarnings,
  detectLossySaveWarnings,
  saveDocx,
  type DocxBundle,
  type LossySaveWarningCategory,
} from '../index'
import { friendlyDocxErrorMessage, type Range } from '../editor'
import type { Document as DocxDocument } from '../model'
import { useTranslate, type TranslateFn } from '../../i18n'
import {
  useRegisterViewerSave,
  useRegisterViewerSaveAs,
  useReportSavedPath,
  useSetViewerDirty,
} from '../../viewers/shared/useViewerContext'
import { syncSelectionToDom } from './selectionDom'

/** The file name a save dialog should suggest, from a full path. */
function getSuggestedFileName(path: string): string {
  const normalized = path.replace(/\\/g, '/')
  const parts = normalized.split('/')
  return parts[parts.length - 1] || 'document.docx'
}

export type DocxSaveOptions = {
  readonly bundle: DocxBundle
  /** History's revision counter for the current render — one half of the dirty compare. */
  readonly documentRevision: number
  readonly range: Range | null
  /**
   * The path this viewer instance was opened from. ViewerRouter remounts on path
   * change, so it never changes for the life of this hook — which is exactly why
   * it, and not the mutable `savePath`, identifies the document to the shell.
   */
  readonly filePath: string
  /** Commits any still-pending header/footer field edit and returns the document to write. */
  readonly flushHeaderFooterEdits: () => DocxDocument
  /** Reads History's revision synchronously — a ref read, so it is already current the instant a flush returns. */
  readonly getRevision: () => number
  readonly editorRootRef: RefObject<HTMLDivElement | null>
  readonly documentModelRef: RefObject<DocxDocument>
  /** The viewer's shared error banner. */
  readonly setSaveError: (message: string | null) => void
}

export type DocxSave = {
  /** Where the document currently lives; moves on a successful Save As. */
  readonly savePath: string
  readonly handleSave: () => Promise<boolean>
  readonly handleSaveAs: () => Promise<boolean>
  readonly fidelityWarningMessage: string | null
  readonly dismissFidelityWarning: () => void
}

/**
 * Round-trip fidelity audit, DXS round 2 follow-up — plain-language,
 * non-technical copy for what `detectLossySaveWarnings` found (see
 * `categorizeLossySaveWarnings` in `docx/fidelity/lossySaveWarnings.ts`).
 * Built here rather than in that module because it has (and should keep)
 * no dependency on the i18n layer. Names the feature the user actually
 * recognizes ("content controls", "a text box's fallback drawing"), never
 * the underlying XML element — see `handleSaveInternal`'s doc comment for
 * when/how often this is shown. `null` when there's nothing to say.
 */
function buildFidelitySaveWarningMessage(
  categories: ReadonlySet<LossySaveWarningCategory>,
  t: TranslateFn,
): string | null {
  const sentences: string[] = []
  if (categories.has('content-control')) {
    sentences.push(t('docx.viewer.fidelityWarningContentControls'))
  }
  if (categories.has('shape-fallback')) {
    sentences.push(t('docx.viewer.fidelityWarningShapeFallback'))
  }
  if (categories.has('other')) {
    sentences.push(t('docx.viewer.fidelityWarningOther'))
  }
  if (sentences.length === 0) {
    return null
  }
  return `${t('docx.viewer.fidelityWarningIntro')} ${sentences.join(' ')}`
}

export function useDocxSave({
  bundle,
  documentRevision,
  range,
  filePath,
  flushHeaderFooterEdits,
  getRevision,
  editorRootRef,
  documentModelRef,
  setSaveError,
}: DocxSaveOptions): DocxSave {
  const t = useTranslate()
  const reportSavedPath = useReportSavedPath()
  const setDirty = useSetViewerDirty()
  const registerSave = useRegisterViewerSave()
  const registerSaveAs = useRegisterViewerSaveAs()

  const [savePath, setSavePath] = useState(filePath)
  const [lastSavedRevision, setLastSavedRevision] = useState(0)
  // Once per document, not once per save: a document that still trips the same
  // warning on every subsequent save should not keep re-announcing it. Not reset
  // when `savePath` changes via Save As — that is the same document.
  const warnedAboutFidelityRef = useRef(false)
  const [fidelityWarningMessage, setFidelityWarningMessage] = useState<string | null>(null)

  const dismissFidelityWarning = useCallback(() => {
    setFidelityWarningMessage(null)
  }, [])

  const handleSaveInternal = useCallback(
    async (options?: { readonly forceDialog?: boolean }): Promise<boolean> => {
      setSaveError(null)

      // D29 follow-up — commit any header/footer field the user is still
      // typing in (Ctrl+S never blurs it) before reading the document to
      // save; see `flushHeaderFooterEdits`'s own doc comment above.
      const documentToSave = flushHeaderFooterEdits()
      // DIRTY-1 — read synchronously, right alongside `documentToSave`, not
      // from the `documentRevision` state variable: `flushHeaderFooterEdits`
      // may just have pushed one more command to `historyRef.current`
      // (a pending field edit) via a `commitState` call whose `setState`s
      // are still only scheduled, not yet reflected in this render's
      // closure. `historyRef.current` itself, being a plain mutable ref, is
      // already up to date the instant that call returns.
      const revisionToSave = getRevision()

      try {
        const nextBytes = await saveDocx({
          ...bundle,
          document: documentToSave,
        })

        const result = await window.electronAPI?.saveBinaryFile?.({
          content: nextBytes,
          suggestedName: getSuggestedFileName(savePath),
          // DXE-24 — omitting `existingPath` makes the main process show the
          // save dialog instead of silently overwriting `savePath`, which is
          // exactly "Save As".
          ...(options?.forceDialog ? {} : { existingPath: savePath }),
          filters: [{ name: t('docx.viewer.saveAsWordFilter'), extensions: ['docx'] }],
        })

        if (result?.saved && result.path) {
          setSavePath(result.path)
          // Save As: tell the shell where this document now lives, or its tab
          // keeps pointing at the file it was opened from. `file.path` (this
          // viewer instance's own identity — ViewerRouter remounts on path
          // change, so it never changes across this component's lifetime) is
          // passed through as the path this save started from, so the shell
          // can route the update to the RIGHT tab even if the user has since
          // switched away (SAVE-1) — `savePath` itself isn't safe for that:
          // it's already been reassigned above by the time this fires for a
          // second save, and (in the spreadsheet editor's legacy-format case)
          // can be `undefined` before the first save ever completes.
          if (result.path !== savePath) reportSavedPath(filePath, result.path)
        }

        if (!result?.saved) {
          setSaveError(result?.error ?? t('docx.viewer.saveCancelled'))
          return false
        }

        // DOCX-3 — Save As (`options?.forceDialog`) shows a native OS save
        // dialog via the main process; that dialog takes OS-level focus away
        // from the whole Electron window, and nothing gives it back to
        // `editorRootRef` once it closes (Chromium doesn't restore focus to
        // whatever had it before an OS dialog interrupted the page).
        // `document.activeElement` is left as `<body>`, and the model
        // `range`/native `Selection` are never resynced because the
        // selection-sync effect only re-runs when `pages` or `range` change
        // — neither does just from saving — so the user has to click back
        // into the document before typing does anything again. Plain Save
        // never shows a dialog (confirmed against the real app: no observed
        // focus loss), so this is scoped to the Save As path specifically.
        //
        // This is necessary but NOT sufficient end-to-end: confirmed against
        // the real app that once `reportSavedPath` (above) updates the
        // active tab's path, `src/components/ViewerRouter.tsx` — which keys
        // its `ViewerErrorBoundary` (and so this whole component) by
        // `file.path` — remounts a BRAND NEW `DocxEditor` a render or two
        // later, discarding the focus this restores along with it. Fixing
        // that fully needs a change in ViewerRouter.tsx (outside this
        // component's ownership), described in this task's final report;
        // this restoration still stands on its own for any Save As that
        // doesn't trigger that remount, and is the half of the fix that
        // belongs here.
        if (options?.forceDialog) {
          const root = editorRootRef.current
          if (root !== null) {
            root.focus({ preventScroll: true })
            syncSelectionToDom(root, range, documentModelRef.current)
          }
        }

        // Record *the revision that was actually written*
        // (`revisionToSave`, captured alongside `documentToSave` above,
        // including any header/footer edit `flushHeaderFooterEdits` just
        // committed) as the new saved baseline. We deliberately do NOT also
        // call `setDirty(false)` here: if the user kept editing while the
        // `await`s above were in flight, `documentRevision` may already have
        // moved on since this snapshot, and the effect below recomputes
        // dirty from whatever the *current* render's `documentRevision` is
        // once this state update lands — never from this stale closure.
        setLastSavedRevision(revisionToSave)

        // Round-trip fidelity audit, DXS round 2 follow-up — tell the user,
        // once per document, when a save couldn't fully preserve something
        // (see `buildFidelitySaveWarningMessage`'s doc comment for the
        // wording rules). Fire-and-forget: the file is already written by
        // this point (`result.saved` above), so this is purely an
        // informational follow-up, never something the save itself should
        // wait on or that should block/undo an already-successful write.
        // Guarded by `warnedAboutFidelityRef` (see its own doc comment) so
        // this only ever surfaces once for a given document, not on every
        // save that happens to still be affected.
        if (!warnedAboutFidelityRef.current) {
          void detectLossySaveWarnings(bundle, nextBytes)
            .then((warnings) => {
              const message = buildFidelitySaveWarningMessage(categorizeLossySaveWarnings(warnings), t)
              if (message !== null) {
                warnedAboutFidelityRef.current = true
                setFidelityWarningMessage(message)
              }
            })
            .catch(() => {
              // Best-effort notice only — a failure detecting/describing it
              // must never surface as a save error; the save already
              // succeeded.
            })
        }

        return true
      } catch (error) {
        // RUN-14 — wrap JSZip/fast-xml-parser/DocxSaveError internals in a
        // friendly, actionable message instead of showing the raw exception.
        setSaveError(friendlyDocxErrorMessage(error, 'save'))
        return false
      }
    },
    [bundle, flushHeaderFooterEdits, getRevision, savePath, filePath, range, reportSavedPath, setSaveError, t, documentModelRef, editorRootRef],
  )

  const handleSave = useCallback((): Promise<boolean> => handleSaveInternal(), [handleSaveInternal])
  const handleSaveAs = useCallback(
    (): Promise<boolean> => handleSaveInternal({ forceDialog: true }),
    [handleSaveInternal],
  )

  // See the "dirty is an integer compare" note in this module's header. Derived
  // in an effect from both pieces of state together, rather than set from inside
  // the save, so it stays correct when an edit lands while a save is in flight:
  // this always compares the CURRENT render's revision, never a snapshot taken
  // before the write resolved.
  useEffect(() => {
    setDirty(documentRevision !== lastSavedRevision)
  }, [documentRevision, lastSavedRevision, setDirty])

  useEffect(() => {
    registerSave(handleSave)
    return () => registerSave(null)
  }, [registerSave, handleSave])

  // Same registration for the global Ctrl+Shift+S / App.tsx `saveFileAs()` —
  // previously unregistered, so that shortcut silently did nothing for a DOCX
  // document even though the toolbar's own "Save As" button worked.
  useEffect(() => {
    registerSaveAs(handleSaveAs)
    return () => registerSaveAs(null)
  }, [registerSaveAs, handleSaveAs])

  return { savePath, handleSave, handleSaveAs, fidelityWarningMessage, dismissFidelityWarning }
}
