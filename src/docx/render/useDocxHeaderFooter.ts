/**
 * The DOCX viewer's header/footer editing panel.
 *
 * Extracted from `src/viewers/DocxViewer.tsx` as part of breaking up that file's
 * one very large component. Like `useDocxComments`, it mutates the document, so
 * it takes named slices of the mutation surface (`commitState`,
 * `applyEditorCommand`, `pushUndo`) rather than the viewer's internals.
 *
 * **The pending-edit map is the whole point of this module.** A header/footer
 * field commits as ONE undo step on blur, not per keystroke, so `onChange` only
 * remembers the latest uncommitted value here. That is also what fixes
 * Ctrl+S-while-a-field-is-focused: the field used to commit only on blur, so a
 * save (which reads `documentModel`) while the field still had focus wrote the
 * PRE-edit text. `flushHeaderFooterEdits` exists for the save path to call
 * first, and returns the resulting document SYNCHRONOUSLY, because
 * `commitState`'s own `setDocumentModel` is async and the save needs the
 * just-committed text in hand immediately rather than on the next render.
 *
 * Two details in `flushHeaderFooterEdits` that look like over-engineering and
 * are not:
 *   - edits are applied ONE AT A TIME, each against the document the previous
 *     one produced, rather than pre-computing every command against the same
 *     snapshot. Two different pending edits can target text segments of the very
 *     SAME paragraph (a `'mixed'` row has one input per segment), and building
 *     both from one stale snapshot makes the second overwrite the first.
 *   - a failing command is skipped rather than aborting the flush, so one bad
 *     field cannot cost the user every other pending edit at save time.
 */
import { useCallback, useEffect, useMemo, useRef, useState, type Dispatch, type RefObject, type SetStateAction } from 'react'

import {
  applyCommand,
  buildHeaderFooterSegmentEdit,
  buildHeaderFooterTextEdit,
  buildInsertHeaderFooterParagraph,
  buildRemoveHeaderFooterParagraph,
  listHeaderFooterParts,
  type Command,
  type HeaderFooterKind,
  type Range,
} from '../editor'
import type { Document as DocxDocument } from '../model'

type PendingHeaderFooterEdit = {
  readonly kind: HeaderFooterKind
  readonly id: string
  readonly blockIndex: number
  /** Present only for a `'mixed'` row's text segment; absent for a plain `'text'` row's single field. */
  readonly segmentIndex?: number
  readonly text: string
}

/**
 * `segmentIndex` is omitted for a plain `'text'` row's single field and present
 * for one text segment of a `'mixed'` row, so the key folds in `'text'` for the
 * former — otherwise a plain row would collide with an actual segment 0.
 */
function pendingKey(kind: HeaderFooterKind, id: string, blockIndex: number, segmentIndex?: number): string {
  return `${kind}:${id}:${blockIndex}:${segmentIndex ?? 'text'}`
}

/** Builds the right command for a pending edit — a whole-paragraph rewrite for a
 * plain `'text'` row, or one text segment's rewrite for a `'mixed'` row — against
 * whatever document is passed in, which is not necessarily the latest
 * `documentModel` (see `flushHeaderFooterEdits`). */
function buildPendingCommand(doc: DocxDocument, edit: PendingHeaderFooterEdit): Command | null {
  return edit.segmentIndex === undefined
    ? buildHeaderFooterTextEdit(doc, edit.kind, edit.id, edit.blockIndex, edit.text)
    : buildHeaderFooterSegmentEdit(doc, edit.kind, edit.id, edit.blockIndex, edit.segmentIndex, edit.text)
}

export type DocxHeaderFooter = {
  readonly headerFooterOpen: boolean
  /** Full setter, not `(open: boolean) => void`: the toolbar toggle uses the
   * updater form (`setHeaderFooterOpen(open => !open)`). */
  readonly setHeaderFooterOpen: Dispatch<SetStateAction<boolean>>
  readonly headerFooterPanelRef: RefObject<HTMLDivElement | null>
  readonly headerFooterToggleRef: RefObject<HTMLButtonElement | null>
  readonly headerFooterParts: ReturnType<typeof listHeaderFooterParts>
  /**
   * Commits every still-pending field edit as one combined undo step and returns
   * the resulting document synchronously. A no-op returning `documentModel`
   * unchanged when nothing is pending. The save path must call this first.
   */
  readonly flushHeaderFooterEdits: () => DocxDocument
  readonly handleHeaderFooterFieldChange: (
    kind: HeaderFooterKind,
    id: string,
    blockIndex: number,
    text: string,
    segmentIndex?: number,
  ) => void
  readonly handleHeaderFooterFieldBlur: (
    kind: HeaderFooterKind,
    id: string,
    blockIndex: number,
    text: string,
    segmentIndex?: number,
  ) => void
  readonly handleAddHeaderFooterParagraph: (kind: HeaderFooterKind, id: string) => void
  readonly handleRemoveHeaderFooterParagraph: (kind: HeaderFooterKind, id: string, blockIndex: number) => void
}

/**
 * @param documentModel      - the live model.
 * @param range              - kept as the selection across a flush; editing a
 *                             header must not move the caret in the body.
 * @param commitState        - the viewer's "next document + selection" entry point.
 * @param applyEditorCommand - for the edits that commit immediately (blur, add,
 *                             remove), as opposed to the pending ones.
 * @param pushUndo           - records one inverse command on the undo history.
 *                             The narrow slice of `History` a flush needs; a
 *                             flush pushes its own combined inverse rather than
 *                             going through `applyEditorCommand`, since it
 *                             applies several commands as one step.
 */
export function useDocxHeaderFooter(
  documentModel: DocxDocument,
  range: Range | null,
  commitState: (nextDocument: DocxDocument, nextRange: Range | null) => void,
  applyEditorCommand: (command: Command) => boolean,
  pushUndo: (inverse: Command) => void,
): DocxHeaderFooter {
  const [headerFooterOpen, setHeaderFooterOpen] = useState(false)
  const headerFooterPanelRef = useRef<HTMLDivElement>(null)
  const headerFooterToggleRef = useRef<HTMLButtonElement>(null)
  const pendingEditsRef = useRef(new Map<string, PendingHeaderFooterEdit>())

  const headerFooterParts = useMemo(() => listHeaderFooterParts(documentModel), [documentModel])

  // UX — move focus into the panel (its close button) when it opens, and back to
  // the toolbar toggle that opened it once it closes, matching the focus
  // management every other shell dialog/panel already has.
  useEffect(() => {
    if (!headerFooterOpen) return
    const panel = headerFooterPanelRef.current
    const toggle = headerFooterToggleRef.current
    panel?.querySelector<HTMLElement>('button, input')?.focus()
    return () => {
      toggle?.focus()
    }
  }, [headerFooterOpen])

  const flushHeaderFooterEdits = useCallback((): DocxDocument => {
    const pending = pendingEditsRef.current
    if (pending.size === 0) {
      return documentModel
    }

    const edits = [...pending.values()]
    pending.clear()

    let workingDocument = documentModel
    const inverses: Command[] = []
    for (const edit of edits) {
      const command = buildPendingCommand(workingDocument, edit)
      if (command === null) {
        continue
      }
      try {
        const result = applyCommand(workingDocument, command)
        workingDocument = result.document
        inverses.push(result.inverse)
      } catch {
        // Skip this one; keep committing the rest.
      }
    }
    if (inverses.length === 0) {
      return documentModel
    }

    inverses.reverse()
    const inverse: Command = inverses.length === 1 ? inverses[0] : { kind: 'composite', commands: inverses }
    pushUndo(inverse)
    commitState(workingDocument, range)
    return workingDocument
  }, [commitState, documentModel, pushUndo, range])

  const handleHeaderFooterFieldChange = useCallback(
    (kind: HeaderFooterKind, id: string, blockIndex: number, text: string, segmentIndex?: number) => {
      pendingEditsRef.current.set(pendingKey(kind, id, blockIndex, segmentIndex), {
        kind,
        id,
        blockIndex,
        segmentIndex,
        text,
      })
    },
    [],
  )

  const handleHeaderFooterFieldBlur = useCallback(
    (kind: HeaderFooterKind, id: string, blockIndex: number, text: string, segmentIndex?: number) => {
      pendingEditsRef.current.delete(pendingKey(kind, id, blockIndex, segmentIndex))
      const command = buildPendingCommand(documentModel, { kind, id, blockIndex, segmentIndex, text })
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [applyEditorCommand, documentModel],
  )

  const handleAddHeaderFooterParagraph = useCallback(
    (kind: HeaderFooterKind, id: string) => {
      const command = buildInsertHeaderFooterParagraph(documentModel, kind, id)
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [applyEditorCommand, documentModel],
  )

  const handleRemoveHeaderFooterParagraph = useCallback(
    (kind: HeaderFooterKind, id: string, blockIndex: number) => {
      // Clears every pending edit for this paragraph — a `'mixed'` row can have
      // more than one (one per text segment), unlike a plain `'text'` row's
      // single field an exact-key delete would assume.
      const prefix = `${kind}:${id}:${blockIndex}:`
      for (const key of pendingEditsRef.current.keys()) {
        if (key.startsWith(prefix)) {
          pendingEditsRef.current.delete(key)
        }
      }
      const command = buildRemoveHeaderFooterParagraph(documentModel, kind, id, blockIndex)
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [applyEditorCommand, documentModel],
  )

  return {
    headerFooterOpen,
    setHeaderFooterOpen,
    headerFooterPanelRef,
    headerFooterToggleRef,
    headerFooterParts,
    flushHeaderFooterEdits,
    handleHeaderFooterFieldChange,
    handleHeaderFooterFieldBlur,
    handleAddHeaderFooterParagraph,
    handleRemoveHeaderFooterParagraph,
  }
}
