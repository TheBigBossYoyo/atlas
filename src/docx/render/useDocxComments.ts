/**
 * The DOCX viewer's comments pane: whether it is open, which threads the reader
 * has resolved, and the mutations that act on a comment.
 *
 * Extracted from `src/viewers/DocxViewer.tsx` as part of breaking up that file's
 * one very large component. It came out later than the pagination/find/selection
 * hooks because it is the first cluster that mutates the document, and that
 * needed the mutation surface named rather than handed over wholesale — see
 * `bumpRevision` below, which is the one narrow slice of the undo history this
 * needs instead of the whole `History` instance.
 *
 * **Resolved is a document edit, not a session view** (fixed 2026-09-26). It used
 * to be the opposite: the pane filtered a local set of ids and wrote nothing, so
 * a thread resolved in Word showed as open in Atlas and one resolved in Atlas was
 * forgotten the moment the file was reopened — even though the parser already read
 * `w15:done` into `Comment.resolved` and the serializer already wrote it back.
 *
 * Resolving now goes through `setCommentResolved`, so it round-trips: the pane
 * filters on `Comment.resolved` (whatever its source — this document's own
 * `commentsExtended.xml`, or a resolve done here), and `docx/index.ts` registers
 * the part, its relationship and its content-type on save when any comment
 * carries resolved state, so it works for a document that never had one.
 *
 * It follows that resolving marks the document dirty, which is correct — it is a
 * change that needs saving, and Word behaves the same way.
 *
 * The three writes below all share one non-obvious requirement, marked DIRTY-1
 * at each site in the original: a comment edit is NOT undoable through the
 * command history, so nothing bumps the revision that the dirty check compares
 * against. Without an explicit `bumpRevision()` the edit is silently missed and
 * the document looks saved when it is not.
 */
import { useCallback, useMemo, useState } from 'react'

import {
  addCommentToDocument,
  deleteCommentFromDocument,
  replyToComment,
  setCommentResolved,
} from '../editor/commentMutations'
import type { Range } from '../editor'
import type { Document as DocxDocument } from '../model'

export type DocxComments = {
  readonly commentsPaneOpen: boolean
  readonly setCommentsPaneOpen: (open: boolean) => void
  /** The model with resolved threads filtered out — what the pane renders. */
  readonly commentsDocument: DocxDocument
  readonly handleResolveComment: (commentId: string) => void
  readonly handleDeleteComment: (commentId: string) => void
  /**
   * Adds a comment on `selection` with the text the prompt dialog collected.
   * Lives here rather than in the viewer's prompt handler because it needs
   * exactly what this hook already holds, and because "add a comment" is a
   * comment operation, not a dialog one.
   */
  readonly confirmComment: (selection: Range, text: string) => void
  /** Replies to `commentId` with the text the prompt dialog collected. */
  readonly confirmReply: (commentId: string, range: Range | null, text: string) => void
}

/**
 * @param documentModel  - the live model.
 * @param range          - current selection, kept as the selection after a
 *                         delete (deleting a comment must not move the caret).
 * @param commitState    - the viewer's "here is the next document + selection"
 *                         entry point.
 * @param bumpRevision   - allocates a fresh document revision for an edit the
 *                         command history does not record. See the DIRTY-1 note
 *                         in this module's header for why every write needs it.
 * @param documentCommentCount - `bundle.document.comments.size`. Changing it
 *                         means a different file was opened, which resets both
 *                         the pane's default state and the resolved set.
 * @param setSaveError   - the viewer's shared error banner. Despite the name it
 *                         is not save-specific; it is where every failed editor
 *                         action in this viewer reports itself.
 */
export function useDocxComments(
  documentModel: DocxDocument,
  range: Range | null,
  commitState: (nextDocument: DocxDocument, nextRange: Range | null) => void,
  bumpRevision: () => void,
  documentCommentCount: number,
  setSaveError: (message: string | null) => void,
): DocxComments {
  const [commentsPaneOpen, setCommentsPaneOpen] = useState(documentCommentCount > 0)
  const [lastCommentCount, setLastCommentCount] = useState(documentCommentCount)

  // setState-during-render reset, the same idiom `viewers/shared/ViewerContext.tsx`
  // uses for its reset-on-filePath-change: a different comment count means a
  // different file, so the pane's default state and the session's resolved set
  // both start over. React discards the in-progress render and re-runs with the
  // new state, which is the lint-clean way to reset state when an input changes
  // — no render-time ref mutation and no cascading-render effect. As an effect
  // this showed the previous document's resolved set for one committed frame.
  if (lastCommentCount !== documentCommentCount) {
    setLastCommentCount(documentCommentCount)
    setCommentsPaneOpen(documentCommentCount > 0)
  }

  const handleResolveComment = useCallback(
    (commentId: string) => {
      const next = setCommentResolved(documentModel, commentId, true)
      if (next === documentModel) return
      bumpRevision()
      commitState(next, range)
    },
    [bumpRevision, commitState, documentModel, range],
  )

  const handleDeleteComment = useCallback(
    (commentId: string) => {
      bumpRevision()
      commitState(deleteCommentFromDocument(documentModel, commentId), range)
    },
    [bumpRevision, commitState, documentModel, range],
  )



  const confirmComment = useCallback(
    (selection: Range, text: string) => {
      try {
        const result = addCommentToDocument(documentModel, selection, text, 'Atlas')
        bumpRevision()
        commitState(result.document, selection)
        setCommentsPaneOpen(true)
      } catch (error) {
        setSaveError(error instanceof Error ? error.message : String(error))
      }
    },
    [bumpRevision, commitState, documentModel, setSaveError],
  )

  const confirmReply = useCallback(
    (commentId: string, range: Range | null, text: string) => {
      // A reply on a resolved thread re-opens it — leaving it filtered out would
      // hide the activity just added. Composed into the SAME document as the
      // reply and committed once, rather than as a second `commitState`: a
      // follow-up call would close over the pre-reply `documentModel` (this
      // render's copy, since setState has not re-rendered yet) and commit that,
      // silently discarding the reply. That is a bug this shape prevents, and it
      // is exactly what a second commit caused when it was written that way.
      const replied = replyToComment(documentModel, commentId, text, 'Atlas')
      bumpRevision()
      commitState(setCommentResolved(replied, commentId, false), range)
      setCommentsPaneOpen(true)
    },
    [bumpRevision, commitState, documentModel],
  )

  const commentsDocument = useMemo<DocxDocument>(() => {
    const comments = documentModel.comments
    const hasResolved = [...comments.values()].some((comment) => comment.resolved === true)
    if (!hasResolved) {
      return documentModel
    }

    const filtered = new Map(
      [...comments.entries()].filter(([, comment]) => {
        if (comment.resolved === true) {
          return false
        }
        // A reply whose parent thread is resolved goes with it. Replies have no
        // independent resolved state in the format, so the parent is the only
        // place to look.
        if (comment.parentId === undefined) return true
        return comments.get(comment.parentId)?.resolved !== true
      }),
    )

    return {
      ...documentModel,
      comments: filtered,
    }
  }, [documentModel])

  return {
    commentsPaneOpen,
    setCommentsPaneOpen,
    commentsDocument,
    handleResolveComment,
    handleDeleteComment,
    confirmComment,
    confirmReply,
  }
}
