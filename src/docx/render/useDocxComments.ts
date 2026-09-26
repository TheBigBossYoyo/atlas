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
 * **Resolved is a session view, not a document edit.** Resolving hides a thread
 * from the pane; it writes nothing to the file, which is why `resolvedCommentIds`
 * is local state and `commentsDocument` is a filtered *view* of the model handed
 * to the pane rather than a mutation of the model itself.
 *
 * This is deliberately NOT connected to Word's own resolved flag, and the
 * disconnect runs both ways (verified while extracting this, and now recorded in
 * docs/KNOWN_LIMITATIONS.md): `w15:done` is parsed into `Comment.resolved` by
 * `docx/parser/commentsExtended.ts` and written back by
 * `serializer/commentsExtendedWriter.ts`, so a file's flags survive a round trip
 * — but nothing reads `Comment.resolved` to filter the pane, and resolving here
 * never sets it. A thread resolved in Word therefore shows as open in Atlas, and
 * one resolved in Atlas is forgotten when the file is reopened.
 *
 * The three writes below all share one non-obvious requirement, marked DIRTY-1
 * at each site in the original: a comment edit is NOT undoable through the
 * command history, so nothing bumps the revision that the dirty check compares
 * against. Without an explicit `bumpRevision()` the edit is silently missed and
 * the document looks saved when it is not.
 */
import { useCallback, useMemo, useState } from 'react'

import { addCommentToDocument, deleteCommentFromDocument, replyToComment } from '../editor/commentMutations'
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
  /**
   * Call after a new comment or reply lands on `commentId`: opens the pane and
   * clears any stale resolved flag, so a thread the reader had resolved comes
   * back into view when it gets new activity.
   */
  readonly revealCommentThread: (commentId: string) => void
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
  const [resolvedCommentIds, setResolvedCommentIds] = useState<ReadonlySet<string>>(new Set())
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
    setResolvedCommentIds(new Set())
  }

  const unresolve = useCallback((commentId: string) => {
    setResolvedCommentIds((current) => {
      const next = new Set(current)
      next.delete(commentId)
      return next
    })
  }, [])

  const handleResolveComment = useCallback((commentId: string) => {
    setResolvedCommentIds((current) => new Set(current).add(commentId))
  }, [])

  const handleDeleteComment = useCallback(
    (commentId: string) => {
      bumpRevision()
      commitState(deleteCommentFromDocument(documentModel, commentId), range)
      unresolve(commentId)
    },
    [bumpRevision, commitState, documentModel, range, unresolve],
  )

  const revealCommentThread = useCallback(
    (commentId: string) => {
      setCommentsPaneOpen(true)
      unresolve(commentId)
    },
    [unresolve],
  )

  const confirmComment = useCallback(
    (selection: Range, text: string) => {
      try {
        const result = addCommentToDocument(documentModel, selection, text, 'Atlas')
        bumpRevision()
        commitState(result.document, selection)
        revealCommentThread(result.commentId)
      } catch (error) {
        setSaveError(error instanceof Error ? error.message : String(error))
      }
    },
    [bumpRevision, commitState, documentModel, revealCommentThread, setSaveError],
  )

  const confirmReply = useCallback(
    (commentId: string, range: Range | null, text: string) => {
      bumpRevision()
      commitState(replyToComment(documentModel, commentId, text, 'Atlas'), range)
      revealCommentThread(commentId)
    },
    [bumpRevision, commitState, documentModel, revealCommentThread],
  )

  const commentsDocument = useMemo<DocxDocument>(() => {
    if (resolvedCommentIds.size === 0) {
      return documentModel
    }

    const filtered = new Map(
      [...documentModel.comments.entries()].filter(([id, comment]) => {
        if (resolvedCommentIds.has(id)) {
          return false
        }
        // A reply whose parent thread is resolved goes with it, even though the
        // reply's own id was never resolved.
        return comment.parentId === undefined || !resolvedCommentIds.has(comment.parentId)
      }),
    )

    return {
      ...documentModel,
      comments: filtered,
    }
  }, [documentModel, resolvedCommentIds])

  return {
    commentsPaneOpen,
    setCommentsPaneOpen,
    commentsDocument,
    handleResolveComment,
    handleDeleteComment,
    revealCommentThread,
    confirmComment,
    confirmReply,
  }
}
