import { memo, useMemo } from 'react'

import { Check, MessageSquare, Reply, Trash2, Undo2 } from 'lucide-react'

import type { Comment, Document as DocxDocument } from '../model/document'
import { extractCommentText, findCommentAnchors } from './comments'
import { useTranslate } from '../../i18n'
import './__styles__/comments-pane.css'

export interface CommentsPaneProps {
  readonly document: DocxDocument
  readonly onScrollToParagraph: (paragraphIndex: number) => void
  readonly onAddComment: () => void
  readonly onReply: (commentId: string) => void
  readonly onResolve: (commentId: string) => void
  readonly onUnresolve: (commentId: string) => void
  readonly onDelete: (commentId: string) => void
  /** When true, resolved threads are listed (marked, with Unresolve in place of Resolve). */
  readonly showResolved: boolean
  readonly onToggleShowResolved: () => void
}

interface CommentThread {
  readonly id: string
  readonly comment: Comment
  readonly anchorParagraphIndex: number | null
  readonly anchorKey: string
  readonly replies: ReadonlyArray<CommentThread>
}

interface CommentGroup {
  readonly anchorKey: string
  readonly anchorParagraphIndex: number | null
  readonly threads: ReadonlyArray<CommentThread>
}

function CommentsPaneBase({
  document,
  onScrollToParagraph,
  onAddComment,
  onReply,
  onResolve,
  onUnresolve,
  onDelete,
  showResolved,
  onToggleShowResolved,
}: CommentsPaneProps) {
  const t = useTranslate()
  const groups = useMemo(() => buildCommentGroups(document), [document])
  const count = document.comments.size

  return (
    <aside className="comments-pane" aria-label={t('docx.comments.aria')}>
      <header className="comments-pane__header">
        <div className="comments-pane__title">
          <MessageSquare size={16} aria-hidden="true" />
          <span>{t('docx.comments.aria')}</span>
        </div>
        <button type="button" className="comments-pane__add-button" onClick={onAddComment}>
          {t('docx.comments.addComment')}
        </button>
      </header>

      {/* Resolving is a document edit that survives reopening the file, so this is the
          only way back from a mis-click. Off by default: getting a thread out of the
          way is the point of resolving it. */}
      <label className="comments-pane__show-resolved">
        <input type="checkbox" checked={showResolved} onChange={onToggleShowResolved} />
        <span>{t('docx.comments.showResolved')}</span>
      </label>

      <div className="comments-pane__count">{t('docx.comments.count', { count })}</div>

      {groups.length === 0 ? (
        <p className="comments-pane__empty">{t('docx.comments.empty')}</p>
      ) : (
        <div className="comments-pane__groups">
          {groups.map((group) => (
            <section key={group.anchorKey} className="comments-pane__group">
              <button
                type="button"
                className="comments-pane__anchor"
                onClick={() => {
                  if (group.anchorParagraphIndex !== null) {
                    onScrollToParagraph(group.anchorParagraphIndex)
                  }
                }}
                disabled={group.anchorParagraphIndex === null}
              >
                {group.anchorParagraphIndex === null
                  ? t('docx.comments.unanchoredThread')
                  : t('docx.comments.anchorParagraph', { n: group.anchorParagraphIndex + 1 })}
              </button>

              <div className="comments-pane__thread-list">
                {group.threads.map((thread) => (
                  <CommentThreadView
                    key={thread.id}
                    thread={thread}
                    onScrollToParagraph={onScrollToParagraph}
                    onReply={onReply}
                    onResolve={onResolve}
                    onUnresolve={onUnresolve}
                    onDelete={onDelete}
                  />
                ))}
              </div>
            </section>
          ))}
        </div>
      )}
    </aside>
  )
}

interface CommentThreadViewProps {
  readonly thread: CommentThread
  readonly onScrollToParagraph: (paragraphIndex: number) => void
  readonly onReply: (commentId: string) => void
  readonly onResolve: (commentId: string) => void
  readonly onUnresolve: (commentId: string) => void
  readonly onDelete: (commentId: string) => void
}

function CommentThreadView({
  thread,
  onScrollToParagraph,
  onReply,
  onResolve,
  onUnresolve,
  onDelete,
}: CommentThreadViewProps) {
  const t = useTranslate()
  const author = thread.comment.author ?? t('docx.comments.unknownAuthor')
  const text = extractCommentText(thread.comment)
  const dateLabel = formatCommentDate(thread.comment.date)
  const isResolved = thread.comment.resolved === true

  return (
    <article className={`comments-pane__item${isResolved ? ' comments-pane__item--resolved' : ''}`}>
      <button
        type="button"
        className="comments-pane__card"
        onClick={() => {
          if (thread.anchorParagraphIndex !== null) {
            onScrollToParagraph(thread.anchorParagraphIndex)
          }
        }}
      >
        <div className="comments-pane__meta">
          <span className="comments-pane__author">{author}</span>
          {dateLabel !== null ? <span className="comments-pane__date">{dateLabel}</span> : null}
        </div>
        <div className="comments-pane__body">{text.length > 0 ? text : t('docx.comments.emptyCommentBody')}</div>
        {/* Text, not only a CSS style: a screen reader has nothing to go on otherwise. */}
        {isResolved ? <div className="comments-pane__resolved-badge">{t('docx.comments.resolvedBadge')}</div> : null}
      </button>

      <div className="comments-pane__actions">
        <button type="button" className="comments-pane__action" onClick={() => onReply(thread.id)}>
          <Reply size={14} aria-hidden="true" />
          <span>{t('docx.comments.reply')}</span>
        </button>
        {isResolved ? (
          <button type="button" className="comments-pane__action" onClick={() => onUnresolve(thread.id)}>
            <Undo2 size={14} aria-hidden="true" />
            <span>{t('docx.comments.unresolve')}</span>
          </button>
        ) : (
          <button type="button" className="comments-pane__action" onClick={() => onResolve(thread.id)}>
            <Check size={14} aria-hidden="true" />
            <span>{t('docx.comments.resolve')}</span>
          </button>
        )}
        <button type="button" className="comments-pane__action comments-pane__action--danger" onClick={() => onDelete(thread.id)}>
          <Trash2 size={14} aria-hidden="true" />
          <span>{t('docx.comments.delete')}</span>
        </button>
      </div>

      {thread.replies.length > 0 ? (
        <div className="comments-pane__replies">
          {thread.replies.map((reply) => (
            <CommentThreadView
              key={reply.id}
              thread={reply}
              onScrollToParagraph={onScrollToParagraph}
              onReply={onReply}
              onResolve={onResolve}
              onUnresolve={onUnresolve}
              onDelete={onDelete}
            />
          ))}
        </div>
      ) : null}
    </article>
  )
}

function buildCommentGroups(document: DocxDocument): ReadonlyArray<CommentGroup> {
  const anchorMap = findCommentAnchors(document)
  const commentMap = new Map(document.comments)
  const replyMap = new Map<string, CommentThread[]>()
  const roots: CommentThread[] = []

  for (const [id, comment] of commentMap) {
    const anchorParagraphIndex = anchorMap.get(id) ?? (comment.parentId !== undefined ? anchorMap.get(comment.parentId) ?? null : null)
    const thread: CommentThread = {
      id,
      comment,
      anchorParagraphIndex,
      anchorKey: anchorParagraphIndex === null ? `orphan-${comment.parentId ?? id}` : String(anchorParagraphIndex),
      replies: [],
    }

    if (comment.parentId !== undefined && commentMap.has(comment.parentId)) {
      const existing = replyMap.get(comment.parentId) ?? []
      replyMap.set(comment.parentId, [...existing, thread])
    } else {
      roots.push(thread)
    }
  }

  const hydratedRoots = roots.map((thread) => hydrateReplies(thread, replyMap))
  const groups = new Map<string, CommentGroup>()

  for (const thread of hydratedRoots) {
    const existing = groups.get(thread.anchorKey)
    if (existing === undefined) {
      groups.set(thread.anchorKey, {
        anchorKey: thread.anchorKey,
        anchorParagraphIndex: thread.anchorParagraphIndex,
        threads: [thread],
      })
      continue
    }

    groups.set(thread.anchorKey, {
      ...existing,
      threads: [...existing.threads, thread],
    })
  }

  return [...groups.values()].sort((left, right) => {
    if (left.anchorParagraphIndex === null) {
      return 1
    }
    if (right.anchorParagraphIndex === null) {
      return -1
    }
    return left.anchorParagraphIndex - right.anchorParagraphIndex
  })
}

function hydrateReplies(thread: CommentThread, replyMap: ReadonlyMap<string, ReadonlyArray<CommentThread>>): CommentThread {
  const replies = replyMap.get(thread.id) ?? []
  return {
    ...thread,
    replies: replies.map((reply) => hydrateReplies(reply, replyMap)),
  }
}

function formatCommentDate(rawDate: string | undefined): string | null {
  if (rawDate === undefined) {
    return null
  }

  const date = new Date(rawDate)
  if (Number.isNaN(date.getTime())) {
    return rawDate
  }

  return date.toLocaleString(undefined, {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}

export const CommentsPane = memo(CommentsPaneBase)
