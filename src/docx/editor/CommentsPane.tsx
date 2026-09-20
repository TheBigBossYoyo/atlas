import { memo, useMemo } from 'react'

import { Check, MessageSquare, Reply, Trash2 } from 'lucide-react'

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
  readonly onDelete: (commentId: string) => void
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
  onDelete,
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
  readonly onDelete: (commentId: string) => void
}

function CommentThreadView({
  thread,
  onScrollToParagraph,
  onReply,
  onResolve,
  onDelete,
}: CommentThreadViewProps) {
  const t = useTranslate()
  const author = thread.comment.author ?? t('docx.comments.unknownAuthor')
  const text = extractCommentText(thread.comment)
  const dateLabel = formatCommentDate(thread.comment.date)

  return (
    <article className="comments-pane__item">
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
      </button>

      <div className="comments-pane__actions">
        <button type="button" className="comments-pane__action" onClick={() => onReply(thread.id)}>
          <Reply size={14} aria-hidden="true" />
          <span>{t('docx.comments.reply')}</span>
        </button>
        <button type="button" className="comments-pane__action" onClick={() => onResolve(thread.id)}>
          <Check size={14} aria-hidden="true" />
          <span>{t('docx.comments.resolve')}</span>
        </button>
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
