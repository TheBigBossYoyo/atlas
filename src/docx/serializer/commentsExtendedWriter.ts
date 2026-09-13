/**
 * Atlas — DOCX commentsExtended.xml writer (D16 / DXS-11)
 *
 * Inverse of `parser/commentsExtended.ts`. Writes one `<w15:commentEx>` per
 * comment whose `resolved` state is known, keyed by the same `w15:paraId`
 * join key the parser reads.
 */

import { XMLBuilder } from 'fast-xml-parser'

import type { Comment } from '../model'

const WORD_2012_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml'
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

/**
 * Resolves the join key a comment's `<w15:commentEx>` entry should use.
 *
 * Prefers the `w14:paraId` Word itself stamped onto the comment's first
 * body paragraph (captured as `Paragraph.paraId`, D19/DXS-10) — the same
 * key a real Word-authored `commentsExtended.xml` uses. Falls back to the
 * comment's own `id` when no paraId was captured (a comment created by
 * Atlas's own editor, or one from a document old enough to lack paraIds),
 * so Atlas's own resolve -> save -> reopen round-trip still works even
 * though that fallback key won't correlate with a paraId Word itself would
 * recognize.
 */
export function resolveCommentExtendedKey(comment: Comment): string {
  return comment.body[0]?.paraId ?? comment.id
}

export function writeCommentsExtendedXml(comments: ReadonlyArray<Comment>): string {
  const entries = comments.filter((comment) => comment.resolved !== undefined)

  const root = {
    'w15:commentsEx': {
      '@_xmlns:w15': WORD_2012_NAMESPACE,
      'w15:commentEx': entries.map((comment) => ({
        '@_w15:paraId': resolveCommentExtendedKey(comment),
        '@_w15:done': comment.resolved ? '1' : '0',
      })),
    },
  }

  return `${XML_DECLARATION}${xmlBuilder.build(root)}`
}
