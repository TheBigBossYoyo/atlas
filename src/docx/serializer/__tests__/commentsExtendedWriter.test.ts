/**
 * Tests for src/docx/serializer/commentsExtendedWriter.ts (D16 / DXS-11)
 */
import { describe, expect, it } from 'vitest'

import type { Comment } from '../../model'
import { parseCommentsExtended } from '../../parser/commentsExtended'
import { resolveCommentExtendedKey, writeCommentsExtendedXml } from '../commentsExtendedWriter'

function createComment(overrides: Partial<Comment> & Pick<Comment, 'id'>): Comment {
  return {
    kind: 'comment',
    body: [],
    ...overrides,
  }
}

describe('resolveCommentExtendedKey', () => {
  it('prefers the paraId captured on the comment body\'s first paragraph', () => {
    const comment = createComment({
      id: '0',
      body: [{ kind: 'paragraph', paraId: '12AB34CD', children: [] }],
    })

    expect(resolveCommentExtendedKey(comment)).toBe('12AB34CD')
  })

  it('falls back to the comment id when no paraId was captured', () => {
    const comment = createComment({
      id: '7',
      body: [{ kind: 'paragraph', children: [] }],
    })

    expect(resolveCommentExtendedKey(comment)).toBe('7')
  })

  it('falls back to the comment id when the comment body is empty', () => {
    const comment = createComment({ id: '3', body: [] })

    expect(resolveCommentExtendedKey(comment)).toBe('3')
  })
})

describe('writeCommentsExtendedXml', () => {
  it('emits nothing for comments with no resolved state', () => {
    const comments = [createComment({ id: '0' })]

    const xml = writeCommentsExtendedXml(comments)

    expect(xml).not.toContain('w15:commentEx')
  })

  it('emits a done="1" entry for a resolved comment', () => {
    const comments = [
      createComment({
        id: '0',
        resolved: true,
        body: [{ kind: 'paragraph', paraId: '12AB34CD', children: [] }],
      }),
    ]

    const xml = writeCommentsExtendedXml(comments)

    expect(xml).toContain('w15:paraId="12AB34CD"')
    expect(xml).toContain('w15:done="1"')
  })

  it('emits a done="0" entry for an explicitly-unresolved comment', () => {
    const comments = [
      createComment({
        id: '0',
        resolved: false,
        body: [{ kind: 'paragraph', paraId: '12AB34CD', children: [] }],
      }),
    ]

    const xml = writeCommentsExtendedXml(comments)

    expect(xml).toContain('w15:done="0"')
  })

  it('round-trips resolved state through parse -> write -> parse', () => {
    const comments = [
      createComment({
        id: '0',
        resolved: true,
        body: [{ kind: 'paragraph', paraId: '12AB34CD', children: [] }],
      }),
      createComment({
        id: '1',
        resolved: false,
        body: [{ kind: 'paragraph', paraId: 'AAAA0002', children: [] }],
      }),
    ]

    const xml = writeCommentsExtendedXml(comments)
    const reparsed = parseCommentsExtended(xml)

    expect(reparsed.get('12AB34CD')).toBe(true)
    expect(reparsed.get('AAAA0002')).toBe(false)
  })
})
