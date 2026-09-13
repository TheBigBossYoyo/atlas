import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { CommentsPane } from '../CommentsPane'
import type { Document } from '../../model'

function createDocument(): Document {
  return {
    kind: 'document',
    sections: [
      {
        kind: 'section',
        props: {},
        blocks: [
          {
            kind: 'paragraph',
            children: [
              { kind: 'comment-range', id: '1', boundary: 'start' },
              { kind: 'run', children: [{ kind: 'text', value: 'Hello' }] },
              { kind: 'comment-range', id: '1', boundary: 'end' },
              { kind: 'run', children: [{ kind: 'comment-reference', id: '1' }] },
            ],
          },
        ],
      },
    ],
    styles: new Map(),
    numbering: new Map(),
    comments: new Map([
      [
        '1',
        {
          kind: 'comment',
          id: '1',
          author: 'Alice',
          date: '2024-01-15T10:30:00Z',
          body: [
            {
              kind: 'paragraph',
              children: [{ kind: 'run', children: [{ kind: 'text', value: 'Root comment' }] }],
            },
          ],
        },
      ],
      [
        '2',
        {
          kind: 'comment',
          id: '2',
          author: 'Bob',
          parentId: '1',
          body: [
            {
              kind: 'paragraph',
              children: [{ kind: 'run', children: [{ kind: 'text', value: 'Reply comment' }] }],
            },
          ],
        },
      ],
    ]),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

describe('CommentsPane', () => {
  it('renders grouped comments with threaded replies and actions', () => {
    const onScrollToParagraph = vi.fn()
    const onAddComment = vi.fn()
    const onReply = vi.fn()
    const onResolve = vi.fn()
    const onDelete = vi.fn()

    render(
      <CommentsPane
        document={createDocument()}
        onScrollToParagraph={onScrollToParagraph}
        onAddComment={onAddComment}
        onReply={onReply}
        onResolve={onResolve}
        onDelete={onDelete}
      />,
    )

    expect(screen.getByText('Root comment')).toBeInTheDocument()
    expect(screen.getByText('Reply comment')).toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: 'Add comment' }))
    expect(onAddComment).toHaveBeenCalledTimes(1)

    fireEvent.click(screen.getByRole('button', { name: /anchor paragraph 1/i }))
    expect(onScrollToParagraph).toHaveBeenCalledWith(0)

    fireEvent.click(screen.getAllByRole('button', { name: /reply/i })[0])
    expect(onReply).toHaveBeenCalledWith('1')

    fireEvent.click(screen.getAllByRole('button', { name: /resolve/i })[0])
    expect(onResolve).toHaveBeenCalledWith('1')

    fireEvent.click(screen.getAllByRole('button', { name: /delete/i })[0])
    expect(onDelete).toHaveBeenCalledWith('1')
  })
})
