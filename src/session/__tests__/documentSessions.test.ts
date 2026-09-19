/** SHELL-17 — the open-documents model. */
import { describe, expect, it, vi } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import {
  EMPTY_SESSIONS,
  activateSession,
  activeSession,
  closeSession,
  moveSession,
  neighbourSession,
  openSession,
  renameSession,
  reloadSessionFile,
  reopenLastClosed,
} from '../documentSessions'

function file(path: string, content = 'x'): LoadedFile {
  return { kind: 'text', content, path, format: 'markdown' }
}

const openAll = (...paths: ReadonlyArray<string>) =>
  paths.reduce((state, path) => openSession(state, file(path)), EMPTY_SESSIONS)

describe('opening', () => {
  it('appends a document and makes it active', () => {
    const state = openAll('/a.md', '/b.docx')
    expect(state.sessions.map((session) => session.name)).toEqual(['a.md', 'b.docx'])
    expect(activeSession(state)?.id).toBe('/b.docx')
  })

  it('re-activates (and refreshes) a path that is already open instead of duplicating it', () => {
    const state = openSession(openAll('/a.md', '/b.docx'), file('/a.md', 'reloaded'))
    expect(state.sessions).toHaveLength(2)
    expect(activeSession(state)?.id).toBe('/a.md')
    expect(activeSession(state)?.file.kind === 'text' && activeSession(state)?.file.content).toBe('reloaded')
  })
})

describe('closing', () => {
  it('activates the document on the right, or the left when it was last', () => {
    const three = openAll('/a.md', '/b.md', '/c.md')
    const closedMiddle = closeSession(activateSession(three, '/b.md'), '/b.md')
    expect(activeSession(closedMiddle)?.id).toBe('/c.md')

    const closedLast = closeSession(three, '/c.md')
    expect(activeSession(closedLast)?.id).toBe('/b.md')
  })

  it('leaves nothing active once the last document is closed', () => {
    const state = closeSession(openAll('/a.md'), '/a.md')
    expect(state.sessions).toHaveLength(0)
    expect(state.activeId).toBeNull()
  })

  it('closing an inactive document does not change which one is showing', () => {
    const state = closeSession(openAll('/a.md', '/b.md'), '/a.md')
    expect(activeSession(state)?.id).toBe('/b.md')
  })

  it('reopens the last closed document with the bytes it had', () => {
    const closed = closeSession(openAll('/a.md', '/b.md'), '/b.md')
    const reopened = reopenLastClosed(closed)
    expect(activeSession(reopened)?.id).toBe('/b.md')
    expect(reopenLastClosed(reopened)).toBe(reopened) // nothing left to reopen
  })

  it('remembers only the last few closed documents', () => {
    let state = openAll('/1.md', '/2.md', '/3.md', '/4.md', '/5.md', '/6.md', '/7.md')
    for (const path of ['/1.md', '/2.md', '/3.md', '/4.md', '/5.md', '/6.md']) {
      state = closeSession(state, path)
    }
    expect(state.recentlyClosed).toHaveLength(5)
    expect(state.recentlyClosed[0].id).toBe('/6.md')
  })
})

describe('navigating and reordering', () => {
  it('walks to the next and previous document, wrapping around', () => {
    const state = openAll('/a.md', '/b.md', '/c.md')
    expect(neighbourSession(state, 1)?.id).toBe('/a.md')
    expect(neighbourSession(state, -1)?.id).toBe('/b.md')
    expect(neighbourSession(openAll('/only.md'), 1)).toBeNull()
  })

  it('moves a document to another position', () => {
    const state = moveSession(openAll('/a.md', '/b.md', '/c.md'), 2, 0)
    expect(state.sessions.map((session) => session.name)).toEqual(['c.md', 'a.md', 'b.md'])
    expect(activeSession(state)?.id).toBe('/c.md')
    expect(moveSession(state, 0, 9)).toBe(state)
  })

  it('follows a document to its new path after Save As', () => {
    const state = renameSession(openAll('/a.md', '/b.md'), '/b.md', file('/renamed.md'))
    expect(state.sessions.map((session) => session.name)).toEqual(['a.md', 'renamed.md'])
    expect(activeSession(state)?.id).toBe('/renamed.md')
  })
})

describe('reloadSessionFile', () => {
  it('shows a text document as it is on disk now, not as it was opened', async () => {
    const reader = { openFileByPath: vi.fn(async () => ({ content: 'saved later' })), readBinaryByPath: vi.fn() }
    const reloaded = await reloadSessionFile(file('/a.md', 'as opened'), reader)
    expect(reader.openFileByPath).toHaveBeenCalledWith('/a.md')
    expect(reloaded).toEqual(file('/a.md', 'saved later'))
  })

  it('re-reads a binary document too', async () => {
    const fresh = new ArrayBuffer(4)
    const stored: LoadedFile = { kind: 'binary', content: new ArrayBuffer(2), path: '/b.docx', format: 'docx' }
    const reader = { openFileByPath: vi.fn(), readBinaryByPath: vi.fn(async () => ({ buffer: fresh })) }
    const reloaded = await reloadSessionFile(stored, reader)
    expect(reloaded.content).toBe(fresh)
    expect(reloaded.path).toBe('/b.docx')
  })

  it('keeps the stored bytes when the file can no longer be read', async () => {
    const stored = file('/gone.md', 'kept')
    const failing = {
      openFileByPath: vi.fn(async () => {
        throw new Error('ENOENT')
      }),
      readBinaryByPath: vi.fn(),
    }
    expect(await reloadSessionFile(stored, failing)).toBe(stored)
    expect(await reloadSessionFile(stored, { openFileByPath: vi.fn(async () => null), readBinaryByPath: vi.fn() })).toBe(stored)
    expect(await reloadSessionFile(stored, undefined)).toBe(stored)
  })
})
