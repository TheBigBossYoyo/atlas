/**
 * SHELL-17 — open documents ("tabs").
 *
 * A pure, immutable model of which files are open, which one is showing, and
 * which ones were closed recently (so Ctrl+Shift+T can bring one back). The
 * shell owns the side effects: prompting about unsaved changes before it
 * switches away from a document, and handing the newly-activated session's
 * already-loaded bytes back to the file handler instead of re-reading disk.
 *
 * A file is identified by its path: opening a path that is already open just
 * activates that tab, exactly like every editor does.
 */
import type { LoadedFile } from '../formats/types'

export type DocumentSession = {
  /** The file's path — also the session's identity. */
  readonly id: string
  readonly name: string
  readonly file: LoadedFile
}

export type DocumentSessionsState = {
  readonly sessions: ReadonlyArray<DocumentSession>
  readonly activeId: string | null
  /** Most recently closed first; bounded, since each one keeps the file's bytes alive. */
  readonly recentlyClosed: ReadonlyArray<DocumentSession>
}

/** How many closed documents Ctrl+Shift+T can walk back through. */
const MAX_RECENTLY_CLOSED = 5

export const EMPTY_SESSIONS: DocumentSessionsState = { sessions: [], activeId: null, recentlyClosed: [] }

export function fileName(path: string): string {
  return path.replace(/\\/g, '/').split('/').pop() || path
}

function toSession(file: LoadedFile): DocumentSession {
  return { id: file.path, name: fileName(file.path), file }
}

export function activeSession(state: DocumentSessionsState): DocumentSession | null {
  return state.sessions.find((session) => session.id === state.activeId) ?? null
}

export function sessionAt(state: DocumentSessionsState, index: number): DocumentSession | null {
  return state.sessions[index] ?? null
}

export function activeIndex(state: DocumentSessionsState): number {
  return state.sessions.findIndex((session) => session.id === state.activeId)
}

/** Opens (or re-activates) a document, keeping its latest bytes. */
export function openSession(state: DocumentSessionsState, file: LoadedFile): DocumentSessionsState {
  const session = toSession(file)
  const existingIndex = state.sessions.findIndex((candidate) => candidate.id === session.id)
  const sessions =
    existingIndex >= 0
      ? state.sessions.map((candidate, index) => (index === existingIndex ? session : candidate))
      : [...state.sessions, session]
  return {
    sessions,
    activeId: session.id,
    recentlyClosed: state.recentlyClosed.filter((closed) => closed.id !== session.id),
  }
}

export function activateSession(state: DocumentSessionsState, id: string): DocumentSessionsState {
  if (state.activeId === id || !state.sessions.some((session) => session.id === id)) return state
  return { ...state, activeId: id }
}

/**
 * Closes a document. The neighbour to its right becomes active (or the one to
 * its left when it was last), mirroring every tabbed editor.
 */
export function closeSession(state: DocumentSessionsState, id: string): DocumentSessionsState {
  const index = state.sessions.findIndex((session) => session.id === id)
  if (index < 0) return state

  const closed = state.sessions[index]
  const sessions = state.sessions.filter((session) => session.id !== id)
  const activeId =
    state.activeId === id ? (sessions[index] ?? sessions[index - 1] ?? null)?.id ?? null : state.activeId

  return {
    sessions,
    activeId,
    recentlyClosed: [closed, ...state.recentlyClosed.filter((session) => session.id !== id)].slice(
      0,
      MAX_RECENTLY_CLOSED,
    ),
  }
}

/** Ctrl+Shift+T — reopens the last closed document with the bytes it had. */
export function reopenLastClosed(state: DocumentSessionsState): DocumentSessionsState {
  const [last, ...rest] = state.recentlyClosed
  if (!last) return state
  return { sessions: [...state.sessions, last], activeId: last.id, recentlyClosed: rest }
}

/** Ctrl+Tab / Ctrl+Shift+Tab — the next document in tab order, wrapping around. */
export function neighbourSession(state: DocumentSessionsState, delta: 1 | -1): DocumentSession | null {
  if (state.sessions.length < 2) return null
  const index = activeIndex(state)
  if (index < 0) return state.sessions[0] ?? null
  const next = (index + delta + state.sessions.length) % state.sessions.length
  return state.sessions[next] ?? null
}

/** Drag to reorder. */
export function moveSession(state: DocumentSessionsState, fromIndex: number, toIndex: number): DocumentSessionsState {
  if (
    fromIndex === toIndex ||
    fromIndex < 0 ||
    toIndex < 0 ||
    fromIndex >= state.sessions.length ||
    toIndex >= state.sessions.length
  ) {
    return state
  }
  const sessions = [...state.sessions]
  const [moved] = sessions.splice(fromIndex, 1)
  sessions.splice(toIndex, 0, moved)
  return { ...state, sessions }
}

/** After Save As: the document keeps its place in the tab strip under its new path. */
export function renameSession(state: DocumentSessionsState, id: string, file: LoadedFile): DocumentSessionsState {
  const index = state.sessions.findIndex((session) => session.id === id)
  if (index < 0) return state
  const renamed = toSession(file)
  return {
    sessions: state.sessions.map((session, position) => (position === index ? renamed : session)),
    activeId: state.activeId === id ? renamed.id : state.activeId,
    recentlyClosed: state.recentlyClosed,
  }
}
