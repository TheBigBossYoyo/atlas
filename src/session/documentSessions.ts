/**
 * SHELL-17 — open documents ("tabs").
 *
 * A pure, immutable model of which files are open, which one is showing, and
 * which ones were closed recently (so Ctrl+Shift+T can bring one back). The
 * shell owns the side effects: prompting about unsaved changes before it
 * switches away from a document, and re-reading the newly-activated document
 * from disk (`reloadSessionFile`) so it never shows bytes older than its last
 * save.
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
 * Strips a document down to an empty placeholder of the same kind, keeping
 * its path/format/name identity. `recentlyClosed` entries only ever exist so
 * Ctrl+Shift+T can bring a document back — and bringing it back always
 * re-reads it from disk via `reloadSessionFile` first (SHELL-17), which
 * ignores whatever bytes a session was holding. Keeping the FULL original
 * bytes alive in `recentlyClosed` (up to `MAX_RECENTLY_CLOSED` of them) was
 * therefore pure waste for a closed binary document: confirmed by a heap
 * snapshot after closing a 100k-row .xlsx tab, where the closed session's
 * retained `ArrayBuffer` was the largest single object kept alive by the
 * whole renderer.
 *
 * The only time these placeholder bytes are ever actually shown is if that
 * re-read itself fails (the file was moved/deleted while its tab was
 * closed) — previously that fell back to the last bytes the tab had, and
 * now it shows an empty document instead. Accepted trade-off: reopening a
 * file is expected to reflect what's on disk, not resurrect a memory of
 * bytes that may no longer correspond to anything.
 */
function shedBytes(file: LoadedFile): LoadedFile {
  return file.kind === 'text' ? { ...file, content: '' } : { ...file, content: new ArrayBuffer(0) }
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

  // Only the path/name/format need to survive into `recentlyClosed` — see
  // `shedBytes`'s own comment for why the bytes themselves don't.
  const shed: DocumentSession = { ...closed, file: shedBytes(closed.file) }

  return {
    sessions,
    activeId,
    recentlyClosed: [shed, ...state.recentlyClosed.filter((session) => session.id !== id)].slice(
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

/** The two reads `reloadSessionFile` needs — a subset of `window.electronAPI`. */
export type SessionFileReader = {
  readonly openFileByPath: (path: string) => Promise<{ content: string } | null>
  readonly readBinaryByPath: (path: string) => Promise<{ buffer: ArrayBuffer }>
}

/**
 * The current contents of a document being shown again (tab switch, close,
 * Ctrl+Shift+T). A session keeps the bytes it was opened with, but every save
 * since then went to disk only, so showing the kept bytes would display the
 * pre-save version, and the next save would overwrite the saved work with
 * edits made on top of that stale copy. Falls back to the kept bytes when the
 * file can no longer be read (moved, deleted, no Electron bridge).
 */
export async function reloadSessionFile(file: LoadedFile, reader: SessionFileReader | undefined): Promise<LoadedFile> {
  if (!reader) return file
  try {
    if (file.kind === 'text') {
      const data = await reader.openFileByPath(file.path)
      return data ? { ...file, content: data.content } : file
    }
    const { buffer } = await reader.readBinaryByPath(file.path)
    return { ...file, content: buffer }
  } catch {
    return file
  }
}
