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

/**
 * Ctrl+Shift+T's other half: the path to reopen, and the state with that entry
 * popped. A closed session keeps no bytes (see `shedBytes`), so the shell
 * re-opens it through the normal load path — which re-reads the file, reports
 * a file that has since been moved or deleted instead of showing an empty
 * document, and re-registers it as a tab on success.
 */
export function takeLastClosedPath(
  state: DocumentSessionsState,
): { readonly state: DocumentSessionsState; readonly path: string } | null {
  const [last, ...rest] = state.recentlyClosed
  if (!last) return null
  return { state: { ...state, recentlyClosed: rest }, path: last.file.path }
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

/**
 * After Save As: the document keeps its place in the tab strip under its new
 * path.
 *
 * SESS-1 — a Save As can land on a path that is *already* open as a
 * different session (save untitled/renamed document A onto the path of
 * already-open document B). `id` is every other function's notion of
 * identity (`sessions.find`, `activateSession`, `closeSessionById`), so
 * silently producing two sessions that share an id would corrupt that
 * invariant outright — `find` would return whichever collides with it first,
 * `closeSessionById` could close the wrong tab, etc. This never showed up in
 * `openSession`, which already de-dupes by construction (opening a path
 * either creates or re-activates the one session for it); `renameSession`
 * had no equivalent check.
 *
 * Chosen semantics, deliberately conservative:
 *
 * 1. Ordinarily, the session being renamed (A) just had its bytes physically
 *    written to that path on disk — it is now the authoritative content for
 *    that path. The colliding session (B) is therefore dropped from the tab
 *    strip: keeping it would show B's in-memory copy as if it still reflected
 *    that path, which is no longer true, and nothing is actually lost — a
 *    fresh open of that path would show exactly A's bytes anyway (every
 *    reactivation already re-reads from disk via `reloadSessionFile`). A
 *    still ends up at the id/path it was renamed to, keeping A's original
 *    tab position; only B's tab disappears.
 *
 * 2. The one exception: if B is the session currently on screen
 *    (`state.activeId`) and A is a *different*, background session, we do
 *    NOT drop B or hand its identity to A. This module has no way to also
 *    swap what's rendered — `App.tsx`'s `file`/`filePath` state lives outside
 *    it, and its callers for a save that finished after the user switched
 *    away (`applySavedPathToInactiveTab`, `handleViewerSavedPath`'s
 *    background branch) deliberately do NOT call `adoptFile`/`showSessionFile`
 *    in that case, specifically so a stale async save can never yank the
 *    view out from under whatever the user is now looking at (which may
 *    itself be dirty). Silently repointing `activeId` at A here, with
 *    nothing to render it, would desync the highlighted tab from the actual
 *    content on screen — worse than the alternative. So instead A (the
 *    background rename) is the one dropped: its bytes are already safely on
 *    disk under that path, the user isn't looking at its tab, so nothing
 *    visible is lost — only its now-redundant tab goes away, since B already
 *    represents that path as far as the screen is concerned.
 */
export function renameSession(state: DocumentSessionsState, id: string, file: LoadedFile): DocumentSessionsState {
  const target = state.sessions.find((session) => session.id === id)
  if (!target) return state
  const renamed = toSession(file)

  const collision = state.sessions.find((session) => session.id === renamed.id && session.id !== id)

  if (collision && collision.id === state.activeId && state.activeId !== id) {
    // Case 2 above: the colliding tab is the one on screen right now — leave
    // it (and `activeId`) untouched, drop the just-renamed background
    // session instead.
    return {
      sessions: state.sessions.filter((session) => session.id !== id),
      activeId: state.activeId,
      recentlyClosed: state.recentlyClosed,
    }
  }

  // Case 1 (the common case: no collision at all, or a collision with some
  // other background tab) — drop the collision (if any) and rename `id` in
  // place.
  const sessions = state.sessions
    .filter((session) => session.id !== renamed.id || session.id === id)
    .map((session) => (session.id === id ? renamed : session))

  return {
    sessions,
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
