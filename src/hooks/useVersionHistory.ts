/**
 * VERSIONS-1 — the renderer's view of a document's version history.
 *
 * Versions themselves are recorded in the main process, at the moment a document
 * is written (see `recordVersionAfterWrite` in `electron/main.cjs`), so this hook
 * does not create them for saves — it reads the list, fetches one version's bytes
 * on request, and offers the two things the renderer IS in a position to do:
 *
 *   - `captureNow`, for the "save a version of this right now" action and for the
 *     periodic capture of a document being typed into, where what matters is the
 *     state in front of the user rather than the last thing written to disk;
 *   - `refresh`, because a save happens through a path this hook cannot see.
 *
 * Desktop-only: in a plain browser tab there is no main process to hold a
 * history, so `available` is false and the UI hides rather than offering
 * something that cannot work.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

import type { DocumentVersion } from '../electron'

export type VersionHistoryState = {
  /** Whether a history can exist at all — false in a plain browser tab. */
  readonly available: boolean
  /** Newest FIRST, which is the order the panel shows them in. */
  readonly versions: ReadonlyArray<DocumentVersion>
  readonly loading: boolean
  readonly error: string | null
  /** Re-reads the list; call after anything that may have written a version. */
  readonly refresh: () => Promise<void>
  /** Records the bytes the user currently has, whether or not they are on disk. */
  readonly captureNow: (bytes: Uint8Array, label?: string) => Promise<boolean>
  /** One version's bytes, or null if it is gone. */
  readonly readVersion: (id: string) => Promise<Uint8Array | null>
  readonly clear: () => Promise<void>
}

export function useVersionHistory(filePath: string | null): VersionHistoryState {
  const api = typeof window === 'undefined' ? undefined : window.electronAPI
  const available = Boolean(filePath) && typeof api?.historyList === 'function'

  const [versions, setVersions] = useState<ReadonlyArray<DocumentVersion>>([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  // Guards against a slow list for document A landing after the user has already
  // switched to document B and overwriting B's list with A's.
  const requestRef = useRef(0)

  const refresh = useCallback(async () => {
    if (!filePath || typeof api?.historyList !== 'function') {
      setVersions([])
      return
    }
    const request = ++requestRef.current
    setLoading(true)
    try {
      const result = await api.historyList(filePath)
      if (request !== requestRef.current) return
      // Newest first for display; the store keeps them oldest first.
      setVersions([...result.versions].reverse())
      setError(result.error ?? null)
    } catch (err) {
      if (request !== requestRef.current) return
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      if (request === requestRef.current) setLoading(false)
    }
  }, [api, filePath])

  useEffect(() => {
    void refresh()
  }, [refresh])

  const captureNow = useCallback(
    async (bytes: Uint8Array, label?: string) => {
      if (!filePath || typeof api?.historySnapshot !== 'function') return false
      try {
        const result = await api.historySnapshot({ path: filePath, bytes, label: label ?? null })
        // An unchanged document is not a failure and not a new version either, so
        // there is nothing to re-read.
        if (result.stored) await refresh()
        return result.stored
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err))
        return false
      }
    },
    [api, filePath, refresh],
  )

  const readVersion = useCallback(
    async (id: string) => {
      if (!filePath || typeof api?.historyRead !== 'function') return null
      try {
        const result = await api.historyRead({ path: filePath, id })
        if (result.error !== undefined || result.bytes === undefined) {
          setError(result.error ?? null)
          return null
        }
        return result.bytes
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err))
        return null
      }
    },
    [api, filePath],
  )

  const clear = useCallback(async () => {
    if (!filePath || typeof api?.historyClear !== 'function') return
    await api.historyClear(filePath)
    await refresh()
  }, [api, filePath, refresh])

  return useMemo(
    () => ({ available, versions, loading, error, refresh, captureNow, readVersion, clear }),
    [available, versions, loading, error, refresh, captureNow, readVersion, clear],
  )
}

/**
 * How often a document being actively typed into is captured.
 *
 * VERSIONS-1 — the request was to see a writing PROCESS, and a history made only
 * of saves shows the process of someone who remembers to press Ctrl+S. Two
 * minutes is frequent enough to show how a paragraph came together and slow
 * enough that the store's dedup (identical content is refused) keeps an idle
 * document from accumulating anything at all.
 */
export const AUTO_CAPTURE_INTERVAL_MS = 2 * 60 * 1000

/**
 * Captures the document periodically while it differs from what was last
 * recorded.
 *
 * `getBytes` is called only when a capture is actually due, never on render: for
 * a large document, serialising it is the expensive part and doing it on a timer
 * that usually has nothing to do would be a tax on every open document.
 */
export function useAutoCapture(
  history: VersionHistoryState,
  enabled: boolean,
  getBytes: () => Uint8Array | null,
): void {
  const getBytesRef = useRef(getBytes)
  useEffect(() => {
    getBytesRef.current = getBytes
  }, [getBytes])

  const captureRef = useRef(history.captureNow)
  useEffect(() => {
    captureRef.current = history.captureNow
  }, [history.captureNow])

  useEffect(() => {
    if (!enabled || !history.available) return undefined
    const timer = setInterval(() => {
      const bytes = getBytesRef.current()
      // The store refuses content identical to the newest version, so an idle
      // document costs one comparison here and nothing on disk.
      if (bytes !== null && bytes.length > 0) void captureRef.current(bytes)
    }, AUTO_CAPTURE_INTERVAL_MS)
    return () => clearInterval(timer)
  }, [enabled, history.available])
}
