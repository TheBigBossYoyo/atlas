/**
 * USR-19 — renderer side of "Run": asks the main process to run the file on
 * disk (which asks the user for confirmation in a native dialog), then
 * streams the program's output into the panel. The renderer never sends code
 * to be executed — only the path of the file the user already opened.
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import { t } from '../../i18n'

export type RunChunk = { readonly stream: 'stdout' | 'stderr' | 'system'; readonly text: string }

export type RunState =
  | { readonly status: 'idle' }
  | { readonly status: 'running'; readonly runId: number }
  | { readonly status: 'finished'; readonly summary: string; readonly failed: boolean }

/** Extensions `electron/lib/codeRunner.cjs` knows how to run. */
const RUNNABLE_EXTENSIONS: ReadonlySet<string> = new Set(['js', 'mjs', 'cjs', 'ts', 'mts', 'cts', 'py', 'pyw'])

export function isRunnable(filePath: string): boolean {
  const match = /\.([^./\\]+)$/.exec(filePath)
  return match !== null && RUNNABLE_EXTENSIONS.has(match[1].toLowerCase())
}

function exitSummary(payload: {
  code: number | null
  timedOut: boolean
  stopped: boolean
  error?: string
}): { summary: string; failed: boolean } {
  if (payload.error) return { summary: payload.error, failed: true }
  if (payload.stopped) return { summary: t('codeRun.stopped'), failed: false }
  if (payload.timedOut) return { summary: t('codeRun.timedOut'), failed: true }
  if (payload.code === 0) return { summary: t('codeRun.finishedOk'), failed: false }
  return { summary: t('codeRun.exitedWithCode', { code: payload.code ?? 'unknown' }), failed: true }
}

export function useCodeRun(filePath: string) {
  const [state, setState] = useState<RunState>({ status: 'idle' })
  const [chunks, setChunks] = useState<ReadonlyArray<RunChunk>>([])
  const [isOpen, setIsOpen] = useState(false)
  const runIdRef = useRef<number | null>(null)

  useEffect(() => {
    const api = window.electronAPI?.codeRun
    if (!api) return undefined
    const offOutput = api.onOutput((payload) => {
      if (runIdRef.current !== null && payload.runId !== runIdRef.current) return
      setChunks((current) => [...current, { stream: payload.stream, text: payload.text }])
    })
    const offExit = api.onExit((payload) => {
      if (runIdRef.current !== null && payload.runId !== runIdRef.current) return
      runIdRef.current = null
      setState({ status: 'finished', ...exitSummary(payload) })
    })
    return () => {
      offOutput()
      offExit()
    }
  }, [])

  // A different file means a different program: drop the previous output.
  // Render-time state adjustment (the same idiom SpreadsheetViewer uses for
  // its active-sheet reset) rather than an effect, so the panel never shows
  // one file's output under another file's name for a frame.
  const [trackedPath, setTrackedPath] = useState(filePath)
  if (trackedPath !== filePath) {
    setTrackedPath(filePath)
    setChunks([])
    setState({ status: 'idle' })
  }

  const run = useCallback(async (): Promise<void> => {
    const api = window.electronAPI?.codeRun
    if (!api) return
    setIsOpen(true)
    setChunks([])
    const result = await api.start(filePath)
    if (!result?.ok || result.runId === undefined) {
      runIdRef.current = null
      if (result?.cancelled) {
        setState({ status: 'idle' })
        setIsOpen(false)
        return
      }
      setState({ status: 'finished', summary: result?.error ?? t('codeRun.couldNotStart'), failed: true })
      return
    }
    runIdRef.current = result.runId
    setState({ status: 'running', runId: result.runId })
  }, [filePath])

  const stop = useCallback((): void => {
    const runId = runIdRef.current
    if (runId !== null) void window.electronAPI?.codeRun?.stop(runId)
  }, [])

  return {
    state,
    chunks,
    isOpen,
    isAvailable: typeof window !== 'undefined' && window.electronAPI?.codeRun !== undefined && isRunnable(filePath),
    open: () => setIsOpen(true),
    close: () => setIsOpen(false),
    clear: () => setChunks([]),
    run,
    stop,
  }
}
