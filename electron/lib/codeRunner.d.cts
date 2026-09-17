/** Types for `codeRunner.cjs` (USR-19). */
export type CodeRunCandidate = {
  readonly command: string
  readonly args: string[]
  readonly env?: Record<string, string>
}

export type CodeRuntime = {
  readonly language: string
  readonly candidates: CodeRunCandidate[]
}

export type CodeRunStartResult = { ok: true; runId: number } | { ok: false; error: string }

export type CodeRunner = {
  start(filePath: string): CodeRunStartResult
  stop(runId: number): boolean
  dispose(): void
}

export declare function runtimeFor(filePath: string): CodeRuntime | null
export declare function childEnv(extra?: Record<string, string>): Record<string, string>
export declare function createCodeRunner(options: {
  send: (channel: string, payload: Record<string, unknown>) => void
  spawnImpl?: unknown
  timeoutMs?: number
}): CodeRunner
export declare const MAX_OUTPUT_BYTES: number
