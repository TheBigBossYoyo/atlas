/**
 * PDF worker setup + packaged-build fake-worker detection (PDF-15/P11).
 *
 * pdfjs-dist tries to load its module worker (`pdf.worker.min.mjs`) off the
 * `workerSrc` URL; if that fails — e.g. a packaged `file://` build serving
 * the worker from an unexpected path — it silently falls back to running
 * entirely on the main thread ("fake worker") via
 * `console.warn("Warning: Setting up fake worker.")`. That fallback still
 * *works*, but it blocks the renderer's main thread during parsing/rendering
 * and defeats the point of a worker, so it's worth surfacing loudly instead
 * of failing silently. pdfjs-dist doesn't expose a public flag for this, so
 * detection works by recognizing its exact warning text.
 */

const FAKE_WORKER_WARNING = 'Warning: Setting up fake worker.'

export function isFakeWorkerWarning(message: unknown): boolean {
  return typeof message === 'string' && message.includes(FAKE_WORKER_WARNING)
}

/**
 * Runs `task` while temporarily wrapping `console.warn` to detect pdfjs-dist's
 * fake-worker fallback warning, restoring the original `console.warn`
 * afterwards regardless of outcome. Returns `task`'s result alongside whether
 * the fallback was detected.
 */
export async function runDetectingFakeWorkerFallback<T>(
  task: () => Promise<T>,
): Promise<{ result: T; usedFakeWorker: boolean }> {
  const originalWarn = console.warn
  let usedFakeWorker = false

  console.warn = (...args: unknown[]) => {
    if (args.some(isFakeWorkerWarning)) {
      usedFakeWorker = true
    }
    originalWarn.apply(console, args)
  }

  try {
    const result = await task()
    return { result, usedFakeWorker }
  } finally {
    console.warn = originalWarn
  }
}
