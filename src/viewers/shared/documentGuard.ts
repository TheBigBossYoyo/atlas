/**
 * Size cap + render timeout guard for RTF/ODT conversion (T9/DAT-20).
 *
 * Both `rtf.js` and `odf-kit` run a single, uninterruptible conversion pass
 * over the whole document with no built-in limit — a pathological input
 * (a maliciously or accidentally deeply-nested RTF group, a multi-hundred-MB
 * ODT) can otherwise hang the viewer indefinitely with no feedback.
 *
 * Two independent mitigations, matching the two independent failure modes:
 *  - `DOCUMENT_SIZE_CAP_BYTES` rejects an oversized file up front, before any
 *    conversion is attempted at all — this is the real guard against a huge
 *    file, and it is instant and total (`byteLength` is a synchronous check
 *    that costs nothing to run before "spawning" the risk).
 *  - `withRenderTimeout` races the conversion against a wall-clock timeout.
 *    Honest limitation: JavaScript is single-threaded, so a timer can only
 *    fire once the current synchronous call stack yields; if `render()`/
 *    `readOdt()` never awaits internally, the timeout cannot preempt it
 *    mid-loop. It still bounds the worst case for the parts of the pipeline
 *    that genuinely are async (the dynamic `import()` of the conversion
 *    library, `rtf.js`'s awaited WMF/EMF image decoding) and turns an
 *    otherwise-silent indefinite "Loading…" into a friendly error after a
 *    fixed wait. A real preemptive guard would require moving the entire
 *    conversion into a cancellable Worker; both libraries hand back live DOM
 *    nodes / call into `dompurify`+`innerHTML` on the caller's document,
 *    which isn't something a Worker can produce, so that rewrite is out of
 *    scope here (tracked as a known limitation, not silently pretended away).
 */

/** Above this size, RTF/ODT conversion is refused outright rather than attempted. */
export const DOCUMENT_SIZE_CAP_BYTES = 25 * 1024 * 1024 // 25 MB

/** Wall-clock budget for a single RTF/ODT conversion before it is treated as hung. */
export const DOCUMENT_RENDER_TIMEOUT_MS = 15_000

/** Thrown by `withRenderTimeout` when the budget is exceeded. */
export class DocumentRenderTimeoutError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'DocumentRenderTimeoutError'
  }
}

/**
 * Formats the friendly "file too large to preview" message shared by
 * RtfViewer/OdtViewer, given the file's actual byte size.
 */
export function formatSizeCapMessage(byteLength: number, label: string): string {
  const mb = (byteLength / (1024 * 1024)).toFixed(1)
  const capMb = (DOCUMENT_SIZE_CAP_BYTES / (1024 * 1024)).toFixed(0)
  return `This file is too large to preview: ${mb} MB, over the ${capMb} MB ${label.toUpperCase()} preview limit. Open it in a dedicated ${label.toUpperCase()} editor instead.`
}

/**
 * Races `promise` against a `ms`-millisecond timer. Rejects with a
 * `DocumentRenderTimeoutError` carrying `timeoutMessage` if the timer wins;
 * otherwise settles exactly as `promise` does. The timer is always cleared,
 * so a fast `promise` never leaves a dangling timeout behind.
 */
export function withRenderTimeout<T>(promise: Promise<T>, ms: number, timeoutMessage: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new DocumentRenderTimeoutError(timeoutMessage))
    }, ms)

    promise.then(
      (value) => {
        clearTimeout(timer)
        resolve(value)
      },
      (err: unknown) => {
        clearTimeout(timer)
        reject(err instanceof Error ? err : new Error(String(err)))
      },
    )
  })
}
