/**
 * Renderer-side last-resort error reporting.
 *
 * The main process has had `process.on('uncaughtException')` and
 * `('unhandledRejection')` plus a crash log since P1.15 (`electron/main.cjs`,
 * `electron/lib/crashLog.cjs`), and `render-process-gone` handling for when the
 * whole renderer dies. The renderer itself had neither: `ViewerErrorBoundary`
 * catches errors thrown during React *rendering*, but a promise rejected in an
 * async path outside its own try/catch — a save, an export, a parse kicked off
 * from an event handler — resolved into nothing at all. No toast, no log line,
 * no clue for the user beyond an action that appeared to do nothing. Since the
 * renderer is where every parse, edit and export actually happens, that was the
 * larger of the two gaps.
 *
 * This is deliberately a backstop, not an error-handling strategy: the paths
 * that matter already catch their own failures and surface a specific, useful
 * message (`utils/export/*` toast the export that failed, the save handlers
 * return `{ saved: false, error }`). Anything reaching here is by definition a
 * case nobody anticipated, so it gets a generic message — the value is that the
 * user learns something went wrong instead of silently distrusting the app, and
 * that the details land in the console where Electron's renderer log keeps them.
 */
import { useEffect } from 'react';

import type { ToastVariant } from '../components/ToastContext';

/** How long an identical message is suppressed for, so one failing loop can't bury the UI in toasts. */
const DEDUPE_WINDOW_MS = 5000;

/** Hard cap on toasts from this reporter per session — past this, the console still gets everything. */
const MAX_REPORTS = 10;

function messageOf(reason: unknown): string {
  if (reason instanceof Error) return reason.message;
  if (typeof reason === 'string') return reason;
  // A rejection can carry anything at all (`Promise.reject({ code: 500 })`).
  try {
    return JSON.stringify(reason) ?? String(reason);
  } catch {
    return String(reason);
  }
}

/**
 * Whether an `error` event is a failed subresource load rather than a thrown
 * error. A broken `<img>`/`<link>` fires `error` ON THE ELEMENT and it bubbles
 * to window with no `error` property — and markdown regularly references remote
 * images (README badges, hotlinked screenshots, which the production CSP
 * deliberately still allows: see `electron/lib/csp.cjs`), so without this check
 * every dead image URL in an opened document would raise a scary toast about
 * something the user cannot act on and Atlas did not get wrong.
 */
function isResourceLoadFailure(event: ErrorEvent): boolean {
  return event.target !== null && event.target !== window && !(event.error instanceof Error);
}

/**
 * Installs window-level `unhandledrejection`/`error` listeners that log the
 * failure and show one generic toast. Call once, from a component inside
 * `<ToastProvider>`; `showToast` and `describe` are passed in rather than read
 * from context here so this stays a plain hook that a test can drive directly.
 *
 * @param showToast - the app's toast queue (from `useToast()`).
 * @param describe  - formats the user-facing message, given the raw detail
 *                    text; the caller supplies it so the string comes from the
 *                    i18n catalogue rather than being hard-coded here.
 */
export function useUnhandledErrorReporter(
  showToast: (message: string, variant?: ToastVariant) => void,
  describe: (detail: string) => string,
): void {
  useEffect(() => {
    let reports = 0;
    const recent = new Map<string, number>();

    const report = (detail: string, logLabel: string, raw: unknown): void => {
      // Always log, even past the cap or inside the dedupe window: the console
      // is the diagnostic record, the toast is only the user-facing hint.
      console.error(`[atlas] ${logLabel}:`, raw);

      const now = Date.now();
      const lastSeen = recent.get(detail);
      if (lastSeen !== undefined && now - lastSeen < DEDUPE_WINDOW_MS) return;
      recent.set(detail, now);
      for (const [key, at] of recent) {
        if (now - at >= DEDUPE_WINDOW_MS) recent.delete(key);
      }

      if (reports >= MAX_REPORTS) return;
      reports += 1;
      showToast(describe(detail), 'error');
    };

    const onRejection = (event: PromiseRejectionEvent): void => {
      report(messageOf(event.reason), 'unhandled rejection', event.reason);
    };

    const onError = (event: ErrorEvent): void => {
      if (isResourceLoadFailure(event)) return;
      report(messageOf(event.error ?? event.message), 'uncaught error', event.error ?? event.message);
    };

    window.addEventListener('unhandledrejection', onRejection);
    // Capture phase: a subresource `error` event does not bubble in every
    // browser, and this listener needs to see (and ignore) those rather than
    // miss the ones that do bubble.
    window.addEventListener('error', onError, true);

    return () => {
      window.removeEventListener('unhandledrejection', onRejection);
      window.removeEventListener('error', onError, true);
    };
  }, [describe, showToast]);
}
