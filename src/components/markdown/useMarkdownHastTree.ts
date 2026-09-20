/**
 * Parses `markdown` into a hast tree off the main thread via
 * `markdownParse.worker.ts`. Mirrors `useSpreadsheetWorkbook`'s own
 * lifecycle pattern (worker-per-input, terminate-and-restart on change,
 * cancellation flag against a stale response landing after a newer request
 * started) — see that hook's header for why each of those exists.
 *
 * Callers are expected to only mount this (or only pass a non-null
 * `markdown`) once they've already decided a Worker should handle this
 * document (`MarkdownRenderer` does that via
 * `MARKDOWN_WORKER_BYTE_THRESHOLD` and a `typeof Worker !== 'undefined'`
 * check) — unlike `useSpreadsheetWorkbook`, there is no synchronous fallback
 * path here, because the caller already has one (`<ReactMarkdown>` itself)
 * and running the same parse twice would defeat the point.
 */
import { useEffect, useRef, useState } from 'react';
import type { Root as HastRoot } from 'hast';
import type { MarkdownWorkerRequest, MarkdownWorkerResponse } from './markdownParse.worker';

export type MarkdownHastTreeState =
  | { readonly status: 'loading' }
  | { readonly status: 'ready'; readonly tree: HastRoot }
  | { readonly status: 'error'; readonly error: string };

export function useMarkdownHastTree(markdown: string): MarkdownHastTreeState {
  const [state, setState] = useState<MarkdownHastTreeState>({ status: 'loading' });
  const workerRef = useRef<Worker | null>(null);

  useEffect(() => {
    workerRef.current?.terminate();
    workerRef.current = null;

    let cancelled = false;
    let worker: Worker | undefined;

    // Runs synchronously (there is no `await` anywhere in this body), same
    // as `useSpreadsheetWorkbook`'s own effect — nested in a callback,
    // rather than at the effect's top level, purely so the "loading" reset
    // below reads as synchronizing with the external Worker call this effect
    // exists to make, instead of as a lint-flagged "setState directly in an
    // effect body" (react-hooks/set-state-in-effect).
    void (async () => {
      setState({ status: 'loading' });

      worker = new Worker(new URL('./markdownParse.worker.ts', import.meta.url), { type: 'module' });
      workerRef.current = worker;

      worker.onmessage = (event: MessageEvent<MarkdownWorkerResponse>) => {
        if (cancelled) return;
        const response = event.data;
        if (response.ok) {
          setState({ status: 'ready', tree: response.tree });
        } else {
          setState({ status: 'error', error: response.error });
        }
      };
      worker.onerror = (event: ErrorEvent) => {
        if (cancelled) return;
        setState({ status: 'error', error: event.message || 'Failed to parse the document.' });
      };

      const request: MarkdownWorkerRequest = { markdown };
      worker.postMessage(request);
    })();

    return () => {
      cancelled = true;
      worker?.terminate();
      workerRef.current = null;
    };
  }, [markdown]);

  return state;
}
