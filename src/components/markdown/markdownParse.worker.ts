/**
 * Worker entry for off-main-thread markdown parsing.
 *
 * Markdown parsing (micromark, under remark-parse) degrades superlinearly on
 * this project's own generated corpus: 0.5 MB took ~2 s, 1 MB ~9 s, 2 MB
 * ~44 s, and a 5 MB document took over two minutes at a ~2.6 GB peak — all
 * of it, previously, on the main thread (see `LargeMarkdownNotice.tsx`'s
 * header for the full measurement). None of that cost goes away by moving
 * it here — a Worker doesn't make micromark faster, and a document large
 * enough can still take minutes or exhaust this Worker's own heap — but it
 * does mean that cost no longer blocks paint, input, or any other tab: this
 * Worker gets its own V8 isolate, so even an out-of-memory crash here ends
 * this parse (reported back as an error) rather than the renderer process.
 *
 * Mirrors `spreadsheetWorker.worker.ts`'s own pattern: this file is only the
 * postMessage plumbing, so the actual pipeline (`markdownPipeline.ts`) stays
 * plain, synchronous, side-effect-free TS that's also used directly by
 * `MarkdownRenderer`'s small-document path's tests.
 */
import { parseMarkdownToHast } from './markdownPipeline';
import type { Root as HastRoot } from 'hast';

export type MarkdownWorkerRequest = {
  readonly markdown: string;
};

export type MarkdownWorkerResponse =
  | { readonly ok: true; readonly tree: HastRoot }
  | { readonly ok: false; readonly error: string };

self.onmessage = (event: MessageEvent<MarkdownWorkerRequest>) => {
  try {
    const tree = parseMarkdownToHast(event.data.markdown);
    const response: MarkdownWorkerResponse = { ok: true, tree };
    self.postMessage(response);
  } catch (err) {
    const response: MarkdownWorkerResponse = {
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    };
    self.postMessage(response);
  }
};
