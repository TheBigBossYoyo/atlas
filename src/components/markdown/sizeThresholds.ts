/**
 * Above this many characters, `MarkdownRenderer` parses off the main thread
 * via `markdownParse.worker.ts` instead of handing the string straight to
 * `<ReactMarkdown>`. Below it, an inline parse is fast enough (and a
 * Worker's own startup latency high enough relatively) that going off-thread
 * would cost more than it saves.
 *
 * Deliberately well under `LARGE_MARKDOWN_PREVIEW_BYTES`
 * (`LargeMarkdownNotice.tsx`): a Worker is what makes documents in that gap
 * — big enough that a synchronous parse would show up, nowhere near big
 * enough to need an opt-in warning — render automatically, with no
 * frozen frame and no click required. Chosen from this project's own
 * corpus measurements (see `markdownParse.worker.ts`'s header): 0.1 MB
 * parses in well under a frame's worth of main-thread time inline, so
 * nothing below this line needs to move off-thread at all.
 */
export const MARKDOWN_WORKER_BYTE_THRESHOLD = 100_000;
