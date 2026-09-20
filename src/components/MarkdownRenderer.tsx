import { forwardRef, lazy, Suspense, useMemo, useId } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import rehypeKatex from 'rehype-katex';
import rehypeHighlight from 'rehype-highlight';
import rehypeSlug from 'rehype-slug';
import rehypeRaw from 'rehype-raw';
import { MARKDOWN_WORKER_BYTE_THRESHOLD } from './markdown/sizeThresholds';
import { buildMarkdownComponents } from './markdown/markdownComponents';

interface MarkdownRendererProps {
  markdown: string;
  searchQuery?: string;
}

// PERF — `WorkerMarkdownBody` (and everything it pulls in: `renderHastTree`,
// `hast-util-to-jsx-runtime`, `html-url-attributes`, `unist-util-visit`,
// `useMarkdownHastTree`) is real weight that only the Worker branch below
// ever needs. `lazy()`-loading it, the same way `App.tsx` already
// `lazy()`-loads this whole module out of the entry chunk (see that file's
// own PERF-01 comment), keeps every small document's import graph exactly
// as light as it was before this file gained a Worker path at all — an
// eager import here measurably slowed `MarkdownRenderer`'s own dynamic
// import (caught by `App.dirtyState.characterization.test.tsx`'s
// `waitFor(...)` timing out under the *default* 1s `waitFor` budget; it is
// not this suite's job to grow a longer timeout to accommodate a heavier
// common path).
const WorkerMarkdownBody = lazy(() =>
  import('./markdown/WorkerMarkdownBody').then((m) => ({ default: m.WorkerMarkdownBody }))
);

export const MarkdownRenderer = forwardRef<HTMLDivElement, MarkdownRendererProps>(
  function MarkdownRenderer({ markdown, searchQuery }, ref) {
    const baseId = useId();
    const remarkPlugins = useMemo(() => [remarkGfm, remarkMath], []);
    const rehypePlugins = useMemo(() => [rehypeKatex, rehypeHighlight, rehypeSlug, rehypeRaw], []);

    const components = useMemo(
      () => buildMarkdownComponents(searchQuery, baseId),
      [searchQuery, baseId],
    );

    // T2-equivalent for markdown (see markdownParse.worker.ts's header):
    // above this size, parsing off the main thread is what keeps the app
    // responsive instead of freezing on micromark's superlinear cost.
    // `typeof Worker !== 'undefined'` mirrors useSpreadsheetWorkbook's own
    // guard — jsdom (this project's test environment) has no `Worker`
    // global, so tests below the threshold exercise this exact branch, and
    // every characterization fixture is far below it, so none of them ever
    // take the Worker path (see renderHastTree.parity.test.tsx for proof
    // that branch renders identically when it does run).
    if (markdown.length >= MARKDOWN_WORKER_BYTE_THRESHOLD && typeof Worker !== 'undefined') {
      return (
        <Suspense fallback={<div ref={ref} className="markdown-body" id="markdown-content" />}>
          <WorkerMarkdownBody ref={ref} markdown={markdown} components={components} />
        </Suspense>
      );
    }

    return (
      <div ref={ref} className="markdown-body" id="markdown-content">
        <ReactMarkdown
          remarkPlugins={remarkPlugins}
          rehypePlugins={rehypePlugins}
          components={components}
        >
          {markdown}
        </ReactMarkdown>
      </div>
    );
  }
);
