/**
 * The remark/rehype pipeline `MarkdownRenderer` renders with, factored out
 * so it can run identically in two places:
 *
 *  - inline, on the main thread, for documents small enough that parsing
 *    them synchronously is not noticeable (`MarkdownRenderer`'s own
 *    `<ReactMarkdown>` branch, unchanged from before this file existed); and
 *  - inside `markdownParse.worker.ts`, off the main thread, for documents
 *    large enough that micromark's superlinear parse cost would otherwise
 *    freeze the UI (see that worker's header, and `LargeMarkdownNotice.tsx`,
 *    for the measurements behind that threshold).
 *
 * `createMarkdownProcessor` mirrors react-markdown's own (unexported)
 * `createProcessor` in `node_modules/react-markdown/lib/index.js`'s
 * structure (remark plugins, then `remarkRehype`, then rehype plugins), and
 * mirrors this project's OWN prior `<ReactMarkdown remarkPlugins rehypePlugins>`
 * call — including its exact rehype plugin order, `[rehypeKatex,
 * rehypeHighlight, rehypeSlug, rehypeRaw]` (rehypeRaw deliberately LAST, not
 * first) — byte-for-byte, since that order is exactly what the P0.4
 * characterization suite's snapshots were frozen against; reordering it
 * (e.g. to the more common raw-first order) would change output for any
 * fixture containing raw HTML. And
 * `parseMarkdownToHast` mirrors its `Markdown()` component's
 * `processor.runSync(processor.parse(file), file)` call. That parity is
 * deliberate: it is what lets a document parsed here, off-thread, render
 * through `renderHastTree.tsx` to byte-identical output as one parsed
 * in-place by `<ReactMarkdown>` — see
 * `src/components/markdown/__tests__/renderHastTree.parity.test.tsx`, which
 * checks exactly that against the same fixtures the P0.4 characterization
 * suite uses.
 */
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import remarkParse from 'remark-parse';
import remarkRehype from 'remark-rehype';
import rehypeKatex from 'rehype-katex';
import rehypeHighlight from 'rehype-highlight';
import rehypeSlug from 'rehype-slug';
import rehypeRaw from 'rehype-raw';
import { unified } from 'unified';
import { VFile } from 'vfile';
import type { Root as HastRoot } from 'hast';

export function createMarkdownProcessor() {
  return unified()
    .use(remarkParse)
    .use(remarkGfm)
    .use(remarkMath)
    .use(remarkRehype, { allowDangerousHtml: true })
    .use(rehypeKatex)
    .use(rehypeHighlight)
    .use(rehypeSlug)
    .use(rehypeRaw);
}

/**
 * Parses `markdown` all the way to its final hast tree — every remark/rehype
 * plugin above has run — synchronously. Every plugin in this pipeline is
 * itself synchronous (no plugin here returns a promise from its
 * transformer), so `runSync` is safe: there is no async work this function
 * would silently fail to wait for.
 */
export function parseMarkdownToHast(markdown: string): HastRoot {
  const processor = createMarkdownProcessor();
  const file = new VFile();
  file.value = markdown;
  const tree = processor.runSync(processor.parse(file), file);
  return tree as HastRoot;
}
