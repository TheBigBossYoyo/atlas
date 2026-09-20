/**
 * Proves the Worker-backed rendering path (`parseMarkdownToHast` +
 * `renderHastTree`, used by `WorkerMarkdownBody` for large documents) is not
 * a second, silently-diverging markdown pipeline: for every fixture the P0.4
 * characterization suite (`MarkdownRenderer.characterization.test.tsx`)
 * pins, this asserts that path's output is BYTE-IDENTICAL to
 * `<ReactMarkdown>`'s own output for the same markdown and the same
 * `components` map.
 *
 * jsdom (this project's test environment) has no `Worker` global, so
 * `MarkdownRenderer` itself never actually takes the Worker branch in any
 * test — this file is what stands in for that missing coverage, by testing
 * the two rendering mechanisms directly against each other instead of via
 * the size-gated branch. If this suite passes and the characterization
 * suite passes, a document large enough to take the Worker path in a real
 * Electron renderer is guaranteed to render the same way a small one does.
 */
import { describe, expect, it } from 'vitest';
import { cleanup, render } from '@testing-library/react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import rehypeKatex from 'rehype-katex';
import rehypeHighlight from 'rehype-highlight';
import rehypeSlug from 'rehype-slug';
import rehypeRaw from 'rehype-raw';

import { parseMarkdownToHast } from '../markdownPipeline';
import { renderHastTree } from '../renderHastTree';
import { buildMarkdownComponents } from '../markdownComponents';

import headingsFixture from '../../../__tests__/fixtures/markdown/headings.md?raw';
import gfmTableFixture from '../../../__tests__/fixtures/markdown/gfm-table.md?raw';
import taskListFixture from '../../../__tests__/fixtures/markdown/task-list.md?raw';
import strikethroughAutolinksFixture from '../../../__tests__/fixtures/markdown/strikethrough-autolinks.md?raw';
import codeFenceFixture from '../../../__tests__/fixtures/markdown/code-fence.md?raw';
import mathFixture from '../../../__tests__/fixtures/markdown/math.md?raw';
import nestedListsFixture from '../../../__tests__/fixtures/markdown/nested-lists.md?raw';
import blockquoteFixture from '../../../__tests__/fixtures/markdown/blockquote.md?raw';
import rawHtmlFixture from '../../../__tests__/fixtures/markdown/raw-html.md?raw';
import linksImagesFixture from '../../../__tests__/fixtures/markdown/links-images.md?raw';
import kitchenSinkFixture from '../../../__tests__/fixtures/markdown/kitchen-sink.md?raw';

// Mermaid's fixture is deliberately excluded: `Mermaid.tsx` renders
// asynchronously (an effect that calls `mermaid.render`), which is orthogonal
// to which of these two code paths produced the surrounding tree — the
// characterization suite's own mermaid-specific assertions already cover it,
// and diffing pre-effect markup here would just compare two identical
// "mermaid hasn't run yet" placeholders.
const FIXTURES: ReadonlyArray<readonly [name: string, markdown: string]> = [
  ['headings', headingsFixture],
  ['gfm-table', gfmTableFixture],
  ['task-list', taskListFixture],
  ['strikethrough-autolinks', strikethroughAutolinksFixture],
  ['code-fence', codeFenceFixture],
  ['math', mathFixture],
  ['nested-lists', nestedListsFixture],
  ['blockquote', blockquoteFixture],
  ['raw-html', rawHtmlFixture],
  ['links-images', linksImagesFixture],
  ['kitchen-sink', kitchenSinkFixture],
];

const remarkPlugins = [remarkGfm, remarkMath];
const rehypePlugins = [rehypeKatex, rehypeHighlight, rehypeSlug, rehypeRaw];

function renderReactMarkdown(markdown: string, components: ReturnType<typeof buildMarkdownComponents>): string {
  const { container } = render(
    <ReactMarkdown remarkPlugins={remarkPlugins} rehypePlugins={rehypePlugins} components={components}>
      {markdown}
    </ReactMarkdown>,
  );
  const html = container.innerHTML;
  cleanup();
  return html;
}

function renderViaHastTree(markdown: string, components: ReturnType<typeof buildMarkdownComponents>): string {
  const tree = parseMarkdownToHast(markdown);
  const { container } = render(renderHastTree(tree, components));
  const html = container.innerHTML;
  cleanup();
  return html;
}

describe('Worker-path parity: renderHastTree(parseMarkdownToHast(x)) === <ReactMarkdown>{x}</ReactMarkdown>', () => {
  for (const [name, markdown] of FIXTURES) {
    it(`renders the "${name}" fixture identically to <ReactMarkdown>`, () => {
      // A fixed baseId (not React's real useId()) so mermaid-hash strings —
      // absent from every fixture here anyway — can't introduce a spurious
      // difference; kept only for signature parity with MarkdownRenderer's
      // own call.
      const components = buildMarkdownComponents(undefined, 'parity-test');

      const expected = renderReactMarkdown(markdown, components);
      const actual = renderViaHastTree(markdown, components);

      expect(actual).toBe(expected);
    });

    it(`renders the "${name}" fixture identically to <ReactMarkdown> with an active search query`, () => {
      const components = buildMarkdownComponents('the', 'parity-test');

      const expected = renderReactMarkdown(markdown, components);
      const actual = renderViaHastTree(markdown, components);

      expect(actual).toBe(expected);
    });
  }
});
