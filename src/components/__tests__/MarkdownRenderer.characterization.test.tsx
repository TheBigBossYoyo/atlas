/**
 * MarkdownRenderer — characterization suite (P0.4 / QA-07).
 *
 * Freezes today's exact rendered HTML for the current markdown pipeline
 * (react-markdown + remark-gfm/remark-math + rehype-katex/rehype-highlight/
 * rehype-slug/rehype-raw). Per the improvement plan's Section 7 "shared-code
 * change rule", these snapshots must stay byte-identical across every later
 * wave unless a markdown-rendering change is explicitly reviewed and
 * approved by the owner — a snapshot diff on a PR that isn't supposed to
 * touch markdown rendering is a red flag, not something to `-u` away.
 *
 * CSS blind spot: jsdom has no layout/paint engine and this project's
 * vitest.config.ts runs with `css: false` (stylesheet imports are stubbed to
 * nothing), so every snapshot below freezes DOM *structure* only — tags,
 * classes, attributes, text content. A purely visual regression (wrong
 * color, broken layout, a CSS variable typo) will NOT be caught here. The
 * `getComputedStyle` probes near the bottom are a cheap, partial mitigation
 * for *structural* CSS regressions (see their own comment); real pixel-diff
 * visual regression testing is out of scope for this plan (DEFER-8).
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { render, waitFor } from '@testing-library/react';
import { MarkdownRenderer } from '../MarkdownRenderer';

import headingsFixture from '../../__tests__/fixtures/markdown/headings.md?raw';
import gfmTableFixture from '../../__tests__/fixtures/markdown/gfm-table.md?raw';
import taskListFixture from '../../__tests__/fixtures/markdown/task-list.md?raw';
import strikethroughAutolinksFixture from '../../__tests__/fixtures/markdown/strikethrough-autolinks.md?raw';
import codeFenceFixture from '../../__tests__/fixtures/markdown/code-fence.md?raw';
import mathFixture from '../../__tests__/fixtures/markdown/math.md?raw';
import mermaidFixture from '../../__tests__/fixtures/markdown/mermaid.md?raw';
import nestedListsFixture from '../../__tests__/fixtures/markdown/nested-lists.md?raw';
import blockquoteFixture from '../../__tests__/fixtures/markdown/blockquote.md?raw';
import rawHtmlFixture from '../../__tests__/fixtures/markdown/raw-html.md?raw';
import linksImagesFixture from '../../__tests__/fixtures/markdown/links-images.md?raw';
import kitchenSinkFixture from '../../__tests__/fixtures/markdown/kitchen-sink.md?raw';

// ---------------------------------------------------------------------------
// Mermaid mock — "Mermaid will not render in jsdom" per the task brief.
//
// Declared via vi.hoisted so the mock functions exist before vi.mock's
// hoisted factory runs. The mocked svg deliberately does NOT echo back the
// `id` argument it's called with — that id is derived from React's useId()
// (see MarkdownRenderer.tsx) and is not guaranteed stable across unrelated
// renders elsewhere in the same test file, so keeping it out of the DOM
// output sidesteps the non-determinism entirely rather than papering over
// it with a regex. `stabilizeHtml` below still normalizes the pattern
// defensively, so a future change that DOES leak a useId string into
// markup fails loudly (a snapshot full of "STABLE-ID") instead of flaking.
// ---------------------------------------------------------------------------

const { mermaidRenderMock, mermaidInitializeMock } = vi.hoisted(() => ({
  // Signature intentionally matches mermaid's real `render(id, code)` so the
  // recorded `.mock.calls` are typed as [string, string] for the assertions
  // below; the arguments themselves aren't used in the return value (see the
  // comment above), just captured — hence the `void` no-ops.
  mermaidRenderMock: vi.fn(async (id: string, code: string) => {
    void id;
    void code;
    return { svg: '<svg data-mock-mermaid="true" role="img"><text>MOCKED MERMAID SVG</text></svg>' };
  }),
  mermaidInitializeMock: vi.fn(),
}));

vi.mock('mermaid', () => ({
  default: {
    initialize: mermaidInitializeMock,
    render: mermaidRenderMock,
  },
}));

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function stabilizeHtml(html: string): string {
  return html
    .replace(/:r[0-9a-z]+:/gi, ':rSTABLE:')
    .replace(/mermaid-[a-zA-Z0-9]+-\d+/g, 'mermaid-STABLE-ID');
}

/** Pulls a fenced code block's exact body (minus its one trailing newline,
 * mirroring MarkdownRenderer's own `raw` computation) out of a fixture. */
function extractFencedBlock(source: string, lang: string): string {
  const re = new RegExp('```' + lang + '\\n([\\s\\S]*?)```', 'm');
  const match = re.exec(source);
  if (!match) throw new Error(`fixture is missing a \`\`\`${lang} fence`);
  return match[1].replace(/\n$/, '');
}

async function renderMarkdown(markdown: string) {
  const utils = render(<MarkdownRenderer markdown={markdown} />);
  if (markdown.includes('```mermaid')) {
    await waitFor(() => {
      expect(utils.container.querySelector('[data-mock-mermaid]')).toBeTruthy();
    });
  }
  return utils;
}

const FIXTURES: ReadonlyArray<readonly [name: string, markdown: string]> = [
  ['headings', headingsFixture],
  ['gfm-table', gfmTableFixture],
  ['task-list', taskListFixture],
  ['strikethrough-autolinks', strikethroughAutolinksFixture],
  ['code-fence', codeFenceFixture],
  ['math', mathFixture],
  ['mermaid', mermaidFixture],
  ['nested-lists', nestedListsFixture],
  ['blockquote', blockquoteFixture],
  ['raw-html', rawHtmlFixture],
  ['links-images', linksImagesFixture],
  ['kitchen-sink', kitchenSinkFixture],
];

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

beforeEach(() => {
  mermaidRenderMock.mockClear();
  mermaidInitializeMock.mockClear();
});

afterEach(() => {
  document.documentElement.removeAttribute('data-theme');
});

describe('MarkdownRenderer characterization', () => {
  for (const [name, markdown] of FIXTURES) {
    it(`renders the "${name}" fixture unchanged`, async () => {
      const { container } = await renderMarkdown(markdown);
      expect(stabilizeHtml(container.innerHTML)).toMatchSnapshot();
    });
  }

  it('hands the exact mermaid source, and a "mermaid-"-prefixed id, to mermaid.render', async () => {
    await renderMarkdown(mermaidFixture);
    const expectedSource = extractFencedBlock(mermaidFixture, 'mermaid');

    expect(mermaidRenderMock).toHaveBeenCalledTimes(1);
    const [id, code] = mermaidRenderMock.mock.calls[0]!;
    expect(id).toMatch(/^mermaid-[a-zA-Z0-9]+-\d+$/);
    expect(code).toBe(expectedSource);
  });

  it('initializes mermaid with the "default" theme when no data-theme is set', async () => {
    await renderMarkdown(mermaidFixture);
    expect(mermaidInitializeMock).toHaveBeenCalledWith(
      expect.objectContaining({ theme: 'default' }),
    );
  });

  it.each(['dark', 'dracula', 'nord'] as const)(
    'initializes mermaid with the "dark" theme when data-theme="%s"',
    async (theme) => {
      document.documentElement.setAttribute('data-theme', theme);
      await renderMarkdown(mermaidFixture);
      expect(mermaidInitializeMock).toHaveBeenCalledWith(
        expect.objectContaining({ theme: 'dark' }),
      );
    },
  );

  describe('getComputedStyle probe (partial CSS-blind-spot mitigation)', () => {
    // jsdom loads no stylesheet (vitest.config.ts sets `css: false`), so
    // color/background values below are jsdom's UA-default initial values,
    // not this app's real theme colors — they cannot catch an actual visual
    // regression. What they DO catch cheaply: an element silently losing
    // the class/tag a real stylesheet selects on (e.g. a heading rendering
    // as a <p>, or `.markdown-body` disappearing), which flips `display`
    // for that selector. Kept per the task brief; not a substitute for real
    // pixel-diff visual regression testing (out of scope, see DEFER-8).
    const SELECTORS = ['.markdown-body', 'h1', 'h2', 'pre', 'code', 'blockquote', 'table'] as const;
    const THEMES = ['light', 'dark', 'sepia', 'nord', 'dracula'] as const;

    it.each(THEMES)('snapshots key container computed styles under theme="%s"', (theme) => {
      document.documentElement.setAttribute('data-theme', theme);
      const { container } = render(<MarkdownRenderer markdown={kitchenSinkFixture} />);

      const snapshot: Record<string, { display: string; color: string; backgroundColor: string }> = {};
      for (const selector of SELECTORS) {
        const el = container.querySelector(selector);
        if (!el) continue;
        const computed = getComputedStyle(el);
        snapshot[selector] = {
          display: computed.display,
          color: computed.color,
          backgroundColor: computed.backgroundColor,
        };
      }

      expect(snapshot).toMatchSnapshot();
    });
  });
});
