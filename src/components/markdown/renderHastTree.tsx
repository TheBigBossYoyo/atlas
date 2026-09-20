/**
 * Renders an already-parsed hast tree the same way react-markdown's own
 * (unexported) `post()` step does — see
 * `node_modules/react-markdown/lib/index.js`. That function does two things
 * to a hast tree before handing it to `hast-util-to-jsx-runtime`: it walks
 * the tree once to sanitize URL-bearing attributes (`defaultUrlTransform`),
 * and it calls `toJsxRuntime` with a specific set of options. Neither is
 * exported, so this file vendors both, kept intentionally minimal and in
 * lockstep with that source so a tree parsed off-main-thread by
 * `markdownParse.worker.ts` renders identically to one `<ReactMarkdown>`
 * parses and renders inline.
 *
 * Deliberately NOT reproduced here: `allowedElements` / `disallowedElements`
 * / `allowElement` / `skipHtml` handling, and the `raw`-node-to-text
 * fallback. `MarkdownRenderer` never passes those props, and this
 * pipeline's own `rehypeRaw` (see `markdownPipeline.ts`) has already
 * resolved every raw-HTML node into real hast elements by the time a tree
 * reaches this function — so neither code path can ever run here. Skipping
 * them keeps this file small instead of vendoring dead code.
 */
import { Fragment, jsx, jsxs } from 'react/jsx-runtime';
import { toJsxRuntime } from 'hast-util-to-jsx-runtime';
import { urlAttributes } from 'html-url-attributes';
import { visit } from 'unist-util-visit';
import type { Root as HastRoot, Element as HastElement } from 'hast';
import type { Components } from 'react-markdown';
import type { ReactElement } from 'react';

// Byte-for-byte the same as react-markdown's `safeProtocol` / `defaultUrlTransform`.
const safeProtocol = /^(https?|ircs?|mailto|xmpp)$/i;

function defaultUrlTransform(value: string): string {
  const colon = value.indexOf(':');
  const questionMark = value.indexOf('?');
  const numberSign = value.indexOf('#');
  const slash = value.indexOf('/');

  if (
    colon === -1 ||
    (slash !== -1 && colon > slash) ||
    (questionMark !== -1 && colon > questionMark) ||
    (numberSign !== -1 && colon > numberSign) ||
    safeProtocol.test(value.slice(0, colon))
  ) {
    return value;
  }

  return '';
}

function sanitizeUrls(tree: HastRoot): void {
  visit(tree, 'element', (node: HastElement) => {
    let key: string;
    for (key in urlAttributes) {
      if (Object.hasOwn(urlAttributes, key) && Object.hasOwn(node.properties, key)) {
        const value = node.properties[key];
        const test = urlAttributes[key];
        if (test === null || test.includes(node.tagName)) {
          node.properties[key] = defaultUrlTransform(String(value ?? ''));
        }
      }
    }
  });
}

export function renderHastTree(tree: HastRoot, components?: Partial<Components> | null): ReactElement {
  sanitizeUrls(tree);

  return toJsxRuntime(tree, {
    Fragment,
    components,
    ignoreInvalidStyle: true,
    jsx,
    jsxs,
    passKeys: true,
    passNode: true,
  }) as ReactElement;
}
