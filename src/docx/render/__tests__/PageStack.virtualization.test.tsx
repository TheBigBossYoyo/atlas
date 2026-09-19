/**
 * D23-PERF-2 — `PageStack`'s scroll-driven page virtualization (see
 * `pageVirtualization.ts` and `PageStack.tsx`'s own doc comments for why
 * this exists). `pageVirtualization.test.ts` covers the pure geometry/
 * lookup functions in isolation; this file exercises `PageStack` itself:
 * that mounting only narrows around the measured viewport, that pinned
 * pages and `forceRenderAll` override that, and that a real scroll actually
 * moves the mounted window.
 */
import { createRef } from 'react';
import { describe, expect, it, vi } from 'vitest';
import { render, waitFor } from '@testing-library/react';

import { PageStack } from '../PageStack';
import { paginate } from '../../layout/paginate';
import type { Document, NumberingDef, Section } from '../../model';
import { twip } from '../../model';

function createDocument(sections: Section[], numbering: ReadonlyMap<string, NumberingDef> = new Map()): Document {
  return {
    kind: 'document',
    sections,
    styles: new Map(),
    numbering,
    headers: new Map(),
    footers: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
  } as Document;
}

function createSection(blocks: Section['blocks'], pageHeightPt: number): Section {
  return {
    kind: 'section',
    props: {
      pgSz: { w: twip(400 * 20), h: twip(pageHeightPt * 20) },
      pgMar: { top: twip(0), right: twip(0), bottom: twip(0), left: twip(0), header: twip(0), footer: twip(0), gutter: twip(0) },
      cols: { num: 1, space: twip(0), col: [] },
    },
    blocks,
  };
}

/** One paragraph per page: each page is forced to hold exactly one, via a
 * `pageHeightPt` just tall enough for a single line. */
function createParagraph(text: string) {
  return {
    kind: 'paragraph' as const,
    props: {},
    children: [{ kind: 'run' as const, props: {}, children: [{ kind: 'text' as const, value: text }] }],
  };
}

const mockFontResolver = async () => ({
  unitsPerEm: 1000,
  ascender: 800,
  descender: -200,
  lineGap: 0,
  xHeight: 500,
  capHeight: 700,
  advanceWidth: () => 500,
  hasGlyph: () => true,
});

/** Builds a document that paginates into exactly `pageCount` pages, one
 * short paragraph each (a tiny `pageHeightPt` forces every paragraph onto
 * its own page). */
async function buildPages(pageCount: number) {
  const paragraphs = Array.from({ length: pageCount }, (_, i) => createParagraph(`page ${i}`));
  const doc = createDocument([createSection(paragraphs, 20)]);
  const pages = await paginate({ document: doc, fontResolver: mockFontResolver });
  expect(pages).toHaveLength(pageCount);
  return { doc, pages };
}

/** Stubs `clientHeight`/`scrollTop` on a real jsdom element — jsdom never
 * lays anything out, so both are normally always 0. `configurable: true`
 * lets a later call re-stub the same element (e.g. to simulate a scroll). */
function stubViewport(element: HTMLElement, scrollTop: number, clientHeight: number): void {
  Object.defineProperty(element, 'clientHeight', { value: clientHeight, configurable: true });
  Object.defineProperty(element, 'scrollTop', { value: scrollTop, configurable: true, writable: true });
}

describe('PageStack virtualization', () => {
  it('mounts only pages near a measured viewport, placeholdering the rest', async () => {
    const { doc, pages } = await buildPages(20);
    const containerRef = createRef<HTMLDivElement>();

    // Mount the scroll container FIRST (no PageStack yet) so its ref is
    // already populated — matching how DocxViewer's own editor surface is
    // already mounted, with real geometry, before `pages` (and so
    // PageStack) exists for the first time — then stub its geometry before
    // PageStack ever reads it.
    const { rerender } = render(<div ref={containerRef} />);
    stubViewport(containerRef.current!, 0, 100);

    rerender(
      <div ref={containerRef}>
        <PageStack pages={pages} zoom={1} document={doc} scrollContainerRef={containerRef} />
      </div>,
    );

    const stack = containerRef.current!.querySelector('.docx-page-stack')!;
    expect(stack.getAttribute('data-page-count')).toBe('20');

    const real = stack.querySelectorAll('.docx-page:not([data-virtualized="placeholder"])');
    const placeholders = stack.querySelectorAll('[data-virtualized="placeholder"]');

    // A 100px-tall viewport against pages far taller than that (scaled from
    // 20pt) mounts a handful of pages, not all 20 — and definitely not zero.
    expect(real.length).toBeGreaterThan(0);
    expect(real.length).toBeLessThan(20);
    expect(placeholders.length).toBe(20 - real.length);

    // Every placeholder still reserves the same box a real page would —
    // scroll height / the scrollbar must not depend on which pages happen
    // to be mounted.
    for (const placeholder of Array.from(placeholders)) {
      const el = placeholder as HTMLElement;
      expect(el.style.width).not.toBe('');
      expect(el.style.height).not.toBe('');
    }

    // The first page (closest to scrollTop 0) is definitely mounted.
    expect(stack.querySelector('[data-page-index="0"]:not([data-virtualized="placeholder"])')).not.toBeNull();
    // A page far below the viewport+buffer is not.
    expect(stack.querySelector('[data-page-index="19"]:not([data-virtualized="placeholder"])')).toBeNull();
  });

  it('keeps pinned pages mounted even when they are outside the viewport', async () => {
    const { doc, pages } = await buildPages(20);
    const containerRef = createRef<HTMLDivElement>();

    const { rerender } = render(<div ref={containerRef} />);
    stubViewport(containerRef.current!, 0, 100);

    rerender(
      <div ref={containerRef}>
        <PageStack
          pages={pages}
          zoom={1}
          document={doc}
          scrollContainerRef={containerRef}
          pinnedPageIndices={new Set([19])}
        />
      </div>,
    );

    const stack = containerRef.current!.querySelector('.docx-page-stack')!;
    // Page 19 is far outside the near-scrollTop-0 viewport, but pinned —
    // it must still be a real PageView, not a placeholder (the caret/
    // selection-sync use case this exists for needs its real DOM).
    expect(stack.querySelector('[data-page-index="19"]:not([data-virtualized="placeholder"])')).not.toBeNull();
  });

  it('forceRenderAll mounts every page regardless of the viewport', async () => {
    const { doc, pages } = await buildPages(20);
    const containerRef = createRef<HTMLDivElement>();

    const { rerender } = render(<div ref={containerRef} />);
    stubViewport(containerRef.current!, 0, 100);

    rerender(
      <div ref={containerRef}>
        <PageStack pages={pages} zoom={1} document={doc} scrollContainerRef={containerRef} forceRenderAll />
      </div>,
    );

    const stack = containerRef.current!.querySelector('.docx-page-stack')!;
    expect(stack.querySelectorAll('.docx-page:not([data-virtualized="placeholder"])')).toHaveLength(20);
    expect(stack.querySelectorAll('[data-virtualized="placeholder"]')).toHaveLength(0);
  });

  it('scrolling the container moves the mounted window', async () => {
    const { doc, pages } = await buildPages(20);
    const containerRef = createRef<HTMLDivElement>();

    const { rerender } = render(<div ref={containerRef} />);
    stubViewport(containerRef.current!, 0, 100);

    rerender(
      <div ref={containerRef}>
        <PageStack pages={pages} zoom={1} document={doc} scrollContainerRef={containerRef} />
      </div>,
    );

    const container = containerRef.current!;
    expect(container.querySelector('[data-page-index="0"]:not([data-virtualized="placeholder"])')).not.toBeNull();
    expect(container.querySelector('[data-page-index="19"]:not([data-virtualized="placeholder"])')).toBeNull();

    // jsdom never computes real layout, so `getBoundingClientRect()` always
    // returns a static all-zero rect — `PageStack` uses the stack's rect
    // (relative to the container's) to translate `scrollTop` into the
    // stack's own coordinate space (see `measureViewport`'s doc comment),
    // which needs the stack to actually appear to move as the container
    // scrolls. This nested test's stack has no siblings above it (zero
    // fixed offset), so `stack.top` should simply track `-scrollTop`.
    const stack = container.querySelector<HTMLElement>('.docx-page-stack')!;
    vi.spyOn(stack, 'getBoundingClientRect').mockImplementation(
      () => ({ top: -container.scrollTop, left: 0, right: 0, bottom: 0, width: 0, height: 0, x: 0, y: 0, toJSON: () => ({}) }) as DOMRect,
    );

    // Scroll to just above the last page's offset (20 pages, each ~26.7px
    // tall at this zoom/pageHeightPt, plus the stack's own padding/gap —
    // see `computePageOffsets`) and let the rAF-throttled scroll handler
    // run. Scrolling PAST the document's total height would legitimately
    // produce an empty visible range (nothing to show) rather than "the
    // last page", so this lands just inside it instead.
    Object.defineProperty(container, 'scrollTop', { value: 950, configurable: true, writable: true });
    container.dispatchEvent(new Event('scroll'));

    await waitFor(() => {
      expect(container.querySelector('[data-page-index="19"]:not([data-virtualized="placeholder"])')).not.toBeNull();
    });
    expect(container.querySelector('[data-page-index="0"]:not([data-virtualized="placeholder"])')).toBeNull();
  });
});
