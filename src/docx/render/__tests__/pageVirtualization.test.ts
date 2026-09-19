import { describe, expect, it } from 'vitest';

import type { Page } from '../../layout/pageTypes';
import {
  computePageOffsets,
  findPagesForParagraphPath,
  findVisiblePageRange,
  PAGE_STACK_GAP_PX,
  PAGE_STACK_PADDING_PX,
  PT_TO_PX,
} from '../pageVirtualization';

/** Minimal fake `Page` — only the fields `pageVirtualization.ts` reads. */
function fakePage(overrides: Partial<Page> & { pageIndex: number }): Page {
  return {
    sectionIndex: 0,
    sizePt: { width: 612, height: 792 },
    marginsPt: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36, gutter: 0 },
    columns: [],
    headerLines: [],
    footerLines: [],
    footnoteLines: [],
    hasFootnoteSeparator: false,
    ...overrides,
  } as Page;
}

describe('computePageOffsets', () => {
  it('stacks pages using the same padding/gap .docx-page-stack applies in CSS', () => {
    const pages = [
      fakePage({ pageIndex: 0, sizePt: { width: 100, height: 100 } }),
      fakePage({ pageIndex: 1, sizePt: { width: 100, height: 200 } }),
      fakePage({ pageIndex: 2, sizePt: { width: 100, height: 50 } }),
    ];

    const { offsets, totalHeight } = computePageOffsets(pages, 1);
    const scale = PT_TO_PX;

    expect(offsets).toHaveLength(3);
    expect(offsets[0]).toEqual({ top: PAGE_STACK_PADDING_PX, height: 100 * scale });
    expect(offsets[1].top).toBeCloseTo(PAGE_STACK_PADDING_PX + 100 * scale + PAGE_STACK_GAP_PX);
    expect(offsets[1].height).toBeCloseTo(200 * scale);
    expect(offsets[2].top).toBeCloseTo(offsets[1].top + offsets[1].height + PAGE_STACK_GAP_PX);

    const expectedTotal =
      PAGE_STACK_PADDING_PX * 2 + (100 + 200 + 50) * scale + PAGE_STACK_GAP_PX * 2;
    expect(totalHeight).toBeCloseTo(expectedTotal);
  });

  it('scales with zoom', () => {
    const pages = [fakePage({ pageIndex: 0, sizePt: { width: 100, height: 100 } })];
    const atZoom1 = computePageOffsets(pages, 1);
    const atZoom2 = computePageOffsets(pages, 2);
    expect(atZoom2.offsets[0].height).toBeCloseTo(atZoom1.offsets[0].height * 2);
  });

  it('returns an empty result for no pages', () => {
    const { offsets, totalHeight } = computePageOffsets([], 1);
    expect(offsets).toEqual([]);
    expect(totalHeight).toBe(0);
  });
});

describe('findVisiblePageRange', () => {
  // Three 100px-tall pages, back to back with no gap, for simple round numbers.
  const offsets = [
    { top: 0, height: 100 },
    { top: 100, height: 100 },
    { top: 200, height: 100 },
  ];

  it('returns an empty range for no pages', () => {
    expect(findVisiblePageRange([], 0, 500, 0)).toEqual({ start: 0, end: -1 });
  });

  it('selects only the page(s) intersecting the viewport with no buffer', () => {
    expect(findVisiblePageRange(offsets, 0, 50, 0)).toEqual({ start: 0, end: 0 });
    expect(findVisiblePageRange(offsets, 120, 50, 0)).toEqual({ start: 1, end: 1 });
  });

  it('expands the range by the buffer on both sides', () => {
    // Viewport is entirely inside page 1 (top=120..170); a 100px buffer
    // reaches into pages 0 and 2 too.
    expect(findVisiblePageRange(offsets, 120, 50, 100)).toEqual({ start: 0, end: 2 });
  });

  it('a viewport spanning every page returns the full range', () => {
    expect(findVisiblePageRange(offsets, 0, 10_000, 0)).toEqual({ start: 0, end: 2 });
  });

  it('a viewport entirely past the last page returns an empty range', () => {
    expect(findVisiblePageRange(offsets, 10_000, 50, 0)).toEqual({ start: 0, end: -1 });
  });
});

describe('findPagesForParagraphPath', () => {
  function pageWithLine(pageIndex: number, sectionIndex: number, blockIndex: number): Page {
    return fakePage({
      pageIndex,
      sectionIndex,
      columns: [
        {
          widthPt: 400,
          leftPt: 0,
          tables: [],
          lines: [{ paragraphPath: [blockIndex], lineIndex: 0, line: {} as never, topPt: 0, leftPt: 0 }],
        },
      ],
    });
  }

  function pageWithTable(pageIndex: number, sectionIndex: number, blockIndex: number): Page {
    return fakePage({
      pageIndex,
      sectionIndex,
      columns: [
        {
          widthPt: 400,
          leftPt: 0,
          lines: [],
          tables: [{ blockPath: [blockIndex] } as never],
        },
      ],
    });
  }

  it('returns [] for an undefined or empty path', () => {
    const pages = [pageWithLine(0, 0, 0)];
    expect(findPagesForParagraphPath(pages, undefined)).toEqual([]);
    expect(findPagesForParagraphPath(pages, [])).toEqual([]);
  });

  it('matches a single-section path ([blockIndex]) against section 0', () => {
    const pages = [pageWithLine(0, 0, 0), pageWithLine(1, 0, 1)];
    expect(findPagesForParagraphPath(pages, [1])).toEqual([1]);
  });

  it('matches a multi-section path ([sectionIndex, blockIndex])', () => {
    const pages = [pageWithLine(0, 0, 3), pageWithLine(1, 1, 3)];
    expect(findPagesForParagraphPath(pages, [1, 3])).toEqual([1]);
    expect(findPagesForParagraphPath(pages, [0, 3])).toEqual([0]);
  });

  it('matches a table by its blockPath', () => {
    const pages = [pageWithLine(0, 0, 0), pageWithTable(1, 0, 5)];
    expect(findPagesForParagraphPath(pages, [5])).toEqual([1]);
  });

  it('returns every page a split paragraph appears on, not just the first', () => {
    const pages = [pageWithLine(0, 0, 2), pageWithLine(1, 0, 2), pageWithLine(2, 0, 9)];
    expect(findPagesForParagraphPath(pages, [2])).toEqual([0, 1]);
  });

  it('returns [] when nothing matches', () => {
    const pages = [pageWithLine(0, 0, 0)];
    expect(findPagesForParagraphPath(pages, [99])).toEqual([]);
  });
});
