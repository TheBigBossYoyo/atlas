import { describe, expect, it, vi } from 'vitest';
import { render } from '@testing-library/react';

import { PageView } from '../PageView';
import { PageStack } from '../PageStack';
import { paginate } from '../../layout/paginate';
import type { Document, NumberingDef, Section, ParaProps } from '../../model';
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

function createSection(
  blocks: Section['blocks'],
  options: {
    pageWidthPt?: number;
    pageHeightPt?: number;
    columnCount?: number;
    columnSpacePt?: number;
  } = {}
): Section {
  return {
    kind: 'section',
    props: {
      pgSz: {
        w: twip((options.pageWidthPt ?? 400) * 20),
        h: twip((options.pageHeightPt ?? 200) * 20),
      },
      pgMar: {
        top: twip(0),
        right: twip(0),
        bottom: twip(0),
        left: twip(0),
        header: twip(0),
        footer: twip(0),
        gutter: twip(0),
      },
      cols: {
        num: options.columnCount ?? 1,
        space: twip((options.columnSpacePt ?? 0) * 20),
        col: [],
      },
    },
    blocks,
  };
}

function createParagraph(wordCount: number, props: ParaProps = {}) {
  return {
    kind: 'paragraph' as const,
    props,
    children: [
      {
        kind: 'run' as const,
        props: {},
        children: Array.from({ length: wordCount }, (_, i) => ({
          kind: 'text' as const,
          value: i === wordCount - 1 ? 'word' : 'word ',
        })),
      },
    ],
  };
}

const mockFontResolver = vi.fn().mockResolvedValue({
  unitsPerEm: 1000,
  ascender: 800,
  descender: -200,
  lineGap: 0,
  xHeight: 500,
  capHeight: 700,
  advanceWidth: () => 500,
  hasGlyph: () => true,
});

describe('PageView', () => {
  it('renders a single page with single column', async () => {
    const doc = createDocument([createSection([createParagraph(5)])]); // 25pt long approx
    const pages = await paginate({
      document: doc,
      fontResolver: mockFontResolver,
    });

    expect(pages).toHaveLength(1);

    const { container } = render(<PageView page={pages[0]} zoom={1} document={doc} />);
    const pageEl = container.querySelector('.docx-page');
    expect(pageEl).not.toBeNull();
    
    // Size check
    const outerEl = pageEl as HTMLElement;
    expect(outerEl.style.width).toBe(`${400 * (4/3)}px`);
    expect(outerEl.style.height).toBe(`${200 * (4/3)}px`);

    const lines = container.querySelectorAll('.docx-page__line');
    expect(lines.length).toBeGreaterThan(0);
    // at least 1 line
  });

  it('renders multi-column pages correctly', async () => {
    const doc = createDocument([
      createSection(
        [createParagraph(50)], // long paragraph
        { columnCount: 2, columnSpacePt: 20 }
      ),
    ]);
    const pages = await paginate({
      document: doc,
      fontResolver: mockFontResolver,
    });

    const { container } = render(<PageView page={pages[0]} zoom={2} document={doc} />);
    const columns = container.querySelectorAll('.docx-page__column');
    expect(columns.length).toBe(2);

    const pageEl = container.querySelector('.docx-page') as HTMLElement;
    // zoom = 2
    expect(pageEl.style.width).toBe(`${400 * 2 * (4/3)}px`);
  });

  it('renders a numbered list marker with the docx-list-marker class (D3/DXP-07/DXL-04)', async () => {
    const numbering = new Map<string, NumberingDef>([
      [
        '1',
        {
          numId: '1',
          levels: new Map([
            [0, { level: 0, format: 'decimal', text: { value: '%1.', placeholders: [1] }, suffix: 'tab' }],
          ]),
        },
      ],
    ]);

    const doc = createDocument(
      [
        createSection([
          {
            kind: 'paragraph' as const,
            props: { numPr: { numId: '1', ilvl: 0 }, ind: { left: twip(720), hanging: twip(360) } },
            children: [{ kind: 'run' as const, children: [{ kind: 'text' as const, value: 'first item' }] }],
          },
        ]),
      ],
      numbering,
    );
    const pages = await paginate({ document: doc, fontResolver: mockFontResolver });

    const { container } = render(<PageView page={pages[0]} zoom={1} document={doc} />);
    const marker = container.querySelector('.docx-list-marker');
    expect(marker?.textContent).toBe('1.');
  });

  it('stretches a justified line flush to the column width (D5/DXL-11)', async () => {
    const doc = createDocument([
      createSection([createParagraph(4, { jc: 'both' })], { pageWidthPt: 100 }),
    ]);
    const pages = await paginate({
      document: doc,
      fontResolver: mockFontResolver,
    });

    const { container } = render(<PageView page={pages[0]} zoom={1} document={doc} />);
    const lines = container.querySelectorAll('.docx-page__line');
    // The paragraph wraps to more than one line at this column width, so its
    // non-final first line is eligible for "both" justification.
    expect(lines.length).toBeGreaterThan(1);

    const firstLine = lines[0] as HTMLElement;
    expect(firstLine.style.width).toBe('100px');
  });

  it('leaves a non-justified line at its natural (ragged) width', async () => {
    const doc = createDocument([
      createSection([createParagraph(4)], { pageWidthPt: 100 }),
    ]);
    const pages = await paginate({
      document: doc,
      fontResolver: mockFontResolver,
    });

    const { container } = render(<PageView page={pages[0]} zoom={1} document={doc} />);
    const firstLine = container.querySelector('.docx-page__line') as HTMLElement;
    expect(firstLine.style.width).not.toBe('100px');
  });

  it('renders pages with headers and footers', async () => {
    const doc = createDocument([createSection([createParagraph(5)])]);
    const pages = await paginate({
      document: doc,
      fontResolver: mockFontResolver,
      headerFooterLines: new Map([
        ['header1', [{ items: [], width: 100, ascent: 10, descent: 2, lineHeight: 12, isJustified: false, justificationStretch: 0 }]],
        ['footer1', [{ items: [], width: 100, ascent: 10, descent: 2, lineHeight: 12, isJustified: false, justificationStretch: 0 }]],
      ]),
    });
    
    // Manually push header and footer to the mocked page since paginate tests inject it somehow
    // Let's just create a modified page
    const pageWithHeaderFooter = {
      ...pages[0],
      headerLines: [{ items: [], width: 100, ascent: 10, descent: 2, lineHeight: 12, isJustified: false, justificationStretch: 0 }],
      footerLines: [{ items: [], width: 100, ascent: 10, descent: 2, lineHeight: 12, isJustified: false, justificationStretch: 0 }]
    };

    const { container } = render(<PageView page={pageWithHeaderFooter} zoom={1} document={doc} />);
    
    expect(container.querySelector('.docx-page__header')).not.toBeNull();
    expect(container.querySelector('.docx-page__footer')).not.toBeNull();
  });
});

describe('PageStack', () => {
  it('renders multiple pages stacked', async () => {
    const doc = createDocument([
      createSection([createParagraph(200)], { pageHeightPt: 100 }) // Forces multiple pages
    ]);
    const pages = await paginate({
      document: doc,
      fontResolver: mockFontResolver,
    });

    expect(pages.length).toBeGreaterThan(1);

    const { container } = render(<PageStack pages={pages} zoom={1} document={doc} />);
    const renderedPages = container.querySelectorAll('.docx-page');
    expect(renderedPages.length).toBe(pages.length);
  });
});
