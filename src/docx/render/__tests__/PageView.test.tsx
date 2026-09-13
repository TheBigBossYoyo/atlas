import { describe, expect, it, vi } from 'vitest';
import { render } from '@testing-library/react';

import { PageView } from '../PageView';
import { PageStack } from '../PageStack';
import { paginate } from '../../layout/paginate';
import type { Document, NumberingDef, Section, ParaProps } from '../../model';
import { twip } from '../../model';
import type { Relationship } from '../../parser/relationships';
import { loadDocx } from '../../index';
import { readCorpusFixture } from '../../__tests__/corpusRoundtripHelpers';

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

  it('renders an external hyperlink as a clickable, target=_blank <a> resolved via relationships (D6/DXL-12)', async () => {
    const relationships: Relationship[] = [
      { id: 'rId1', type: 'hyperlink', target: 'https://example.com/atlas', targetMode: 'External' },
    ];
    const doc = createDocument([
      createSection([
        {
          kind: 'paragraph' as const,
          props: {},
          children: [
            {
              kind: 'hyperlink' as const,
              relationshipId: 'rId1',
              children: [{ kind: 'run' as const, children: [{ kind: 'text' as const, value: 'a link' }] }],
            },
          ],
        },
      ]),
    ]);
    const pages = await paginate({ document: doc, fontResolver: mockFontResolver });

    const { container } = render(
      <PageView page={pages[0]} zoom={1} document={doc} relationships={relationships} />,
    );
    const links = Array.from(container.querySelectorAll('a.docx-hyperlink'));
    expect(links.length).toBeGreaterThan(0);
    for (const link of links) {
      expect(link.getAttribute('href')).toBe('https://example.com/atlas');
      expect(link.getAttribute('target')).toBe('_blank');
      expect(link.getAttribute('rel')).toBe('noopener noreferrer');
    }
    // Every item of the hyperlink's text (including the space between
    // words) is individually wrapped, so the whole span is clickable with
    // no gap — see PageView's wrapHyperlink.
    expect(links.map((link) => link.textContent).join('')).toBe('a link');
  });

  it('renders an internal hyperlink as a same-document anchor without target=_blank (D6/DXL-12)', async () => {
    const doc = createDocument([
      createSection([
        {
          kind: 'paragraph' as const,
          props: {},
          children: [
            {
              kind: 'hyperlink' as const,
              anchor: 'targetBookmark',
              children: [{ kind: 'run' as const, children: [{ kind: 'text' as const, value: 'jump' }] }],
            },
          ],
        },
        {
          kind: 'paragraph' as const,
          props: {},
          children: [
            { kind: 'bookmark' as const, id: '1', boundary: 'start' as const, name: 'targetBookmark' },
            { kind: 'run' as const, children: [{ kind: 'text' as const, value: 'Target Section' }] },
          ],
        },
      ]),
    ]);
    const pages = await paginate({ document: doc, fontResolver: mockFontResolver });

    const { container } = render(<PageView page={pages[0]} zoom={1} document={doc} />);
    const link = container.querySelector('a.docx-hyperlink');
    expect(link?.getAttribute('href')).toBe('#docx-bookmark-targetBookmark');
    expect(link?.hasAttribute('target')).toBe(false);
    expect(container.querySelector('#docx-bookmark-targetBookmark')).not.toBeNull();
  });

  it('does not render a hyperlink for a relationship id that is not External', async () => {
    const relationships: Relationship[] = [
      { id: 'rId1', type: 'hyperlink', target: 'word/document.xml', targetMode: 'Internal' },
    ];
    const doc = createDocument([
      createSection([
        {
          kind: 'paragraph' as const,
          props: {},
          children: [
            {
              kind: 'hyperlink' as const,
              relationshipId: 'rId1',
              children: [{ kind: 'run' as const, children: [{ kind: 'text' as const, value: 'suspicious' }] }],
            },
          ],
        },
      ]),
    ]);
    const pages = await paginate({ document: doc, fontResolver: mockFontResolver });

    const { container } = render(
      <PageView page={pages[0]} zoom={1} document={doc} relationships={relationships} />,
    );
    expect(container.querySelector('a.docx-hyperlink')).toBeNull();
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

describe('PageView against the real hyperlinks-bookmarks corpus fixture (D6/DXL-12)', () => {
  it('renders a working external hyperlink and a same-document internal anchor', async () => {
    const buffer = await readCorpusFixture('hyperlinks-bookmarks');
    const bundle = await loadDocx(buffer);
    const pages = await paginate({ document: bundle.document, fontResolver: mockFontResolver });

    const { container } = render(
      <PageStack
        pages={pages}
        zoom={1}
        document={bundle.document}
        relationships={bundle.relationships}
      />,
    );

    const links = Array.from(container.querySelectorAll('a.docx-hyperlink'));
    expect(links.length).toBeGreaterThan(0);

    // Per scripts/generate-docx-corpus.mjs's fixtureHyperlinksBookmarks:
    // one external link to https://example.com/atlas, one internal link to
    // the "targetBookmark" bookmark rendered later in the same document.
    const externalLink = links.find((link) => link.getAttribute('href') === 'https://example.com/atlas');
    expect(externalLink?.getAttribute('target')).toBe('_blank');
    expect(externalLink?.getAttribute('rel')).toBe('noopener noreferrer');

    const internalLink = links.find((link) => link.getAttribute('href') === '#docx-bookmark-targetBookmark');
    expect(internalLink).not.toBeUndefined();
    expect(internalLink?.hasAttribute('target')).toBe(false);
    expect(container.querySelector('#docx-bookmark-targetBookmark')).not.toBeNull();
  });
});
