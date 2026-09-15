import { describe, expect, it, vi } from 'vitest';
import { fireEvent, render } from '@testing-library/react';

import { PageView } from '../PageView';
import { MediaContext, type MediaResolver } from '../mediaContext';
import { paginate } from '../../layout/paginate';
import type { Document, Drawing, RunChild, Section } from '../../model';
import { twip } from '../../model';

const EMUS_PER_POINT = 12700;

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

function createDocument(children: ReadonlyArray<RunChild>): Document {
  const section: Section = {
    kind: 'section',
    props: {
      pgSz: { w: twip(600 * 20), h: twip(800 * 20) },
      pgMar: {
        top: twip(0),
        right: twip(0),
        bottom: twip(0),
        left: twip(0),
        header: twip(0),
        footer: twip(0),
        gutter: twip(0),
      },
      cols: { num: 1, space: twip(0), col: [] },
    },
    blocks: [
      {
        kind: 'paragraph',
        props: {},
        children: [{ kind: 'run', props: {}, children }],
      },
    ],
  };

  return {
    kind: 'document',
    sections: [section],
    styles: new Map(),
    numbering: new Map(),
    headers: new Map(),
    footers: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
  } as Document;
}

function createDrawing(overrides: Partial<Drawing> = {}): Drawing {
  return {
    kind: 'drawing',
    layout: 'inline',
    relationshipId: 'rId5',
    description: 'Quarterly revenue chart',
    extent: { cx: 150 * EMUS_PER_POINT, cy: 90 * EMUS_PER_POINT },
    ...overrides,
  };
}

function resolverFor(entries: Record<string, string>): MediaResolver {
  return { resolve: (relationshipId) => entries[relationshipId] ?? null };
}

async function renderPage(children: ReadonlyArray<RunChild>, resolver: MediaResolver) {
  const doc = createDocument(children);
  const pages = await paginate({ document: doc, fontResolver: mockFontResolver });

  return render(
    <MediaContext.Provider value={resolver}>
      <PageView page={pages[0]} zoom={1} document={doc} />
    </MediaContext.Provider>,
  );
}

describe('PageView drawings', () => {
  it('paints an embedded picture as a real <img> resolved through MediaContext', async () => {
    const { container } = await renderPage(
      [createDrawing()],
      resolverFor({ rId5: 'blob:atlas/chart' }),
    );

    const img = container.querySelector('img');
    expect(img).not.toBeNull();
    expect(img).toHaveAttribute('src', 'blob:atlas/chart');
    expect(img).toHaveAttribute('alt', 'Quarterly revenue chart');
    expect(img?.style.width).toBe('150px');
    expect(img?.style.height).toBe('90px');
  });

  it('paints anchored (floating) pictures instead of dropping them', async () => {
    const { container } = await renderPage(
      [createDrawing({ layout: 'anchor', relationshipId: 'rId9' })],
      resolverFor({ rId9: 'blob:atlas/floating' }),
    );

    expect(container.querySelector('img')).toHaveAttribute('src', 'blob:atlas/floating');
    expect(container.querySelector('[data-drawing-layout="anchor"]')).not.toBeNull();
  });

  it('does not leak the object replacement character into a visible text run', async () => {
    const { container } = await renderPage(
      [{ kind: 'text', value: 'See chart ' }, createDrawing()],
      resolverFor({ rId5: 'blob:atlas/chart' }),
    );

    const runs = Array.from(container.querySelectorAll('span.docx-run'));
    expect(runs.length).toBeGreaterThan(0);
    expect(runs.some((run) => run.textContent?.includes('￼'))).toBe(false);
  });

  it('keeps editor char mapping attributes on the drawing slot', async () => {
    const { container } = await renderPage(
      [{ kind: 'text', value: 'ab' }, createDrawing(), { kind: 'text', value: 'cd' }],
      resolverFor({ rId5: 'blob:atlas/chart' }),
    );

    const slot = container.querySelector('.docx-drawing') as HTMLElement;
    expect(slot).not.toBeNull();
    expect(slot.dataset.runIndex).toBe('0');
    expect(slot.dataset.charStart).toBe('2');
    expect(slot.dataset.charEnd).toBe('3');
    // One DOM character per drawing, matching the model's char offsets.
    expect(slot.textContent).toBe('￼');
  });

  it('shows a visible placeholder when the picture cannot be resolved', async () => {
    const { container, getByRole } = await renderPage([createDrawing()], resolverFor({}));

    expect(container.querySelector('img')).toBeNull();
    const placeholder = getByRole('img', { name: 'Quarterly revenue chart' });
    expect(placeholder).toHaveClass('docx-drawing__placeholder');
    expect(placeholder.style.width).toBe('150px');
    expect(placeholder.style.height).toBe('90px');
  });

  it('falls back to the placeholder when the image fails to decode', async () => {
    const { container, getByRole } = await renderPage(
      [createDrawing()],
      resolverFor({ rId5: 'blob:atlas/unsupported-emf' }),
    );

    const img = container.querySelector('img') as HTMLImageElement;
    fireEvent.error(img);

    expect(container.querySelector('img')).toBeNull();
    expect(getByRole('img', { name: 'Quarterly revenue chart' })).toHaveClass('docx-drawing__placeholder');
  });

  it('positions a floating drawing at its page-absolute wp:positionH/V rect (D4/DXL-03/DXP-09)', async () => {
    const { container } = await renderPage(
      [
        createDrawing({
          layout: 'anchor',
          relationshipId: 'rId9',
          extent: { cx: 40 * EMUS_PER_POINT, cy: 30 * EMUS_PER_POINT },
          positionH: { relativeFrom: 'page', offsetEmu: 20 * EMUS_PER_POINT },
          positionV: { relativeFrom: 'page', offsetEmu: 10 * EMUS_PER_POINT },
          wrap: { mode: 'square' },
        }),
      ],
      resolverFor({ rId9: 'blob:atlas/floating' }),
    );

    const float = container.querySelector('.docx-anchored-drawing') as HTMLElement;
    expect(float).not.toBeNull();
    expect(float.style.left).toBe('20px');
    expect(float.style.top).toBe('10px');
    expect(float.style.width).toBe('40px');
    expect(float.style.height).toBe('30px');
  });

  it('does not let the anchored picture widen its line or reserve baseline clearance (D4/DXL-03)', async () => {
    const { container } = await renderPage(
      [
        { kind: 'text', value: 'Body text ' },
        createDrawing({
          layout: 'anchor',
          relationshipId: 'rId9',
          extent: { cx: 200 * EMUS_PER_POINT, cy: 300 * EMUS_PER_POINT },
          positionH: { relativeFrom: 'page', offsetEmu: 0 },
          positionV: { relativeFrom: 'page', offsetEmu: 0 },
        }),
        { kind: 'text', value: 'after.' },
      ],
      resolverFor({ rId9: 'blob:atlas/floating' }),
    );

    const line = container.querySelector('.docx-page__line') as HTMLElement;
    expect(line.style.paddingTop).toBe(''); // no drawingClearancePt reserved
    expect(container.querySelector('.docx-drawing--anchor-marker')).not.toBeNull();
  });

  it('renders a behindDoc float before the text layer and a normal float after it (z-order)', async () => {
    const { container } = await renderPage(
      [
        createDrawing({
          layout: 'anchor',
          relationshipId: 'rId-behind',
          positionH: { relativeFrom: 'page', offsetEmu: 0 },
          positionV: { relativeFrom: 'page', offsetEmu: 0 },
          behindDoc: true,
        }),
      ],
      resolverFor({ 'rId-behind': 'blob:atlas/behind' }),
    );

    const float = container.querySelector('.docx-anchored-drawing') as HTMLElement;
    expect(float.style.zIndex).toBe('0');
    expect(float.style.pointerEvents).toBe('none');

    const { container: frontContainer } = await renderPage(
      [
        createDrawing({
          layout: 'anchor',
          relationshipId: 'rId-front',
          positionH: { relativeFrom: 'page', offsetEmu: 0 },
          positionV: { relativeFrom: 'page', offsetEmu: 0 },
          behindDoc: false,
        }),
      ],
      resolverFor({ 'rId-front': 'blob:atlas/front' }),
    );

    const frontFloat = frontContainer.querySelector('.docx-anchored-drawing') as HTMLElement;
    expect(frontFloat.style.zIndex).toBe('2');
  });

  it('reserves clearance above the text so tall pictures do not overlap earlier lines', async () => {
    const { container } = await renderPage(
      [createDrawing()],
      resolverFor({ rId5: 'blob:atlas/chart' }),
    );

    const line = container.querySelector('.docx-drawing')?.closest('.docx-page__line') as HTMLElement;
    const lineHeight = parseFloat(line.style.height);
    const clearance = parseFloat(line.style.paddingTop);

    expect(lineHeight).toBeGreaterThanOrEqual(90);
    expect(clearance).toBeGreaterThan(0);
    expect(line.style.boxSizing).toBe('border-box');
    expect(parseFloat(line.style.lineHeight)).toBeCloseTo(lineHeight - clearance, 5);
  });

  it('does not double-count the top margin when positioning a line inside its column (regression)', async () => {
    // `.docx-page__column` renders at `top: marginsPt.top`; a line's own
    // `top` inside it must be column-relative, not `PageLineRef.topPt`
    // as-is (which is already page-absolute) — otherwise every line on
    // every page renders `marginsPt.top` further down than it should.
    // Found while adding `floats.ts` (D4/DXL-03), which renders page-
    // absolute floats directly under the page layer and so exposed the
    // inconsistency between the horizontal (`leftPt - col.leftPt`,
    // correctly column-relative) and vertical axes.
    const topMarginPt = 72;
    const section: Section = {
      kind: 'section',
      props: {
        pgSz: { w: twip(600 * 20), h: twip(800 * 20) },
        pgMar: {
          top: twip(topMarginPt * 20),
          right: twip(0),
          bottom: twip(0),
          left: twip(0),
          header: twip(0),
          footer: twip(0),
          gutter: twip(0),
        },
        cols: { num: 1, space: twip(0), col: [] },
      },
      blocks: [
        {
          kind: 'paragraph',
          props: {},
          children: [{ kind: 'run', props: {}, children: [{ kind: 'text', value: 'hi' }] }],
        },
      ],
    };
    const doc: Document = {
      kind: 'document',
      sections: [section],
      styles: new Map(),
      numbering: new Map(),
      headers: new Map(),
      footers: new Map(),
      comments: new Map(),
      footnotes: new Map(),
      endnotes: new Map(),
    } as Document;
    const pages = await paginate({ document: doc, fontResolver: mockFontResolver });

    const { container } = render(
      <MediaContext.Provider value={resolverFor({})}>
        <PageView page={pages[0]} zoom={1} document={doc} />
      </MediaContext.Provider>,
    );

    const column = container.querySelector('.docx-page__column') as HTMLElement;
    const line = container.querySelector('.docx-page__line') as HTMLElement;

    expect(column.style.top).toBe(`${topMarginPt}px`);
    // The first line sits at the column's own top edge (0), not at
    // `topMarginPt` again.
    expect(line.style.top).toBe('0px');
  });
});
