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
});
