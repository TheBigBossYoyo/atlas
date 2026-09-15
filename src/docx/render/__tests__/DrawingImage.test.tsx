import { describe, expect, it } from 'vitest';
import { render } from '@testing-library/react';

import { DrawingImage } from '../DrawingImage';
import { MediaContext, type MediaResolver } from '../mediaContext';
import type { Drawing } from '../../model';

function resolverFor(entries: Record<string, string>): MediaResolver {
  return { resolve: (relationshipId) => entries[relationshipId] ?? null };
}

function renderImage(drawing: Drawing, resolver: MediaResolver) {
  return render(
    <MediaContext.Provider value={resolver}>
      <DrawingImage
        drawing={drawing}
        widthPt={100}
        heightPt={50}
        positionStyle={{ position: 'absolute', left: 0, top: 0 }}
      />
    </MediaContext.Provider>,
  );
}

const BASE_DRAWING: Drawing = {
  kind: 'drawing',
  layout: 'inline',
  relationshipId: 'rId1',
};

describe('DrawingImage', () => {
  it('renders a single <img> with no wrapper when there is no crop (fast path)', () => {
    const { container } = renderImage(BASE_DRAWING, resolverFor({ rId1: 'blob:atlas/pic' }));

    const img = container.querySelector('img');
    expect(img).not.toBeNull();
    expect(img).toHaveAttribute('src', 'blob:atlas/pic');
    expect(img?.style.width).toBe('100px');
    expect(img?.style.height).toBe('50px');
    expect(img?.style.left).toBe('0px');
    expect(img?.style.top).toBe('0px');
    // No extra wrapper span around the <img> in the uncropped case.
    expect(container.querySelector('span')).toBeNull();
  });

  it('applies rotation directly to the <img> when there is no crop', () => {
    const drawing: Drawing = { ...BASE_DRAWING, transform: { rotation: 2700000 } }; // 45deg
    const { container } = renderImage(drawing, resolverFor({ rId1: 'blob:atlas/pic' }));

    const img = container.querySelector('img') as HTMLImageElement;
    expect(img.style.transform).toBe('rotate(45deg)');
    expect(img.style.transformOrigin).toBe('center');
  });

  it('composes flip and rotation with scale listed after rotate (flip-then-rotate semantics)', () => {
    const drawing: Drawing = {
      ...BASE_DRAWING,
      transform: { rotation: 5400000, flipH: true, flipV: true }, // 90deg
    };
    const { container } = renderImage(drawing, resolverFor({ rId1: 'blob:atlas/pic' }));

    const img = container.querySelector('img') as HTMLImageElement;
    expect(img.style.transform).toBe('rotate(90deg) scale(-1, -1)');
  });

  it('wraps in an overflow:hidden box and scales the inner <img> when a crop is present', () => {
    // 10% off each edge -> remaining 80% width/height. Displayed box is
    // 100x50pt, so the full (uncropped) image scales to 125x62.5pt, offset
    // left/top by 10% of that full size.
    const drawing: Drawing = { ...BASE_DRAWING, crop: { l: 10000, t: 10000, r: 10000, b: 10000 } };
    const { container } = renderImage(drawing, resolverFor({ rId1: 'blob:atlas/pic' }));

    const outer = container.querySelector('span') as HTMLElement;
    expect(outer).not.toBeNull();
    expect(outer.style.overflow).toBe('hidden');
    expect(outer.style.width).toBe('100px');
    expect(outer.style.height).toBe('50px');

    const img = outer.querySelector('img') as HTMLImageElement;
    expect(img.style.width).toBe('125px');
    expect(img.style.height).toBe('62.5px');
    expect(img.style.left).toBe('-12.5px');
    expect(img.style.top).toBe('-6.25px');
  });

  it('rotates the outer cropped box, not the inner scaled image', () => {
    const drawing: Drawing = {
      ...BASE_DRAWING,
      crop: { l: 10000, t: 0, r: 0, b: 0 },
      transform: { rotation: 2700000 },
    };
    const { container } = renderImage(drawing, resolverFor({ rId1: 'blob:atlas/pic' }));

    const outer = container.querySelector('span') as HTMLElement;
    expect(outer.style.transform).toBe('rotate(45deg)');
    const img = outer.querySelector('img') as HTMLImageElement;
    expect(img.style.transform).toBe('');
  });

  it('treats an all-zero/absent crop the same as no crop at all', () => {
    const drawing: Drawing = { ...BASE_DRAWING, crop: {} };
    const { container } = renderImage(drawing, resolverFor({ rId1: 'blob:atlas/pic' }));

    expect(container.querySelector('span')).toBeNull();
    expect(container.querySelector('img')).not.toBeNull();
  });

  it('shows the placeholder (sized to the display box) when the image cannot be resolved', () => {
    const drawing: Drawing = { ...BASE_DRAWING, crop: { l: 10000, r: 10000, t: 0, b: 0 } };
    const { getByRole } = renderImage(drawing, resolverFor({}));

    const placeholder = getByRole('img', { name: 'Image unavailable' });
    expect(placeholder).toHaveClass('docx-drawing__placeholder');
    expect(placeholder.style.width).toBe('100px');
    expect(placeholder.style.height).toBe('50px');
  });
});
