import React from 'react';
import type { PageFloat } from '../layout/floats';
import { DrawingImage } from './DrawingImage';

export type AnchoredDrawingProps = {
  pageFloat: PageFloat;
};

/**
 * Paints one floating (anchored) drawing at its absolute page position
 * (D4/DXL-03). Rendered by `PageView` as a sibling of the page's columns —
 * `behindDoc` floats are placed before the columns in DOM order (so text
 * paints over them), non-`behindDoc` floats after (so they paint over text)
 * — see `PageView.tsx`'s `renderFloats`.
 *
 * `pointer-events: none`: Atlas's editor doesn't yet support dragging a
 * floating image to reposition it, so the float layer stays out of the way
 * of text selection/click handling underneath or above it. The anchor's own
 * character offset (for cursor addressing) is a separate, invisible marker
 * back in the text flow — see `DrawingAnchorMarker`.
 */
export function AnchoredDrawing({ pageFloat }: AnchoredDrawingProps): React.ReactElement {
  const { drawing, rect, behindDoc } = pageFloat;

  return (
    <div
      className="docx-anchored-drawing"
      data-drawing-layout="anchor"
      data-block-index={pageFloat.blockIndex}
      style={{
        position: 'absolute',
        left: `${rect.leftPt}px`,
        top: `${rect.topPt}px`,
        width: `${rect.widthPt}px`,
        height: `${rect.heightPt}px`,
        zIndex: behindDoc ? 0 : 2,
        pointerEvents: 'none',
      }}
    >
      <DrawingImage
        drawing={drawing}
        widthPt={rect.widthPt}
        heightPt={rect.heightPt}
        positionStyle={{ position: 'absolute', left: 0, top: 0 }}
      />
    </div>
  );
}
