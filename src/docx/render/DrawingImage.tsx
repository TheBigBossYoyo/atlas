import React, { useState } from 'react';
import type { Drawing } from '../model/document';
import { useMediaResolver } from './mediaContext';

const CROP_UNIT_FULL = 100 * 1000; // a:srcRect l/t/r/b are in 1000ths of a percent
const ANGLE_UNITS_PER_DEGREE = 60000; // a:xfrm rot is in 60,000ths of a degree
// Guards a pathological/malformed crop (>=100% removed from opposing edges)
// from producing a zero/negative/Infinity-sized inner image.
const MIN_REMAINING_FRACTION = 0.01;

export type DrawingImageProps = {
  drawing: Drawing;
  widthPt: number;
  heightPt: number;
  /**
   * Absolute positioning for the rendered box, supplied by the caller:
   * `InlineDrawing` anchors it to its own zero-height baseline slot
   * (`{ position: 'absolute', left: 0, bottom: 0 }`), `AnchoredDrawing`
   * anchors it to the page layer's computed float rect. `DrawingImage` only
   * owns size/crop/rotation, never where on the page that box sits.
   */
  positionStyle: React.CSSProperties;
};

/**
 * Renders one drawing's picture (or its unavailable-image placeholder),
 * applying `a:srcRect` crop and `a:xfrm` rotation/flip (DXS-09) — shared by
 * `InlineDrawing` and `AnchoredDrawing` so the two layout paths (in-flow vs.
 * page-absolute float) don't duplicate this logic.
 *
 * Crop is applied in the picture's own local, unrotated coordinate frame,
 * matching Word: the visible box is sized to `widthPt`×`heightPt` (the
 * *cropped*, displayed size — `wp:extent`), an inner `<img>` is scaled up so
 * the region outside the crop rectangle is pushed off the visible box's
 * edges, then the whole visible (cropped) box is rotated/flipped about its
 * own center. When there's no crop, a single `<img>` is rendered directly
 * (no wrapper) — the common case keeps the exact DOM shape this component
 * had before DXS-09, so existing consumers/tests are unaffected.
 */
export function DrawingImage({ drawing, widthPt, heightPt, positionStyle }: DrawingImageProps): React.ReactElement {
  const media = useMediaResolver();
  const [failedUrl, setFailedUrl] = useState<string | null>(null);

  const url = drawing.relationshipId !== undefined ? media.resolve(drawing.relationshipId) : null;
  const canPaint = url !== null && url !== failedUrl;
  const label = drawing.description ?? drawing.title ?? drawing.name ?? '';

  const leftFrac = cropFraction(drawing.crop?.l);
  const topFrac = cropFraction(drawing.crop?.t);
  const rightFrac = cropFraction(drawing.crop?.r);
  const bottomFrac = cropFraction(drawing.crop?.b);
  const hasCrop = leftFrac > 0 || topFrac > 0 || rightFrac > 0 || bottomFrac > 0;

  const transform = buildTransformCss(drawing.transform);
  const boxStyle: React.CSSProperties = {
    ...positionStyle,
    width: `${widthPt}px`,
    height: `${heightPt}px`,
    ...(transform !== undefined ? { transform, transformOrigin: 'center' } : {}),
  };

  if (!canPaint) {
    return (
      <span
        className="docx-drawing__placeholder"
        role="img"
        aria-label={label === '' ? 'Image unavailable' : label}
        title={label === '' ? 'Image unavailable' : `Image unavailable: ${label}`}
        style={boxStyle}
      />
    );
  }

  if (!hasCrop) {
    return (
      <img
        className="docx-drawing__image"
        src={url}
        alt={label}
        draggable={false}
        onError={() => setFailedUrl(url)}
        style={boxStyle}
      />
    );
  }

  const remainingWidthFrac = Math.max(MIN_REMAINING_FRACTION, 1 - leftFrac - rightFrac);
  const remainingHeightFrac = Math.max(MIN_REMAINING_FRACTION, 1 - topFrac - bottomFrac);
  const fullWidthPt = widthPt / remainingWidthFrac;
  const fullHeightPt = heightPt / remainingHeightFrac;

  return (
    <span style={{ ...boxStyle, display: 'block', overflow: 'hidden' }}>
      <img
        className="docx-drawing__image"
        src={url}
        alt={label}
        draggable={false}
        onError={() => setFailedUrl(url)}
        style={{
          position: 'absolute',
          left: `${-leftFrac * fullWidthPt}px`,
          top: `${-topFrac * fullHeightPt}px`,
          width: `${fullWidthPt}px`,
          height: `${fullHeightPt}px`,
        }}
      />
    </span>
  );
}

function cropFraction(value: number | undefined): number {
  return value === undefined ? 0 : value / CROP_UNIT_FULL;
}

function buildTransformCss(transform: Drawing['transform']): string | undefined {
  if (transform === undefined) {
    return undefined;
  }

  const parts: string[] = [];
  if (typeof transform.rotation === 'number' && transform.rotation !== 0) {
    parts.push(`rotate(${transform.rotation / ANGLE_UNITS_PER_DEGREE}deg)`);
  }
  if (transform.flipH === true || transform.flipV === true) {
    // Flip happens in the picture's own local frame before rotation — CSS
    // composes `transform` functions right-to-left, so listing `scale`
    // after `rotate` applies it first, matching OOXML's semantics.
    parts.push(`scale(${transform.flipH === true ? -1 : 1}, ${transform.flipV === true ? -1 : 1})`);
  }

  return parts.length > 0 ? parts.join(' ') : undefined;
}
