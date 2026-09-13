import React, { useState } from 'react';
import type { Drawing } from '../model/document';
import { useMediaResolver } from './mediaContext';

// The editor maps DOM selections to model offsets by counting text characters
// inside `[data-run-index]` elements; a drawing occupies exactly one offset.
const DRAWING_OFFSET_CHAR = '￼';

export type InlineDrawingProps = {
  drawing: Drawing;
  widthPt: number;
  heightPt: number;
  runIndex: number;
  charStart: number;
  charEnd: number;
  className?: string;
};

/**
 * Paints a laid-out drawing into its reserved slot on a page line.
 *
 * The slot is a zero-height inline-block sitting on the text baseline, so it
 * never inflates the browser's line box; the picture is absolutely positioned
 * to grow upward from the baseline. The paginator has already reserved that
 * space via `LineBox.drawingClearancePt`.
 *
 * Coordinates are internal "px" units inside the PageView scale wrapper.
 */
export function InlineDrawing({
  drawing,
  widthPt,
  heightPt,
  runIndex,
  charStart,
  charEnd,
  className,
}: InlineDrawingProps): React.ReactElement {
  const media = useMediaResolver();
  const [failedUrl, setFailedUrl] = useState<string | null>(null);

  const url = drawing.relationshipId !== undefined ? media.resolve(drawing.relationshipId) : null;
  const canPaint = url !== null && url !== failedUrl;
  const label = drawing.description ?? drawing.title ?? drawing.name ?? '';

  const boxStyle: React.CSSProperties = {
    position: 'absolute',
    left: 0,
    bottom: 0,
    width: `${widthPt}px`,
    height: `${heightPt}px`,
  };

  return (
    <span
      className={className === undefined ? 'docx-drawing' : `docx-drawing ${className}`}
      data-run-index={runIndex}
      data-char-start={charStart}
      data-char-end={charEnd}
      data-drawing-layout={drawing.layout}
      style={{
        display: 'inline-block',
        position: 'relative',
        width: `${widthPt}px`,
        height: 0,
        verticalAlign: 'baseline',
      }}
    >
      {canPaint ? (
        <img
          className="docx-drawing__image"
          src={url}
          alt={label}
          draggable={false}
          onError={() => setFailedUrl(url)}
          style={boxStyle}
        />
      ) : (
        <span
          className="docx-drawing__placeholder"
          role="img"
          aria-label={label === '' ? 'Image unavailable' : label}
          title={label === '' ? 'Image unavailable' : `Image unavailable: ${label}`}
          style={boxStyle}
        />
      )}
      {/* Out of flow so it cannot give the zero-height slot a line box
          (which would move the slot's baseline off the text baseline). */}
      <span
        className="docx-drawing__offset"
        aria-hidden="true"
        style={{ position: 'absolute', width: 0, height: 0, overflow: 'hidden', opacity: 0 }}
      >
        {DRAWING_OFFSET_CHAR}
      </span>
    </span>
  );
}
