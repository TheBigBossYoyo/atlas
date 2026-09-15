import React from 'react';
import type { Drawing } from '../model/document';
import { DrawingImage } from './DrawingImage';

// The editor maps DOM selections to model offsets by counting text characters
// inside `[data-run-index]` elements; a drawing occupies exactly one offset.
const DRAWING_OFFSET_CHAR = '￼';

const INLINE_POSITION_STYLE: React.CSSProperties = { position: 'absolute', left: 0, bottom: 0 };

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
      <DrawingImage
        drawing={drawing}
        widthPt={widthPt}
        heightPt={heightPt}
        positionStyle={INLINE_POSITION_STYLE}
      />
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

export type DrawingAnchorMarkerProps = {
  runIndex: number;
  charStart: number;
  charEnd: number;
};

/**
 * The in-flow counterpart of an anchored (floating) drawing (D4/DXL-03).
 *
 * A floating picture's visual is rendered separately, absolutely positioned
 * at the page level (see `AnchoredDrawing`/`floats.ts`) — it does not sit at
 * this position in the text. But the run/paragraph editor's cursor mapping
 * (`editor/Cursor.ts`) walks `[data-run-index]` elements and counts DOM text
 * characters to resolve offsets, so the anchor still needs *some* element
 * here carrying the same one-character offset every other drawing kind
 * occupies. This is that marker: invisible and `aria-hidden` (so a screen
 * reader isn't told about the picture twice — once here, once at the float),
 * carrying no image of its own.
 */
export function DrawingAnchorMarker({
  runIndex,
  charStart,
  charEnd,
}: DrawingAnchorMarkerProps): React.ReactElement {
  return (
    <span
      className="docx-drawing docx-drawing--anchor-marker"
      data-run-index={runIndex}
      data-char-start={charStart}
      data-char-end={charEnd}
      data-drawing-layout="anchor"
      aria-hidden="true"
      style={{ position: 'absolute', width: 0, height: 0, overflow: 'hidden', opacity: 0 }}
    >
      {DRAWING_OFFSET_CHAR}
    </span>
  );
}
