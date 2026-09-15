/**
 * The shared slide model produced by `src/viewers/slides/pptx/parser.ts` and
 * `src/viewers/slides/odp/parser.ts`, and rendered by `SlideDeck`.
 *
 * This is intentionally format-agnostic: both parsers resolve masters/layouts,
 * theme colors, run formatting, bullets, tables, group transforms, and fills
 * down to this one flat `SlideShape[]` list per slide, so a later per-slide
 * export (Task X1) can render a deck off-screen without touching PPTX/ODP
 * XML again.
 */

export type SlideTextRun = {
  readonly text: string
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strikethrough?: boolean
  /** Resolved CSS color (theme/scheme colors already applied). */
  readonly color?: string
  readonly fontSizePx?: number
  readonly fontFamily?: string
}

export type SlideBullet = {
  /** 0-based nesting level. */
  readonly level: number
  readonly char?: string
  readonly numbered?: boolean
}

export type SlideParagraphAlign = 'left' | 'center' | 'right' | 'justify'

export type SlideParagraph = {
  readonly runs: ReadonlyArray<SlideTextRun>
  readonly level: number
  readonly bullet?: SlideBullet
  readonly align?: SlideParagraphAlign
  readonly spaceBeforePx?: number
  readonly spaceAfterPx?: number
}

export type SlideFill =
  | { readonly kind: 'solid'; readonly color: string }
  | { readonly kind: 'gradient'; readonly colors: ReadonlyArray<string>; readonly angleDeg?: number }
  | { readonly kind: 'none' }

export type SlideBorder = {
  readonly color: string
  readonly widthPx: number
}

export type SlideGeometry =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow'
  | 'line'
  | 'other'

export type SlideTransform = {
  readonly x: number
  readonly y: number
  readonly w: number
  readonly h: number
  readonly rotationDeg?: number
  readonly flipH?: boolean
  readonly flipV?: boolean
}

type SlideShapeBase = {
  readonly id: string
  readonly transform: SlideTransform
}

export type SlideTextBox = SlideShapeBase & {
  readonly kind: 'text'
  readonly paragraphs: ReadonlyArray<SlideParagraph>
  /** Flattened plain text — used for search, alt/nav labels, and tests. */
  readonly text: string
  readonly fill?: SlideFill
  readonly border?: SlideBorder
  readonly geometry?: SlideGeometry
  /** `a:normAutofit`/ODF equivalent — a multiplier applied on top of the shape's own font sizes. */
  readonly fontScale?: number
  readonly placeholderType?: string
  /** USR-15 — `a:bodyPr` text layout, resolved through the layout/master placeholder chain. */
  readonly body?: SlideTextBody
}

export type SlideTextAnchor = 'top' | 'middle' | 'bottom'

export type SlideTextBody = {
  /** Inner margins in px (OOXML defaults: 0.1in left/right, 0.05in top/bottom). */
  readonly insets: { readonly top: number; readonly right: number; readonly bottom: number; readonly left: number }
  readonly anchor: SlideTextAnchor
  /** `wrap="none"` keeps each paragraph on one line. */
  readonly wrap: boolean
}

export type SlideShapeOnly = SlideShapeBase & {
  readonly kind: 'shape'
  readonly fill?: SlideFill
  readonly border?: SlideBorder
  readonly geometry?: SlideGeometry
}

export type SlideImage = SlideShapeBase & {
  readonly kind: 'image'
  readonly src: string
  readonly alt: string
  /** Crop fractions (0..1) from each edge, already resolved from `a:srcRect`/ODF crop. */
  readonly crop?: { readonly top: number; readonly right: number; readonly bottom: number; readonly left: number }
}

export type SlideTableCell = {
  readonly text: string
  readonly runs: ReadonlyArray<SlideTextRun>
}

export type SlideTable = SlideShapeBase & {
  readonly kind: 'table'
  readonly rows: ReadonlyArray<ReadonlyArray<SlideTableCell>>
}

export type SlideUnsupported = SlideShapeBase & {
  readonly kind: 'unsupported'
  /** e.g. "Chart not supported", "SmartArt not supported", "Embedded object not supported". */
  readonly label: string
}

export type SlideShape = SlideTextBox | SlideShapeOnly | SlideImage | SlideTable | SlideUnsupported

export type SlideData = {
  readonly id: string
  readonly index: number
  /** Extracted from a title placeholder; falls back to "Slide N" in the UI. */
  readonly title?: string
  readonly hidden?: boolean
  readonly notes?: string
  readonly width: number
  readonly height: number
  readonly background?: SlideFill
  readonly shapes: ReadonlyArray<SlideShape>
  /** Set when this slide's own parse failed — the deck still renders every other slide (S9). */
  readonly error?: string
}
