/**
 * Atlas — immutable DOCX document AST (Wave A.2)
 *
 * This model is the shared contract between the hand-coded parser, style
 * resolver, renderer, editor, and serializer.
 */

import type {
  DocDefaults,
  InsetSet,
  JustifyContent,
  NumberingDef,
  OnOff,
  ParaProps,
  RunProps,
  Shading,
  TableCellVerticalAlign,
  TableLayout,
  TableLook,
  TableRowHeight,
  Twip,
  Width,
  BorderSet,
  Style,
} from './styles'

export type { Style, NumberingDef }

// ---------------------------------------------------------------------------
// Kind constants
// ---------------------------------------------------------------------------

export const BlockKind = {
  paragraph: 'paragraph',
  table: 'table',
  unknown: 'unknown',
} as const

export const InlineKind = {
  run: 'run',
  hyperlink: 'hyperlink',
  bookmark: 'bookmark',
  commentRange: 'comment-range',
  commentReference: 'comment-reference',
  footnoteReference: 'footnote-reference',
  endnoteReference: 'endnote-reference',
  drawing: 'drawing',
  field: 'field',
  text: 'text',
  tab: 'tab',
  break: 'break',
  unknown: 'unknown',
} as const

// ---------------------------------------------------------------------------
// Shared nodes
// ---------------------------------------------------------------------------

export interface UnknownNode {
  readonly kind: 'unknown'
  readonly xml: string
}

export interface TextNode {
  readonly kind: 'text'
  readonly value: string
  readonly preserveSpace?: OnOff
}

export interface TabNode {
  readonly kind: 'tab'
}

export type BreakType = 'line' | 'page' | 'column' | 'textWrapping'

export type BreakClear = 'none' | 'left' | 'right' | 'all'

export interface BreakNode {
  readonly kind: 'break'
  readonly breakType?: BreakType
  readonly clear?: BreakClear
}

export interface Bookmark {
  readonly kind: 'bookmark'
  readonly id: string
  readonly boundary: 'start' | 'end'
  readonly name?: string
  readonly colFirst?: number
  readonly colLast?: number
}

export interface CommentRange {
  readonly kind: 'comment-range'
  readonly id: string
  readonly boundary: 'start' | 'end'
}

export interface CommentReference {
  readonly kind: 'comment-reference'
  readonly id: string
}

export interface FootnoteReference {
  readonly kind: 'footnote-reference'
  readonly id: string
  readonly customMarkFollows?: OnOff
}

export interface EndnoteReference {
  readonly kind: 'endnote-reference'
  readonly id: string
  readonly customMarkFollows?: OnOff
}

export type DrawingLayout = 'inline' | 'anchor'

export interface DrawingExtent {
  readonly cx: number
  readonly cy: number
}

/**
 * `wp:effectExtent`'s l/t/r/b (EMU) — extra bleed around the drawing's
 * `extent` box reserved for effects (shadow, glow, rotation overshoot) that
 * extend past the nominal picture rectangle. Valid on both `wp:inline` and
 * `wp:anchor` (D4/DXP-09).
 */
export interface DrawingEffectExtent {
  readonly l: number
  readonly t: number
  readonly r: number
  readonly b: number
}

/**
 * `wp:positionH`'s relativeFrom values (ST_RelFromH). `column`/`character`
 * only make sense for positionH — see `DrawingRelativeFromV` for positionV's
 * distinct set.
 */
export type DrawingRelativeFromH =
  | 'page'
  | 'margin'
  | 'column'
  | 'character'
  | 'leftMargin'
  | 'rightMargin'
  | 'insideMargin'
  | 'outsideMargin'

/** `wp:positionV`'s relativeFrom values (ST_RelFromV). */
export type DrawingRelativeFromV =
  | 'page'
  | 'margin'
  | 'paragraph'
  | 'line'
  | 'topMargin'
  | 'bottomMargin'
  | 'insideMargin'
  | 'outsideMargin'

export type DrawingHorizontalAlign = 'left' | 'center' | 'right' | 'inside' | 'outside'
export type DrawingVerticalAlign = 'top' | 'center' | 'bottom' | 'inside' | 'outside'

/**
 * `wp:positionH`, fully parsed (DXP-09/DXL-03) rather than captured as raw
 * XML: either `align` (`wp:align`) or `offsetEmu` (`wp:posOffset`) is present,
 * matching the mutually-exclusive `EG_WrapPositionType` choice in the OOXML
 * schema — never both, and a schema-valid document always has one.
 */
export interface DrawingPositionH {
  readonly relativeFrom: DrawingRelativeFromH
  readonly align?: DrawingHorizontalAlign
  readonly offsetEmu?: number
}

/** `wp:positionV`, the vertical counterpart of `DrawingPositionH`. */
export interface DrawingPositionV {
  readonly relativeFrom: DrawingRelativeFromV
  readonly align?: DrawingVerticalAlign
  readonly offsetEmu?: number
}

export type DrawingWrapMode = 'none' | 'square' | 'tight' | 'through' | 'topAndBottom'

/** `wrapText` attribute shared by `wp:wrapSquare`/`wrapTight`/`wrapThrough`. */
export type DrawingWrapSide = 'bothSides' | 'left' | 'right' | 'largest'

/**
 * The anchor's wrap choice (`wp:wrapNone`/`wp:wrapSquare`/`wp:wrapTight`/
 * `wp:wrapThrough`/`wp:wrapTopAndBottom`), parsed into one shape regardless
 * of which element was present. `distT/B/L/R` (EMU) are that element's own
 * distance-from-text attributes (absent on `wrapNone`/`wrapThrough`, which
 * don't carry them).
 */
export interface DrawingWrap {
  readonly mode: DrawingWrapMode
  readonly side?: DrawingWrapSide
  readonly distTEmu?: number
  readonly distBEmu?: number
  readonly distLEmu?: number
  readonly distREmu?: number
}

/**
 * `a:srcRect`'s crop rectangle (DXS-09): each edge is in thousandths of a
 * percent of the source image's full width/height (0-100000 for a normal
 * crop; the OOXML spec also allows negative values to pad rather than crop).
 * A field's absence means that edge's attribute was absent in the source
 * (defaults to 0 uncropped per spec) — preserved distinctly from `0` so a
 * save doesn't add attributes the source never had.
 */
export interface DrawingCrop {
  readonly l?: number
  readonly t?: number
  readonly r?: number
  readonly b?: number
}

/**
 * The picture's own rotation/flip (DXS-09), from `pic:spPr/a:xfrm`'s `rot`
 * (60,000ths of a degree, clockwise) and `flipH`/`flipV` attributes. Only
 * populated when at least one of the three is present in the source — an
 * `a:xfrm` present solely for its (unmodeled) `a:off`/`a:ext` children is
 * treated the same as no `a:xfrm` at all, since Atlas derives position from
 * `wp:positionH`/`V` and size from `wp:extent` instead.
 */
export interface DrawingTransform {
  readonly rotation?: number
  readonly flipH?: boolean
  readonly flipV?: boolean
}

/**
 * Marks where one of Atlas's own modeled `wp:anchor` children sits among the
 * raw, unmodeled ones captured in `Drawing.anchorChildren`, so the serializer
 * can rebuild each slot from current model state instead of replaying stale
 * captured XML. `positionH`/`positionV`/`wrap`/`effectExtent` (DXP-09) join
 * the original `extent`/`docPr`/`graphic` slots; anything else Atlas doesn't
 * model — `wp:simplePos`, `wp:cNvGraphicFramePr` — is still replayed
 * verbatim as an `UnknownNode` in `anchorChildren`.
 */
export interface DrawingAnchorSlot {
  readonly kind: 'anchor-slot'
  readonly slot: 'extent' | 'effectExtent' | 'positionH' | 'positionV' | 'wrap' | 'docPr' | 'graphic'
}

export type DrawingAnchorChild = DrawingAnchorSlot | UnknownNode

export interface Drawing {
  readonly kind: 'drawing'
  readonly layout: DrawingLayout
  readonly relationshipId?: string
  readonly title?: string
  readonly description?: string
  readonly name?: string
  readonly extent?: DrawingExtent
  /** Valid for both `inline` and `anchor` (ST schema allows it on either). */
  readonly effectExtent?: DrawingEffectExtent
  /** `a:srcRect` crop, from the picture subtree — valid regardless of `layout` (DXS-09). */
  readonly crop?: DrawingCrop
  /** `a:xfrm` rotation/flip, from the picture subtree — valid regardless of `layout` (DXS-09). */
  readonly transform?: DrawingTransform
  /**
   * The following fields are populated for `layout: 'anchor'` only (parsed
   * straight off `wp:anchor`'s own attributes/children rather than captured
   * via `anchorChildren`, so layout/wrap logic can read them directly
   * without walking the raw child list — DXP-09/DXL-03).
   */
  readonly behindDoc?: boolean
  readonly allowOverlap?: boolean
  readonly positionH?: DrawingPositionH
  readonly positionV?: DrawingPositionV
  readonly wrap?: DrawingWrap
  /**
   * For `layout: 'anchor'` drawings only: the exact original child order of
   * `wp:anchor`, required so a save emits the schema-required position/wrap
   * elements (DXS-03) instead of silently dropping them. Absent for
   * `layout: 'inline'` drawings, which carry no position/wrap and are
   * rebuilt from `extent`/`title`/`description`/`name`/`relationshipId`.
   */
  readonly anchorChildren?: ReadonlyArray<DrawingAnchorChild>
}

export type RunChild =
  | TextNode
  | TabNode
  | BreakNode
  | Drawing
  | CommentReference
  | FootnoteReference
  | EndnoteReference
  | UnknownNode

export interface Run {
  readonly kind: 'run'
  readonly props?: RunProps
  readonly children: ReadonlyArray<RunChild>
  /**
   * Revision-save-ID bookkeeping attributes (`w:rsidR`/`w:rsidRPr`/`w:rsidDel`)
   * captured verbatim from a Word-authored `<w:r>` so a save doesn't strip
   * them (DXS-10). Word uses these to correlate edits across a document's
   * revision-save history; Atlas never generates or interprets them itself.
   */
  readonly rsidR?: string
  readonly rsidRPr?: string
  readonly rsidDel?: string
}

/**
 * DEFER-5 / DXS-20 — a `w:fldSimple` or complex (`w:fldChar` begin/
 * separate/end + `w:instrText`) field, modeled uniformly regardless of
 * which XML form the source used. `result` is the cached display content
 * Word last computed for it (what a viewer that never recalculates fields
 * should show) — recalculating is an explicit, opt-in "Update field(s)"
 * action (`src/docx/fields`), never automatic on load, matching Word's own
 * behavior of trusting a field's last-known-good result until asked
 * otherwise.
 */
export type FieldType =
  | 'DATE'
  | 'TIME'
  | 'AUTHOR'
  | 'TITLE'
  | 'REF'
  | 'PAGEREF'
  | 'SEQ'
  | 'NUMPAGES'
  | 'PAGE'
  | 'HYPERLINK'
  | 'TOC'
  | 'unknown'

export interface Field {
  readonly kind: 'field'
  /** The instruction's first keyword, uppercased (e.g. `DATE`, `REF`), or `unknown` for a field type Atlas doesn't evaluate. */
  readonly fieldType: FieldType
  /** The exact field instruction text (switches included), e.g. `PAGE \* MERGEFORMAT` or `REF _Ref123 \h`. */
  readonly instruction: string
  /** Cached display content — shown as-is until "Update field(s)" regenerates it. */
  readonly result: ReadonlyArray<ParagraphChild>
  /** `true` for the `w:fldSimple` form; absent/`false` for the complex `fldChar`-based form. Both round-trip byte-faithfully via `raw` until updated. */
  readonly simple?: boolean
  readonly locked?: OnOff
  readonly dirty?: OnOff
  /**
   * The exact source XML for this field (the `w:fldSimple` element, or the
   * run sequence from the begin `w:fldChar` through the end `w:fldChar`
   * inclusive), captured the same way `UnknownNode` captures raw source
   * text (D19 / DXS-16). The serializer re-emits this verbatim as long as
   * the field hasn't been regenerated by "Update field(s)"; `undefined`
   * only for a field built in memory rather than parsed from a source file.
   */
  readonly raw?: string
}

export type HyperlinkChild =
  | Run
  | Bookmark
  | CommentRange
  | CommentReference
  | FootnoteReference
  | EndnoteReference
  | Field
  | UnknownNode

export interface Hyperlink {
  readonly kind: 'hyperlink'
  readonly relationshipId?: string
  readonly anchor?: string
  readonly tooltip?: string
  readonly targetFrame?: string
  readonly history?: OnOff
  readonly children: ReadonlyArray<HyperlinkChild>
}

export type Inline =
  | Run
  | Hyperlink
  | Bookmark
  | CommentRange
  | CommentReference
  | FootnoteReference
  | EndnoteReference
  | Drawing
  | Field
  | TextNode
  | TabNode
  | BreakNode
  | UnknownNode

export interface InsRevision {
  readonly kind: 'ins-revision'
  readonly id: string
  readonly author?: string
  readonly date?: string
  readonly children: ReadonlyArray<ParagraphChild> & ReadonlyArray<Run>
}

export interface DelRevision {
  readonly kind: 'del-revision'
  readonly id: string
  readonly author?: string
  readonly date?: string
  readonly children: ReadonlyArray<ParagraphChild> & ReadonlyArray<Run>
}

export type ParagraphChild =
  | Run
  | Hyperlink
  | Bookmark
  | CommentRange
  | CommentReference
  | FootnoteReference
  | EndnoteReference
  | InsRevision
  | DelRevision
  | Field
  | UnknownNode

// ---------------------------------------------------------------------------
// Section props
// ---------------------------------------------------------------------------

export type PageOrientation = 'portrait' | 'landscape'

export interface PageSize {
  readonly w: Twip
  readonly h: Twip
  readonly orient?: PageOrientation
}

export interface PageMargins {
  readonly top?: Twip
  readonly right?: Twip
  readonly bottom?: Twip
  readonly left?: Twip
  readonly header?: Twip
  readonly footer?: Twip
  readonly gutter?: Twip
}

export interface SectionColumn {
  readonly w?: Twip
  readonly space?: Twip
}

export interface SectionColumns {
  readonly num?: number
  readonly space?: Twip
  readonly sep?: OnOff
  readonly equalWidth?: OnOff
  readonly col: ReadonlyArray<SectionColumn>
}

export interface PageNumberType {
  readonly start?: number
  readonly fmt?: string
}

export type HeaderFooterReferenceType = 'default' | 'first' | 'even'

export interface HeaderReference {
  readonly id: string
  readonly type: HeaderFooterReferenceType
}

export interface FooterReference {
  readonly id: string
  readonly type: HeaderFooterReferenceType
}

export type SectionBreakType =
  | 'continuous'
  | 'nextPage'
  | 'nextColumn'
  | 'evenPage'
  | 'oddPage'

export interface LineNumberType {
  readonly countBy?: number
  readonly start?: number
  readonly distance?: Twip
  readonly restart?: 'continuous' | 'newPage' | 'newSection'
}

export type SectionVerticalAlign = 'top' | 'center' | 'both' | 'bottom'

export interface SectionProps {
  readonly pgSz?: PageSize
  readonly pgMar?: PageMargins
  readonly cols?: SectionColumns
  readonly pgNumType?: PageNumberType
  readonly titlePg?: OnOff
  readonly type?: SectionBreakType
  readonly headerReference?: ReadonlyArray<HeaderReference>
  readonly footerReference?: ReadonlyArray<FooterReference>
  readonly lnNumType?: LineNumberType
  readonly vAlign?: SectionVerticalAlign
}

// ---------------------------------------------------------------------------
// Blocks + table nodes
// ---------------------------------------------------------------------------

export interface Paragraph {
  readonly kind: 'paragraph'
  readonly props?: ParaProps
  readonly children: ReadonlyArray<ParagraphChild>
  /**
   * `w14:paraId`/`w14:textId` and `w:rsid*` bookkeeping attributes captured
   * verbatim from a Word-authored `<w:p>` so a save doesn't strip them
   * (DXS-10). `paraId` in particular is the join key `word/commentsExtended.xml`
   * uses to correlate a comment with its resolved/done state (D16/DXS-11) —
   * losing it on save would silently detach that state from the comment.
   * Atlas never generates these for a paragraph that already has them;
   * genuinely new paragraphs are left without them, matching how Word
   * itself only assigns them lazily.
   */
  readonly paraId?: string
  readonly textId?: string
  readonly rsidR?: string
  readonly rsidRDefault?: string
  readonly rsidP?: string
  readonly rsidRPr?: string
}

export interface TableProps {
  readonly tblStyle?: string
  readonly tblW?: Width
  readonly tblInd?: Width
  readonly tblBorders?: BorderSet
  readonly tblCellMar?: InsetSet
  readonly tblLayout?: TableLayout
  readonly tblLook?: TableLook
  readonly jc?: JustifyContent
  readonly shd?: Shading
}

export interface TableRowProps {
  readonly trHeight?: TableRowHeight
  readonly cantSplit?: OnOff
  readonly tblHeader?: OnOff
  readonly jc?: JustifyContent
}

export type TableCellMerge = 'restart' | 'continue'

export interface TableCellProps {
  readonly tcW?: Width
  readonly gridSpan?: number
  readonly vMerge?: TableCellMerge
  readonly tcBorders?: BorderSet
  readonly shd?: Shading
  readonly tcMar?: InsetSet
  readonly vAlign?: TableCellVerticalAlign
  readonly noWrap?: OnOff
  readonly hideMark?: OnOff
}

export interface TableCell {
  readonly kind: 'table-cell'
  readonly props?: TableCellProps
  readonly blocks: ReadonlyArray<Block>
}

export type TableRowChild = TableCell | UnknownNode

export interface TableRow {
  readonly kind: 'table-row'
  readonly props?: TableRowProps
  readonly cells: ReadonlyArray<TableRowChild>
}

export type TableChild = TableRow | UnknownNode

export interface Table {
  readonly kind: 'table'
  readonly props?: TableProps
  /**
   * Column widths from `w:tblGrid` (one entry per `w:gridCol`, in twips).
   * Required by the OOXML schema on every `w:tbl`; layout consumes this
   * directly for `w:tblLayout="fixed"` tables instead of measuring content.
   */
  readonly tblGrid?: ReadonlyArray<Twip>
  readonly rows: ReadonlyArray<TableChild>
}

export type Block = Paragraph | Table | UnknownNode

// ---------------------------------------------------------------------------
// Notes, headers, footers, and document root
// ---------------------------------------------------------------------------

export interface Section {
  readonly kind: 'section'
  readonly props: SectionProps
  readonly blocks: ReadonlyArray<Block>
}

export interface Comment {
  readonly kind: 'comment'
  readonly id: string
  readonly author?: string
  readonly initials?: string
  readonly date?: string
  /**
   * Comment body blocks in source order. Paragraphs are overwhelmingly the
   * common case, but a comment can legitimately contain a table too — see
   * the wave 1 follow-up that stopped `parser/comments.ts` from silently
   * discarding one.
   */
  readonly body: ReadonlyArray<Block>
  readonly parentId?: string
  /**
   * Whether the comment thread is marked resolved, sourced from
   * `word/commentsExtended.xml`'s `w15:done` attribute (D16/DXS-11).
   * `undefined` when the document has no `commentsExtended.xml` part at
   * all (most comment-bearing documents don't) — treat as "not resolved"
   * for display purposes, distinct from an explicit `false`.
   */
  readonly resolved?: boolean
}

export type NoteType = 'normal' | 'separator' | 'continuationSeparator' | 'continuationNotice'

export interface Footnote {
  readonly kind: 'footnote'
  readonly id: string
  readonly noteType?: NoteType
  readonly blocks: ReadonlyArray<Block>
}

export interface Endnote {
  readonly kind: 'endnote'
  readonly id: string
  readonly noteType?: NoteType
  readonly blocks: ReadonlyArray<Block>
}

export interface Header {
  readonly kind: 'header'
  readonly id: string
  readonly blocks: ReadonlyArray<Block>
}

export interface Footer {
  readonly kind: 'footer'
  readonly id: string
  readonly blocks: ReadonlyArray<Block>
}

export interface Document {
  readonly kind: 'document'
  readonly sections: ReadonlyArray<Section>
  readonly styles: ReadonlyMap<string, Style>
  readonly numbering: ReadonlyMap<string, NumberingDef>
  readonly defaults?: DocDefaults
  readonly comments: ReadonlyMap<string, Comment>
  readonly footnotes: ReadonlyMap<string, Footnote>
  readonly endnotes: ReadonlyMap<string, Endnote>
  readonly headers: ReadonlyMap<string, Header>
  readonly footers: ReadonlyMap<string, Footer>
  /**
   * `{prefix: uri}` namespace declarations the source `<w:document>` root
   * element itself carried (D19 / DXS-15). `undefined` when the document
   * was constructed in memory rather than parsed from a source file.
   * `documentWriter.ts` unions this with Atlas's own required baseline
   * namespace set when re-emitting the root, instead of emitting a fixed
   * hardcoded set regardless of what the source actually declared — this
   * preserves any namespace prefix a source document declares that Atlas's
   * own serializer doesn't otherwise know about (most relevantly one used
   * only by unknown-node passthrough content).
   */
  readonly rootNamespaces?: ReadonlyMap<string, string>
  /**
   * The source `<w:document>` root's `mc:Ignorable` attribute value
   * (space-separated namespace prefixes), unioned with Atlas's own
   * baseline token list on save for the same reason as `rootNamespaces`.
   */
  readonly mcIgnorable?: string
}
