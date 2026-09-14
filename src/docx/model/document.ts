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
 * Marks where one of Atlas's own modeled `wp:anchor` children (extent/docPr/
 * graphic) sits among the raw, unmodeled ones captured in `Drawing.anchorChildren`,
 * so the serializer can rebuild each slot from current model state instead of
 * replaying stale captured XML, while still emitting position/wrap children
 * (`wp:simplePos`, `wp:positionH`/`wp:positionV`, the wrap choice,
 * `wp:effectExtent`, `wp:cNvGraphicFramePr`) verbatim in their original spot.
 */
export interface DrawingAnchorSlot {
  readonly kind: 'anchor-slot'
  readonly slot: 'extent' | 'docPr' | 'graphic'
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

export type HyperlinkChild =
  | Run
  | Bookmark
  | CommentRange
  | CommentReference
  | FootnoteReference
  | EndnoteReference
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
