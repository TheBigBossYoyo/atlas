/**
 * Atlas — immutable DOCX style model (Wave A.2)
 *
 * These types model the OOXML style/property surface consumed by the hand-coded
 * DOCX parser, style cascade resolver, renderer, and editor pipeline.
 */

import type { SectionProps, TableCellProps, TableRowProps } from './document'

// ---------------------------------------------------------------------------
// Branded scalar values
// ---------------------------------------------------------------------------

export type HexColor = string & { readonly __brand: 'HexColor' }
export type Color = HexColor | 'auto'

export type Twip = number & { readonly __brand: 'Twip' }
export type HalfPoint = number & { readonly __brand: 'HalfPoint' }
export type EighthPoint = number & { readonly __brand: 'EighthPoint' }
export type Pct = number & { readonly __brand: 'Pct' }

const HEX_COLOR_DIGITS = /^[0-9a-fA-F]{6}$/

/**
 * DOCX-14 fix — `ST_HexColor` (§17.18.41) is either the keyword `"auto"` or
 * *exactly* six hex digits, with no `#`. This used to be a bare `value as
 * HexColor` cast — type-safety theatre that promised a validated value
 * without ever checking one — so every caller that handed it a browser
 * `<input type="color">` value (always `#rrggbb`) produced an invalid
 * `w:val="#rrggbb"` that Word rejects wholesale via its "unreadable
 * content" repair path.
 *
 * This is the single choke point every call site funnels through — both the
 * editor/UI side (`toolbarAdapter.ts`'s color picker and hardcoded table
 * border, `pasteBlocks.ts`'s CSS-derived colors) *and* the parser
 * (`parser/document.ts`'s and `parser/styles.ts`'s `parseColor`, reading a
 * `w:val`/`w:color`/`w:fill` straight out of a `.docx` on disk) — so no call
 * site, current or future, can push a `#`-prefixed or otherwise malformed
 * value through to the writer.
 *
 * Behaviour is deliberately split by how "wrong" the input is:
 *   - A leading `#` (what every color `<input>` and CSS color produces) is
 *     stripped; the remaining six digits are otherwise preserved
 *     byte-for-byte, including case, so a color that came from a file and is
 *     already valid round-trips completely unchanged — normalizing case
 *     here would silently rewrite a file's bytes for a value that was never
 *     wrong.
 *   - Anything that still isn't exactly six hex digits after that strip is
 *     genuinely malformed. This function is also the parser's entry point
 *     (it cannot be given a separate one without touching the off-limits
 *     `src/docx/parser/` files), and a parser must never let a corrupt or
 *     hand-edited `w:val` crash the app on open — so rather than throwing,
 *     this falls back to `'000000'` (black), the same as Word's own repair
 *     behavior for an unreadable color. Every current editor/UI call site
 *     only ever passes already-well-formed input (a color `<input>`'s
 *     value, a hardcoded constant, or `parseCssColor`'s output), so this
 *     fallback is not expected to be reachable from the UI in practice.
 */
export const hexColor = (value: string): HexColor => {
  const stripped = value.startsWith('#') ? value.slice(1) : value
  return (HEX_COLOR_DIGITS.test(stripped) ? stripped : '000000') as HexColor
}
export const twip = (value: number): Twip => value as Twip
export const halfPoint = (value: number): HalfPoint => value as HalfPoint
export const eighthPoint = (value: number): EighthPoint => value as EighthPoint
export const pct = (value: number): Pct => value as Pct

// ---------------------------------------------------------------------------
// Shared scalar enums
// ---------------------------------------------------------------------------

export type OnOff = boolean

export type VerticalAlign = 'baseline' | 'superscript' | 'subscript'

export type TextAlignment = 'top' | 'center' | 'baseline' | 'auto' | 'bottom'

export type JustifyContent = 'start' | 'center' | 'end' | 'both' | 'distribute'

export type LineRule = 'auto' | 'exact' | 'atLeast'

export type UnderlineStyle =
  | 'none'
  | 'single'
  | 'words'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashedHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dashDotHeavy'
  | 'dotDotDash'
  | 'dashDotDotHeavy'
  | 'wave'
  | 'wavyHeavy'
  | 'wavyDouble'

export type BorderStyle =
  | 'nil'
  | 'none'
  | 'single'
  | 'thick'
  | 'double'
  | 'dotted'
  | 'dashed'
  | 'dotDash'
  | 'dotDotDash'
  | 'triple'
  | 'thinThickSmallGap'
  | 'thickThinSmallGap'
  | 'thinThickThinSmallGap'
  | 'thinThickMediumGap'
  | 'thickThinMediumGap'
  | 'thinThickThinMediumGap'
  | 'thinThickLargeGap'
  | 'thickThinLargeGap'
  | 'thinThickThinLargeGap'
  | 'wave'
  | 'doubleWave'
  | 'dashSmallGap'
  | 'dashDotStroked'
  | 'threeDEmboss'
  | 'threeDEngrave'
  | 'outset'
  | 'inset'

export type TabAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'decimal'
  | 'bar'
  | 'clear'
  | 'start'
  | 'end'
  | 'num'

export type TabLeader =
  | 'none'
  | 'dot'
  | 'hyphen'
  | 'underscore'
  | 'heavy'
  | 'middleDot'

export type FontHint = 'default' | 'eastAsia' | 'cs' | 'hAnsi'

/** `w:em` — emphasis mark (`CT_Em`/`ST_Em`, §17.18.24), routine in CJK documents. */
export type EmphasisMark = 'none' | 'dot' | 'comma' | 'circle' | 'underDot'

export type HighlightColor =
  | 'black'
  | 'blue'
  | 'cyan'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGray'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'green'
  | 'lightGray'
  | 'magenta'
  | 'none'
  | 'red'
  | 'white'
  | 'yellow'

export type TableLayout = 'fixed' | 'autofit'

export type WidthType = 'auto' | 'dxa' | 'pct' | 'nil'

export type TableCellVerticalAlign = 'top' | 'center' | 'bottom' | 'both'

export type TableRowHeightRule = 'auto' | 'atLeast' | 'exact'

export type FrameDropCap = 'none' | 'drop' | 'margin'

export type FrameWrap = 'none' | 'around' | 'notBeside' | 'through' | 'tight'

export type FrameHorizontalAlign = 'left' | 'center' | 'right' | 'inside' | 'outside'

export type FrameVerticalAlign = 'top' | 'center' | 'bottom' | 'inside' | 'outside'

export type FrameAnchor = 'page' | 'margin' | 'text'

export type NumberingSuffix = 'tab' | 'space' | 'nothing'

// ---------------------------------------------------------------------------
// Shared structured values
// ---------------------------------------------------------------------------

export type Font = string

export type FontTheme =
  | 'majorAscii'
  | 'majorHAnsi'
  | 'majorBidi'
  | 'majorEastAsia'
  | 'minorAscii'
  | 'minorHAnsi'
  | 'minorBidi'
  | 'minorEastAsia'

export interface FontSet {
  readonly ascii?: Font
  readonly hAnsi?: Font
  readonly cs?: Font
  readonly eastAsia?: Font
  readonly hint?: FontHint
  readonly asciiTheme?: FontTheme
  readonly hAnsiTheme?: FontTheme
  readonly csTheme?: FontTheme
  readonly eastAsiaTheme?: FontTheme
}

export interface Underline {
  readonly style: UnderlineStyle
  readonly color?: Color
}

export interface LineSpacing {
  readonly line?: Twip
  readonly lineRule?: LineRule
}

export interface Indent {
  readonly left?: Twip
  readonly right?: Twip
  readonly firstLine?: Twip
  readonly hanging?: Twip
  readonly start?: Twip
  readonly end?: Twip
}

export interface Spacing extends LineSpacing {
  readonly before?: Twip
  readonly after?: Twip
  readonly beforeAutospacing?: OnOff
  readonly afterAutospacing?: OnOff
}

export interface Border {
  readonly style?: BorderStyle
  readonly color?: Color
  readonly size?: EighthPoint
  readonly space?: Twip
  readonly shadow?: OnOff
  readonly frame?: OnOff
}

export interface BorderSet {
  readonly top?: Border
  readonly left?: Border
  readonly bottom?: Border
  readonly right?: Border
  readonly start?: Border
  readonly end?: Border
  readonly between?: Border
  readonly bar?: Border
  readonly insideH?: Border
  readonly insideV?: Border
}

export interface Shading {
  readonly fill?: Color
  readonly color?: Color
  readonly pattern?: string
}

export interface Tab {
  readonly position: Twip
  readonly alignment?: TabAlignment
  readonly leader?: TabLeader
}

export interface TabSet {
  readonly items: ReadonlyArray<Tab>
  readonly defaultTabStop?: Twip
}

export interface LanguageSet {
  readonly value?: string
  readonly eastAsia?: string
  readonly bidi?: string
}

export interface NumPr {
  readonly ilvl?: number
  readonly numId?: string
}

export interface FrameProps {
  readonly width?: Twip
  readonly height?: Twip
  readonly x?: number
  readonly y?: number
  readonly xAlign?: FrameHorizontalAlign
  readonly yAlign?: FrameVerticalAlign
  readonly hAnchor?: FrameAnchor
  readonly vAnchor?: FrameAnchor
  readonly wrap?: FrameWrap
  readonly lines?: number
  readonly hSpace?: Twip
  readonly vSpace?: Twip
  readonly dropCap?: FrameDropCap
  readonly lockAnchor?: OnOff
}

export interface Width {
  readonly type: WidthType
  readonly value?: Twip | Pct
}

export interface InsetSet {
  readonly top?: Width
  readonly left?: Width
  readonly bottom?: Width
  readonly right?: Width
  readonly start?: Width
  readonly end?: Width
}

export interface TableLook {
  readonly value?: string
  readonly firstRow?: OnOff
  readonly lastRow?: OnOff
  readonly firstColumn?: OnOff
  readonly lastColumn?: OnOff
  readonly noHBand?: OnOff
  readonly noVBand?: OnOff
}

export interface TableRowHeight {
  readonly val: Twip
  readonly hRule?: TableRowHeightRule
}

// ---------------------------------------------------------------------------
// Run + paragraph properties
// ---------------------------------------------------------------------------

export interface RunProps {
  readonly rStyle?: string
  readonly bold?: OnOff
  readonly italic?: OnOff
  /** `w:bCs`/`w:iCs` — bold/italic for complex-script text (round-trip fidelity audit, DXS round 2; RTL-adjacent, like `szCs`). */
  readonly boldCs?: OnOff
  readonly italicCs?: OnOff
  readonly underline?: Underline
  readonly strike?: OnOff
  readonly dstrike?: OnOff
  readonly vertAlign?: VerticalAlign
  readonly color?: Color
  readonly highlight?: HighlightColor
  readonly shd?: Shading
  readonly sz?: HalfPoint
  readonly szCs?: HalfPoint
  readonly rFonts?: FontSet
  readonly spacing?: Twip
  readonly kern?: HalfPoint
  readonly position?: HalfPoint
  readonly lang?: LanguageSet
  readonly caps?: OnOff
  readonly smallCaps?: OnOff
  readonly vanish?: OnOff
  readonly webHidden?: OnOff
  readonly rtl?: OnOff
  /** DOCX-12 — `w:outline`: hollow/outlined character effect. */
  readonly outline?: OnOff
  /** DOCX-12 — `w:emboss`: embossed (raised) character effect. */
  readonly emboss?: OnOff
  /** DOCX-12 — `w:imprint`: engraved/imprinted character effect. */
  readonly imprint?: OnOff
  /** DOCX-12 — `w:em`: emphasis mark, routine in CJK documents. */
  readonly em?: EmphasisMark
  /** DOCX-12 — `w:bdr`: run-level (character) border, `CT_Border`. */
  readonly bdr?: Border
  /**
   * DOCX-12 — `w:w`'s `w:val`: manual character width scaling as a whole
   * percentage (`100` = no scaling). Named `charScale` rather than `w` to
   * stay readable — `RunProps.w` would collide visually with `Width`'s own
   * `w:w`-derived field name used elsewhere in this model.
   */
  readonly charScale?: number
  /**
   * DOCX-12 — opaque passthrough for a `w:rPr` child Atlas has no dedicated
   * field for: a real ECMA-376 `CT_RPr`/`EG_RPrBase` sequence member this
   * model doesn't carry (e.g. `w:effect`, `w:eastAsianLayout`, `w:fitText`,
   * `w:noProof`, `w:snapToGrid`, `w:cs`, `w:specVanish`, `w:oMath`) or any
   * other element a source document happens to contain there. Captured so a
   * save doesn't silently drop it — see `RPrUnknownChild`'s doc comment for
   * how the writer uses `before` to reinsert it in schema order.
   */
  readonly rPrUnknown?: ReadonlyArray<RPrUnknownChild>
}

/**
 * One passthrough `w:rPr` child — see `RunProps.rPrUnknown`. `xml` is the
 * element's own captured source text (or a tree-rebuilt equivalent — see
 * `parser/document.ts`'s `parseUnknownNode`, reused for this capture).
 * `before` is the tag name of the nearest *modeled* `w:rPr` sibling that
 * followed this child in the source document; `buildRunPropertiesNode`
 * (`serializer/documentWriter.ts`) looks that sibling up in the element it
 * is rebuilding and splices this child back in immediately ahead of it, so
 * the `CT_RPr` sequence — a strict `xsd:sequence`, not a bag, and one of
 * Word's "unreadable content" triggers when violated — stays correctly
 * ordered. `undefined` means this child was last among the source's `w:rPr`
 * children (or last among the remaining unmodeled ones), so it is appended
 * at the very end instead.
 */
export interface RPrUnknownChild {
  readonly xml: string
  readonly before?: string
}

export interface ParaProps {
  readonly pStyle?: string
  readonly numPr?: NumPr
  readonly spacing?: Spacing
  readonly ind?: Indent
  readonly jc?: JustifyContent
  readonly keepNext?: OnOff
  readonly keepLines?: OnOff
  readonly pageBreakBefore?: OnOff
  readonly widowControl?: OnOff
  readonly suppressLineNumbers?: OnOff
  readonly suppressAutoHyphens?: OnOff
  readonly contextualSpacing?: OnOff
  readonly mirrorIndents?: OnOff
  readonly outlineLvl?: number
  readonly textAlignment?: TextAlignment
  readonly tabs?: TabSet
  readonly pBdr?: BorderSet
  readonly shd?: Shading
  readonly framePr?: FrameProps
  readonly divId?: number
  /** DEFER-2 — `w:bidi`: this paragraph reads right-to-left (alignment and text direction flip). */
  readonly bidi?: OnOff
  readonly sectPr?: SectionProps
}

// ---------------------------------------------------------------------------
// Styles + numbering
// ---------------------------------------------------------------------------

export interface LvlText {
  readonly value: string
  readonly placeholders: ReadonlyArray<number>
}

export interface LvlDef {
  readonly level: number
  readonly start?: number
  readonly restart?: number
  readonly format?: string
  readonly text?: LvlText
  readonly suffix?: NumberingSuffix
  readonly tentative?: OnOff
  readonly legal?: OnOff
  readonly justification?: JustifyContent
  readonly pStyle?: string
  readonly paragraph?: ParaProps
  readonly run?: RunProps
}

export interface LvlOverride {
  readonly level: number
  readonly startOverride?: number
  readonly levelDefinition?: LvlDef
}

export interface NumberingDef {
  readonly numId: string
  readonly abstractNumId?: string
  readonly styleLink?: string
  readonly numberStyleLink?: string
  readonly levels: ReadonlyMap<number, LvlDef>
  readonly levelOverrides?: ReadonlyMap<number, LvlOverride>
  /**
   * DOCX-15 fix — marks a `NumberingDef` that Atlas itself minted for a
   * bullet/numbered-list toggle or paste (`insertList.ts`'s
   * `createListNumberingEntry`), as opposed to one that came from the
   * source document. `abstractNumId` is now a real `ST_DecimalNumber`
   * integer (see that file's doc comment for why), so it can no longer
   * double as this marker the way the old `"atlas-list-"`-prefixed string
   * value used to; `pickListNumId` checks this instead to recognize its own
   * prior list definitions and reuse them, without ever mistaking a real
   * document's own numbering (which might coincidentally match the wanted
   * level-0 format) for one of its own. Always `undefined` for a
   * `NumberingDef` the parser built from a file.
   */
  readonly atlasManaged?: boolean
}

export interface TableStyleProps {
  readonly width?: Width
  readonly indent?: Width
  readonly borders?: BorderSet
  readonly cellMargin?: InsetSet
  readonly layout?: TableLayout
  readonly look?: TableLook
  readonly justification?: JustifyContent
  readonly shading?: Shading
  /**
   * `w:tblStyleRowBandSize`/`w:tblStyleColBandSize` (D7 / DXP-06) — how many
   * consecutive rows/columns each `band1Horz`/`band2Horz`/`band1Vert`/
   * `band2Vert` stripe covers before alternating. Word defaults both to 1
   * (alternate every row/column) when absent.
   */
  readonly rowBandSize?: number
  readonly colBandSize?: number
}

/**
 * The `w:type` values `<w:tblStylePr>` accepts (D7 / DXP-06, DXL-08,
 * DXS-05) — one conditional-formatting block per table-style "region" Word
 * exposes as Table Style Options checkboxes (Header/Total Row, First/Last
 * Column, Banded Rows/Columns) plus the four corner-cell intersections.
 */
export type TableConditionalFormatType =
  | 'wholeTable'
  | 'firstRow'
  | 'lastRow'
  | 'firstCol'
  | 'lastCol'
  | 'band1Vert'
  | 'band2Vert'
  | 'band1Horz'
  | 'band2Horz'
  | 'neCell'
  | 'nwCell'
  | 'seCell'
  | 'swCell'

/**
 * One `<w:tblStylePr>` block's formatting — the OOXML schema (`CT_TblStylePr`)
 * allows the same paragraph/run/table/row/cell property groups as direct
 * formatting, scoped to whichever table region `TableConditionalFormatType`
 * names.
 */
export interface TableConditionalFormat {
  readonly paragraph?: ParaProps
  readonly run?: RunProps
  readonly table?: TableStyleProps
  readonly row?: TableRowProps
  readonly cell?: TableCellProps
}

export interface NumberingStyleProps {
  readonly numId: string
  readonly ilvl?: number
}

export interface DocDefaults {
  readonly paragraph?: ParaProps
  readonly run?: RunProps
}

export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering'

export interface Style {
  readonly id: string
  readonly type: StyleType
  readonly name?: string
  readonly aliases?: ReadonlyArray<string>
  readonly basedOn?: string
  readonly next?: string
  readonly linked?: string
  readonly custom?: OnOff
  readonly isDefault?: OnOff
  readonly uiPriority?: number
  readonly hidden?: OnOff
  readonly semiHidden?: OnOff
  readonly unhideWhenUsed?: OnOff
  readonly qFormat?: OnOff
  readonly locked?: OnOff
  readonly paragraph?: ParaProps
  readonly run?: RunProps
  readonly table?: TableStyleProps
  readonly numbering?: NumberingStyleProps
  /**
   * `<w:tblStylePr>` conditional-formatting blocks, keyed by `w:type`
   * (D7 / DXP-06, DXL-08, DXS-05). Only populated for `type: 'table'`
   * styles. `cascadeTable.ts`'s `resolveTableCellStyle` walks the
   * `basedOn` chain and applies the blocks whose type is "active" for a
   * given cell position, gated by the table instance's own `tblLook`
   * flags.
   */
  readonly conditionalFormats?: ReadonlyMap<TableConditionalFormatType, TableConditionalFormat>
}
