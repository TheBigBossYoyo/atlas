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

export const hexColor = (value: string): HexColor => value as HexColor
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
