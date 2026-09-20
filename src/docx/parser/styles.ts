/**
 * Atlas — DOCX styles.xml parser (Wave A.4)
 *
 * Parses `word/styles.xml` into immutable style/default structures that match
 * the Wave A.2 model.
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import {
  eighthPoint,
  halfPoint,
  hexColor,
  pct,
  twip,
} from '../model'
import type {
  Border,
  BorderSet,
  Color,
  FontHint,
  FontSet,
  FrameAnchor,
  FrameDropCap,
  FrameHorizontalAlign,
  FrameProps,
  FrameVerticalAlign,
  FrameWrap,
  HighlightColor,
  Indent,
  InsetSet,
  JustifyContent,
  LanguageSet,
  LineRule,
  NumberingStyleProps,
  OnOff,
  ParaProps,
  RunProps,
  SectionColumns,
  SectionProps,
  Shading,
  Spacing,
  Style,
  StyleType,
  Tab,
  TabAlignment,
  TabLeader,
  TabSet,
  TableCellMerge,
  TableCellProps,
  TableCellVerticalAlign,
  TableConditionalFormat,
  TableConditionalFormatType,
  TableLayout,
  TableLook,
  TableRowHeight,
  TableRowHeightRule,
  TableRowProps,
  TableStyleProps,
  TextAlignment,
  Underline,
  UnderlineStyle,
  VerticalAlign,
  Width,
  WidthType,
} from '../model'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'

type XmlScalar = string | number | boolean
type XmlValue = XmlScalar | XmlNode | XmlValue[]

interface XmlNode {
  readonly [key: string]: XmlValue | undefined
}

interface RawStylesDocument {
  readonly 'w:styles'?: XmlValue
}

type Mutable<T> = {
  -readonly [K in keyof T]?: T[K]
}

export interface StylesPart {
  readonly docDefaults: {
    readonly rPr?: RunProps
    readonly pPr?: ParaProps
  }
  readonly styles: ReadonlyMap<string, Style>
  /**
   * Opaque passthrough for `<w:latentStyles>` (D19 / DXS-14) — Word's
   * per-style-name UI defaults (a `w:lsdException` list) that Atlas's
   * model has no first-class representation for. Captured as the raw
   * parsed object shape (this module's plain, non-`preserveOrder` XML
   * shape) and re-emitted as-is by `stylesWriter.ts` instead of being
   * silently dropped on every save.
   */
  readonly latentStyles?: unknown
}

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

export function parseStyles(xml: string): StylesPart {
  assertXmlPartSizeWithinLimit(xml, 'word/styles.xml')
  let raw: RawStylesDocument
  try {
    raw = xmlParser.parse(xml) as RawStylesDocument
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse styles XML: ${msg}`)
  }
  const stylesRoot = asXmlNode(raw['w:styles'])

  const docDefaultsNode = getNode(stylesRoot, 'w:docDefaults')
  const rPrDefaultNode = getNode(docDefaultsNode, 'w:rPrDefault')
  const pPrDefaultNode = getNode(docDefaultsNode, 'w:pPrDefault')

  const rPr = parseRunPropsElement(getNode(rPrDefaultNode, 'w:rPr'))
  const pPr = parseParaPropsElement(getNode(pPrDefaultNode, 'w:pPr'))

  const styles = new Map<string, Style>()

  for (const styleNode of getNodes(stylesRoot, 'w:style')) {
    const style = parseStyleNode(styleNode)
    if (style === undefined) continue
    styles.set(style.id, style)
  }

  const latentStyles = getNode(stylesRoot, 'w:latentStyles')

  return {
    docDefaults: {
      ...(rPr !== undefined ? { rPr } : {}),
      ...(pPr !== undefined ? { pPr } : {}),
    },
    styles,
    ...(latentStyles !== undefined ? { latentStyles } : {}),
  }
}

export function parseRunPropsNode(node: unknown): RunProps | undefined {
  return parseRunPropsElement(asXmlNode(node))
}

export function parseParaPropsNode(node: unknown): ParaProps | undefined {
  return parseParaPropsElement(asXmlNode(node))
}

function parseStyleNode(node: XmlNode): Style | undefined {
  const id = getAttr(node, 'w:styleId')?.trim()
  const type = parseStyleType(getAttr(node, 'w:type'))

  if (id === undefined || id === '' || type === undefined) {
    return undefined
  }

  const paragraph = parseParaPropsElement(getNode(node, 'w:pPr'))
  const run = parseRunPropsElement(getNode(node, 'w:rPr'))
  const table = parseTableStylePropsElement(getNode(node, 'w:tblPr'))
  const numbering = parseNumberingStyleProps(type, paragraph)
  const aliases = parseAliases(getValAttr(getNode(node, 'w:aliases')))
  const conditionalFormats = parseTableConditionalFormats(node)

  return {
    id,
    type,
    ...(withValue('name', getValAttr(getNode(node, 'w:name')))),
    ...(aliases !== undefined ? { aliases } : {}),
    ...(withValue('basedOn', getValAttr(getNode(node, 'w:basedOn')))),
    ...(withValue('next', getValAttr(getNode(node, 'w:next')))),
    ...(withValue('linked', getValAttr(getNode(node, 'w:link')))),
    ...(withValue('custom', parseOnOffAttr(getAttr(node, 'w:customStyle')))),
    ...(withValue('isDefault', parseOnOffAttr(getAttr(node, 'w:default')))),
    ...(withValue('uiPriority', parseInteger(getValAttr(getNode(node, 'w:uiPriority'))))),
    ...(withValue('hidden', parseOnOffElement(node['w:hidden']))),
    ...(withValue('semiHidden', parseOnOffElement(node['w:semiHidden']))),
    ...(withValue('unhideWhenUsed', parseOnOffElement(node['w:unhideWhenUsed']))),
    ...(withValue('qFormat', parseOnOffElement(node['w:qFormat']))),
    ...(withValue('locked', parseOnOffElement(node['w:locked']))),
    ...(paragraph !== undefined ? { paragraph } : {}),
    ...(run !== undefined ? { run } : {}),
    ...(table !== undefined ? { table } : {}),
    ...(numbering !== undefined ? { numbering } : {}),
    ...(conditionalFormats !== undefined ? { conditionalFormats } : {}),
  }
}

/**
 * D7 / DXP-06, DXL-08, DXS-05: parses every `<w:tblStylePr>` sibling of
 * `<w:tblPr>`/`<w:pPr>`/`<w:rPr>` within a `<w:style>` element — one block
 * per table "region" (header row, banded rows/columns, first/last column,
 * corner-cell intersections) a table style conditionally formats. Returns
 * `undefined` when the style has none (the overwhelmingly common case for
 * non-table styles, and even for many table styles).
 */
function parseTableConditionalFormats(
  node: XmlNode,
): ReadonlyMap<TableConditionalFormatType, TableConditionalFormat> | undefined {
  const formats = new Map<TableConditionalFormatType, TableConditionalFormat>()

  for (const tblStylePrNode of getNodes(node, 'w:tblStylePr')) {
    const type = parseTableConditionalFormatType(getAttr(tblStylePrNode, 'w:type'))
    if (type === undefined) continue

    const format = parseTableConditionalFormat(tblStylePrNode)
    if (format !== undefined) {
      formats.set(type, format)
    }
  }

  return formats.size > 0 ? formats : undefined
}

function parseTableConditionalFormat(node: XmlNode): TableConditionalFormat | undefined {
  const paragraph = parseParaPropsElement(getNode(node, 'w:pPr'))
  const run = parseRunPropsElement(getNode(node, 'w:rPr'))
  const table = parseTableStylePropsElement(getNode(node, 'w:tblPr'))
  const row = parseTableStylePrRowProps(getNode(node, 'w:trPr'))
  const cell = parseTableStylePrCellProps(getNode(node, 'w:tcPr'))

  const format: TableConditionalFormat = {
    ...(paragraph !== undefined ? { paragraph } : {}),
    ...(run !== undefined ? { run } : {}),
    ...(table !== undefined ? { table } : {}),
    ...(row !== undefined ? { row } : {}),
    ...(cell !== undefined ? { cell } : {}),
  }

  return hasKeys(format) ? format : undefined
}

const TABLE_CONDITIONAL_FORMAT_TYPES: ReadonlySet<TableConditionalFormatType> = new Set([
  'wholeTable',
  'firstRow',
  'lastRow',
  'firstCol',
  'lastCol',
  'band1Vert',
  'band2Vert',
  'band1Horz',
  'band2Horz',
  'neCell',
  'nwCell',
  'seCell',
  'swCell',
])

function parseTableConditionalFormatType(value: string | undefined): TableConditionalFormatType | undefined {
  if (value === undefined) return undefined
  return (TABLE_CONDITIONAL_FORMAT_TYPES as ReadonlySet<string>).has(value)
    ? (value as TableConditionalFormatType)
    : undefined
}

/**
 * The `w:trPr` subset OOXML's `CT_TrPrBase` allows inside `w:tblStylePr`
 * (a small subset of the full row-properties schema — no `w:divId` or
 * `w:tblCellSpacing`, which Atlas's `TableRowProps` model has no field for
 * anyway).
 */
function parseTableStylePrRowProps(node: XmlNode | undefined): TableRowProps | undefined {
  if (node === undefined) return undefined

  const props: Mutable<TableRowProps> = {}

  const trHeight = parseTableStylePrRowHeight(getNode(node, 'w:trHeight'))
  if (trHeight !== undefined) props.trHeight = trHeight

  const cantSplit = parseOnOffElement(node['w:cantSplit'])
  if (cantSplit !== undefined) props.cantSplit = cantSplit

  const tblHeader = parseOnOffElement(node['w:tblHeader'])
  if (tblHeader !== undefined) props.tblHeader = tblHeader

  const jc = parseJustifyContent(getValAttr(getNode(node, 'w:jc')))
  if (jc !== undefined) props.jc = jc

  return hasKeys(props) ? props : undefined
}

function parseTableStylePrRowHeight(node: XmlNode | undefined): TableRowHeight | undefined {
  if (node === undefined) return undefined

  const val = parseTwipAttr(node, 'w:val')
  if (val === undefined) return undefined

  const hRule = parseTableRowHeightRule(getAttr(node, 'w:hRule'))
  return { val, ...(hRule !== undefined ? { hRule } : {}) }
}

function parseTableRowHeightRule(value: string | undefined): TableRowHeightRule | undefined {
  switch (value) {
    case 'auto':
    case 'atLeast':
    case 'exact':
      return value
    default:
      return undefined
  }
}

/**
 * The `w:tcPr` subset OOXML's `CT_TcPrBase` allows inside `w:tblStylePr` —
 * covers everything `layout/layoutTable.ts` consumes from a cell's direct
 * `w:tcPr` today (shading, borders, margins, vertical alignment).
 */
function parseTableStylePrCellProps(node: XmlNode | undefined): TableCellProps | undefined {
  if (node === undefined) return undefined

  const props: Mutable<TableCellProps> = {}

  const tcW = parseWidth(getNode(node, 'w:tcW'))
  if (tcW !== undefined) props.tcW = tcW

  const gridSpan = parseInteger(getValAttr(getNode(node, 'w:gridSpan')))
  if (gridSpan !== undefined) props.gridSpan = gridSpan

  const vMerge = parseTableStylePrCellMerge(getNode(node, 'w:vMerge'))
  if (vMerge !== undefined) props.vMerge = vMerge

  const tcBorders = parseBorderSet(getNode(node, 'w:tcBorders'))
  if (tcBorders !== undefined) props.tcBorders = tcBorders

  const shd = parseShading(getNode(node, 'w:shd'))
  if (shd !== undefined) props.shd = shd

  const tcMar = parseInsetSet(getNode(node, 'w:tcMar'))
  if (tcMar !== undefined) props.tcMar = tcMar

  const vAlign = parseTableStylePrCellVerticalAlign(getValAttr(getNode(node, 'w:vAlign')))
  if (vAlign !== undefined) props.vAlign = vAlign

  const noWrap = parseOnOffElement(node['w:noWrap'])
  if (noWrap !== undefined) props.noWrap = noWrap

  const hideMark = parseOnOffElement(node['w:hideMark'])
  if (hideMark !== undefined) props.hideMark = hideMark

  return hasKeys(props) ? props : undefined
}

function parseTableStylePrCellMerge(node: XmlNode | undefined): TableCellMerge | undefined {
  if (node === undefined) return undefined
  // `w:val`'s schema default is "continue" — mirrors `parser/document.ts`'s
  // `parseTableCellMerge` for the same element on a direct `w:tcPr`.
  return getValAttr(node) === 'restart' ? 'restart' : 'continue'
}

function parseTableStylePrCellVerticalAlign(value: string | undefined): TableCellVerticalAlign | undefined {
  switch (value) {
    case 'top':
    case 'center':
    case 'bottom':
    case 'both':
      return value
    default:
      return undefined
  }
}

function parseRunPropsElement(node: XmlNode | undefined): RunProps | undefined {
  if (node === undefined) return undefined

  const props: Mutable<RunProps> = {}

  const rStyle = getValAttr(getNode(node, 'w:rStyle'))
  if (rStyle !== undefined) props.rStyle = rStyle

  const bold = parseOnOffElement(node['w:b'])
  if (bold !== undefined) props.bold = bold

  const italic = parseOnOffElement(node['w:i'])
  if (italic !== undefined) props.italic = italic

  // Round-trip fidelity audit (DXS round 2): complex-script bold/italic,
  // mirroring `document.ts`'s direct-formatting `w:bCs`/`w:iCs` handling.
  const boldCs = parseOnOffElement(node['w:bCs'])
  if (boldCs !== undefined) props.boldCs = boldCs

  const italicCs = parseOnOffElement(node['w:iCs'])
  if (italicCs !== undefined) props.italicCs = italicCs

  const underline = parseUnderline(getNode(node, 'w:u'))
  if (underline !== undefined) props.underline = underline

  const strike = parseOnOffElement(node['w:strike'])
  if (strike !== undefined) props.strike = strike

  const dstrike = parseOnOffElement(node['w:dstrike'])
  if (dstrike !== undefined) props.dstrike = dstrike

  const vertAlign = parseVerticalAlign(getValAttr(getNode(node, 'w:vertAlign')))
  if (vertAlign !== undefined) props.vertAlign = vertAlign

  const color = parseColor(getValAttr(getNode(node, 'w:color')))
  if (color !== undefined) props.color = color

  const highlight = parseHighlightColor(getValAttr(getNode(node, 'w:highlight')))
  if (highlight !== undefined) props.highlight = highlight

  const shd = parseShading(getNode(node, 'w:shd'))
  if (shd !== undefined) props.shd = shd

  const sz = parseHalfPointAttr(getNode(node, 'w:sz'), 'w:val')
  if (sz !== undefined) props.sz = sz

  const szCs = parseHalfPointAttr(getNode(node, 'w:szCs'), 'w:val')
  if (szCs !== undefined) props.szCs = szCs

  const rFonts = parseFontSet(getNode(node, 'w:rFonts'))
  if (rFonts !== undefined) props.rFonts = rFonts

  const spacing = parseTwipAttr(getNode(node, 'w:spacing'), 'w:val')
  if (spacing !== undefined) props.spacing = spacing

  const kern = parseHalfPointAttr(getNode(node, 'w:kern'), 'w:val')
  if (kern !== undefined) props.kern = kern

  const position = parseHalfPointAttr(getNode(node, 'w:position'), 'w:val')
  if (position !== undefined) props.position = position

  const lang = parseLanguageSet(getNode(node, 'w:lang'))
  if (lang !== undefined) props.lang = lang

  const caps = parseOnOffElement(node['w:caps'])
  if (caps !== undefined) props.caps = caps

  const smallCaps = parseOnOffElement(node['w:smallCaps'])
  if (smallCaps !== undefined) props.smallCaps = smallCaps

  const vanish = parseOnOffElement(node['w:vanish'])
  if (vanish !== undefined) props.vanish = vanish

  const webHidden = parseOnOffElement(node['w:webHidden'])
  if (webHidden !== undefined) props.webHidden = webHidden

  const rtl = parseOnOffElement(node['w:rtl'])
  if (rtl !== undefined) props.rtl = rtl

  // DOCX-12 — mirrors `document.ts`'s `parseRunProps` (see its doc comment):
  // a style's own run properties (`w:docDefaults`/named-style/table-
  // conditional-format `w:rPr`) previously lost these the same way direct
  // run formatting did.
  const outline = parseOnOffElement(node['w:outline'])
  if (outline !== undefined) props.outline = outline

  const emboss = parseOnOffElement(node['w:emboss'])
  if (emboss !== undefined) props.emboss = emboss

  const imprint = parseOnOffElement(node['w:imprint'])
  if (imprint !== undefined) props.imprint = imprint

  const em = parseEmphasisMark(getValAttr(getNode(node, 'w:em')))
  if (em !== undefined) props.em = em

  const bdr = parseBorder(getNode(node, 'w:bdr'))
  if (bdr !== undefined) props.bdr = bdr

  const charScale = parseCharScale(getValAttr(getNode(node, 'w:w')))
  if (charScale !== undefined) props.charScale = charScale

  return hasKeys(props) ? props : undefined
}

function parseEmphasisMark(value: string | undefined): RunProps['em'] {
  switch (value) {
    case 'none':
    case 'dot':
    case 'comma':
    case 'circle':
    case 'underDot':
      return value
    default:
      return undefined
  }
}

/** `w:w`'s `w:val` (`ST_TextScale`) — see `document.ts`'s `parseCharScale` doc comment. */
function parseCharScale(value: string | undefined): number | undefined {
  if (value === undefined) return undefined

  const normalized = value.endsWith('%') ? value.slice(0, -1) : value
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : undefined
}

function parseParaPropsElement(node: XmlNode | undefined): ParaProps | undefined {
  if (node === undefined) return undefined

  const props: Mutable<ParaProps> = {}

  const pStyle = getValAttr(getNode(node, 'w:pStyle'))
  if (pStyle !== undefined) props.pStyle = pStyle

  const numPr = parseNumPr(getNode(node, 'w:numPr'))
  if (numPr !== undefined) props.numPr = numPr

  const spacing = parseSpacing(getNode(node, 'w:spacing'))
  if (spacing !== undefined) props.spacing = spacing

  const ind = parseIndent(getNode(node, 'w:ind'))
  if (ind !== undefined) props.ind = ind

  const jc = parseJustifyContent(getValAttr(getNode(node, 'w:jc')))
  if (jc !== undefined) props.jc = jc

  const keepNext = parseOnOffElement(node['w:keepNext'])
  if (keepNext !== undefined) props.keepNext = keepNext

  const keepLines = parseOnOffElement(node['w:keepLines'])
  if (keepLines !== undefined) props.keepLines = keepLines

  const pageBreakBefore = parseOnOffElement(node['w:pageBreakBefore'])
  if (pageBreakBefore !== undefined) props.pageBreakBefore = pageBreakBefore

  const widowControl = parseOnOffElement(node['w:widowControl'])
  if (widowControl !== undefined) props.widowControl = widowControl

  const suppressLineNumbers = parseOnOffElement(node['w:suppressLineNumbers'])
  if (suppressLineNumbers !== undefined) props.suppressLineNumbers = suppressLineNumbers

  const suppressAutoHyphens = parseOnOffElement(node['w:suppressAutoHyphens'])
  if (suppressAutoHyphens !== undefined) props.suppressAutoHyphens = suppressAutoHyphens

  const contextualSpacing = parseOnOffElement(node['w:contextualSpacing'])
  if (contextualSpacing !== undefined) props.contextualSpacing = contextualSpacing

  const mirrorIndents = parseOnOffElement(node['w:mirrorIndents'])
  if (mirrorIndents !== undefined) props.mirrorIndents = mirrorIndents

  const outlineLvl = parseInteger(getValAttr(getNode(node, 'w:outlineLvl')))
  if (outlineLvl !== undefined) props.outlineLvl = outlineLvl

  const textAlignment = parseTextAlignment(getValAttr(getNode(node, 'w:textAlignment')))
  if (textAlignment !== undefined) props.textAlignment = textAlignment

  const tabs = parseTabSet(getNode(node, 'w:tabs'))
  if (tabs !== undefined) props.tabs = tabs

  const pBdr = parseBorderSet(getNode(node, 'w:pBdr'))
  if (pBdr !== undefined) props.pBdr = pBdr

  const shd = parseShading(getNode(node, 'w:shd'))
  if (shd !== undefined) props.shd = shd

  const framePr = parseFrameProps(getNode(node, 'w:framePr'))
  if (framePr !== undefined) props.framePr = framePr

  const divId = parseInteger(getValAttr(getNode(node, 'w:divId')))
  if (divId !== undefined) props.divId = divId

  const sectPr = parseSectionProps(getNode(node, 'w:sectPr'))
  if (sectPr !== undefined) props.sectPr = sectPr

  return hasKeys(props) ? props : undefined
}

function parseTableStylePropsElement(node: XmlNode | undefined): TableStyleProps | undefined {
  if (node === undefined) return undefined

  const props: TableStyleProps = {
    ...(withValue('width', parseWidth(getNode(node, 'w:tblW')))),
    ...(withValue('indent', parseWidth(getNode(node, 'w:tblInd')))),
    ...(withValue('borders', parseBorderSet(getNode(node, 'w:tblBorders')))),
    ...(withValue('cellMargin', parseInsetSet(getNode(node, 'w:tblCellMar')))),
    ...(withValue('layout', parseTableLayout(getValAttr(getNode(node, 'w:tblLayout'))))),
    ...(withValue('look', parseTableLook(getNode(node, 'w:tblLook')))),
    ...(withValue('justification', parseJustifyContent(getValAttr(getNode(node, 'w:jc'))))),
    ...(withValue('shading', parseShading(getNode(node, 'w:shd')))),
    ...(withValue('rowBandSize', parseInteger(getValAttr(getNode(node, 'w:tblStyleRowBandSize'))))),
    ...(withValue('colBandSize', parseInteger(getValAttr(getNode(node, 'w:tblStyleColBandSize'))))),
  }

  return hasKeys(props) ? props : undefined
}

function parseNumberingStyleProps(
  type: StyleType,
  paragraph: ParaProps | undefined,
): NumberingStyleProps | undefined {
  if (type !== 'numbering') return undefined
  if (paragraph?.numPr?.numId === undefined) return undefined

  return {
    numId: paragraph.numPr.numId,
    ...(withValue('ilvl', paragraph.numPr.ilvl)),
  }
}

function parseUnderline(node: XmlNode | undefined): Underline | undefined {
  if (node === undefined) return undefined

  const style = parseUnderlineStyle(getAttr(node, 'w:val')) ?? 'single'
  const color = parseColor(getAttr(node, 'w:color'))

  return {
    style,
    ...(withValue('color', color)),
  }
}

function parseVerticalAlign(value: string | undefined): VerticalAlign | undefined {
  switch (value) {
    case 'baseline':
    case 'superscript':
    case 'subscript':
      return value
    default:
      return undefined
  }
}

function parseHighlightColor(value: string | undefined): HighlightColor | undefined {
  switch (value) {
    case 'black':
    case 'blue':
    case 'cyan':
    case 'darkBlue':
    case 'darkCyan':
    case 'darkGray':
    case 'darkGreen':
    case 'darkMagenta':
    case 'darkRed':
    case 'darkYellow':
    case 'green':
    case 'lightGray':
    case 'magenta':
    case 'none':
    case 'red':
    case 'white':
    case 'yellow':
      return value
    default:
      return undefined
  }
}

function parseFontSet(node: XmlNode | undefined): FontSet | undefined {
  if (node === undefined) return undefined

  const fontSet: FontSet = {
    ...(withValue('ascii', getAttr(node, 'w:ascii'))),
    ...(withValue('hAnsi', getAttr(node, 'w:hAnsi'))),
    ...(withValue('cs', getAttr(node, 'w:cs'))),
    ...(withValue('eastAsia', getAttr(node, 'w:eastAsia'))),
    ...(withValue('hint', parseFontHint(getAttr(node, 'w:hint')))),
    ...(withValue('asciiTheme', parseFontTheme(getAttr(node, 'w:asciiTheme')))),
    ...(withValue('hAnsiTheme', parseFontTheme(getAttr(node, 'w:hAnsiTheme')))),
    ...(withValue('csTheme', parseFontTheme(getAttr(node, 'w:cstheme')))),
    ...(withValue('eastAsiaTheme', parseFontTheme(getAttr(node, 'w:eastAsiaTheme')))),
  }

  return hasKeys(fontSet) ? fontSet : undefined
}

function parseFontTheme(value: string | undefined): FontSet['asciiTheme'] | undefined {
  switch (value) {
    case 'majorAscii':
    case 'majorHAnsi':
    case 'majorBidi':
    case 'majorEastAsia':
    case 'minorAscii':
    case 'minorHAnsi':
    case 'minorBidi':
    case 'minorEastAsia':
      return value
    default:
      return undefined
  }
}

function parseLanguageSet(node: XmlNode | undefined): LanguageSet | undefined {
  if (node === undefined) return undefined

  const languageSet: LanguageSet = {
    ...(withValue('value', getAttr(node, 'w:val'))),
    ...(withValue('eastAsia', getAttr(node, 'w:eastAsia'))),
    ...(withValue('bidi', getAttr(node, 'w:bidi'))),
  }

  return hasKeys(languageSet) ? languageSet : undefined
}

function parseNumPr(node: XmlNode | undefined): ParaProps['numPr'] | undefined {
  if (node === undefined) return undefined

  const ilvl = parseInteger(getValAttr(getNode(node, 'w:ilvl')))
  const numId = getValAttr(getNode(node, 'w:numId'))

  if (ilvl === undefined && numId === undefined) return undefined

  return {
    ...(withValue('ilvl', ilvl)),
    ...(withValue('numId', numId)),
  }
}

function parseSpacing(node: XmlNode | undefined): Spacing | undefined {
  if (node === undefined) return undefined

  const spacing: Spacing = {
    ...(withValue('before', parseTwipAttr(node, 'w:before'))),
    ...(withValue('after', parseTwipAttr(node, 'w:after'))),
    ...(withValue('beforeAutospacing', parseOnOffAttr(getAttr(node, 'w:beforeAutospacing')))),
    ...(withValue('afterAutospacing', parseOnOffAttr(getAttr(node, 'w:afterAutospacing')))),
    ...(withValue('line', parseTwipAttr(node, 'w:line'))),
    ...(withValue('lineRule', parseLineRule(getAttr(node, 'w:lineRule')))),
  }

  return hasKeys(spacing) ? spacing : undefined
}

function parseIndent(node: XmlNode | undefined): Indent | undefined {
  if (node === undefined) return undefined

  const indent: Indent = {
    ...(withValue('left', parseTwipAttr(node, 'w:left'))),
    ...(withValue('right', parseTwipAttr(node, 'w:right'))),
    ...(withValue('firstLine', parseTwipAttr(node, 'w:firstLine'))),
    ...(withValue('hanging', parseTwipAttr(node, 'w:hanging'))),
    ...(withValue('start', parseTwipAttr(node, 'w:start'))),
    ...(withValue('end', parseTwipAttr(node, 'w:end'))),
  }

  return hasKeys(indent) ? indent : undefined
}

function parseTabSet(node: XmlNode | undefined): TabSet | undefined {
  if (node === undefined) return undefined

  const items: ReadonlyArray<Tab> = getNodes(node, 'w:tab')
    .map(parseTab)
    .filter((tab): tab is Tab => tab !== undefined)

  const defaultTabStop = parseTwipAttr(node, 'w:defaultTabStop')

  if (items.length === 0 && defaultTabStop === undefined) {
    return undefined
  }

  return {
    items,
    ...(withValue('defaultTabStop', defaultTabStop)),
  }
}

function parseTab(node: XmlNode): Tab | undefined {
  const position = parseTwipAttr(node, 'w:pos')
  if (position === undefined) return undefined

  return {
    position,
    ...(withValue('alignment', parseTabAlignment(getAttr(node, 'w:val')))),
    ...(withValue('leader', parseTabLeader(getAttr(node, 'w:leader')))),
  }
}

function parseBorderSet(node: XmlNode | undefined): BorderSet | undefined {
  if (node === undefined) return undefined

  const borderSet: BorderSet = {
    ...(withValue('top', parseBorder(getNode(node, 'w:top')))),
    ...(withValue('left', parseBorder(getNode(node, 'w:left')))),
    ...(withValue('bottom', parseBorder(getNode(node, 'w:bottom')))),
    ...(withValue('right', parseBorder(getNode(node, 'w:right')))),
    ...(withValue('start', parseBorder(getNode(node, 'w:start')))),
    ...(withValue('end', parseBorder(getNode(node, 'w:end')))),
    ...(withValue('between', parseBorder(getNode(node, 'w:between')))),
    ...(withValue('bar', parseBorder(getNode(node, 'w:bar')))),
    ...(withValue('insideH', parseBorder(getNode(node, 'w:insideH')))),
    ...(withValue('insideV', parseBorder(getNode(node, 'w:insideV')))),
  }

  return hasKeys(borderSet) ? borderSet : undefined
}

function parseBorder(node: XmlNode | undefined): Border | undefined {
  if (node === undefined) return undefined

  const border: Border = {
    ...(withValue('style', parseBorderStyle(getAttr(node, 'w:val')))),
    ...(withValue('color', parseColor(getAttr(node, 'w:color')))),
    ...(withValue('size', parseEighthPointAttr(node, 'w:sz'))),
    ...(withValue('space', parseTwipAttr(node, 'w:space'))),
    ...(withValue('shadow', parseOnOffAttr(getAttr(node, 'w:shadow')))),
    ...(withValue('frame', parseOnOffAttr(getAttr(node, 'w:frame')))),
  }

  return hasKeys(border) ? border : undefined
}

function parseShading(node: XmlNode | undefined): Shading | undefined {
  if (node === undefined) return undefined

  const shading: Shading = {
    ...(withValue('fill', parseColor(getAttr(node, 'w:fill')))),
    ...(withValue('color', parseColor(getAttr(node, 'w:color')))),
    ...(withValue('pattern', getAttr(node, 'w:val'))),
  }

  return hasKeys(shading) ? shading : undefined
}

function parseFrameProps(node: XmlNode | undefined): FrameProps | undefined {
  if (node === undefined) return undefined

  const frameProps: FrameProps = {
    ...(withValue('width', parseTwipAttr(node, 'w:w'))),
    ...(withValue('height', parseTwipAttr(node, 'w:h'))),
    ...(withValue('x', parseInteger(getAttr(node, 'w:x')))),
    ...(withValue('y', parseInteger(getAttr(node, 'w:y')))),
    ...(withValue('xAlign', parseFrameHorizontalAlign(getAttr(node, 'w:xAlign')))),
    ...(withValue('yAlign', parseFrameVerticalAlign(getAttr(node, 'w:yAlign')))),
    ...(withValue('hAnchor', parseFrameAnchor(getAttr(node, 'w:hAnchor')))),
    ...(withValue('vAnchor', parseFrameAnchor(getAttr(node, 'w:vAnchor')))),
    ...(withValue('wrap', parseFrameWrap(getAttr(node, 'w:wrap')))),
    ...(withValue('lines', parseInteger(getAttr(node, 'w:lines')))),
    ...(withValue('hSpace', parseTwipAttr(node, 'w:hSpace'))),
    ...(withValue('vSpace', parseTwipAttr(node, 'w:vSpace'))),
    ...(withValue('dropCap', parseFrameDropCap(getAttr(node, 'w:dropCap')))),
    ...(withValue('lockAnchor', parseOnOffAttr(getAttr(node, 'w:lockAnchor')))),
  }

  return hasKeys(frameProps) ? frameProps : undefined
}

function parseSectionProps(node: XmlNode | undefined): SectionProps | undefined {
  if (node === undefined) return undefined

  const props: SectionProps = {
    ...(withValue('pgSz', parsePageSize(getNode(node, 'w:pgSz')))),
    ...(withValue('pgMar', parsePageMargins(getNode(node, 'w:pgMar')))),
    ...(withValue('cols', parseSectionColumns(getNode(node, 'w:cols')))),
    ...(withValue('pgNumType', parsePageNumberType(getNode(node, 'w:pgNumType')))),
    ...(withValue('titlePg', parseOnOffElement(node['w:titlePg']))),
    ...(withValue('type', parseSectionBreakType(getValAttr(getNode(node, 'w:type'))))),
    ...(withValue('headerReference', parseHeaderReferences(node))),
    ...(withValue('footerReference', parseFooterReferences(node))),
    ...(withValue('lnNumType', parseLineNumberType(getNode(node, 'w:lnNumType')))),
    ...(withValue('vAlign', parseSectionVerticalAlign(getValAttr(getNode(node, 'w:vAlign'))))),
  }

  return hasKeys(props) ? props : undefined
}

function parsePageSize(node: XmlNode | undefined): SectionProps['pgSz'] | undefined {
  if (node === undefined) return undefined

  const w = parseTwipAttr(node, 'w:w')
  const h = parseTwipAttr(node, 'w:h')

  if (w === undefined || h === undefined) return undefined

  return {
    w,
    h,
    ...(withValue('orient', parsePageOrientation(getAttr(node, 'w:orient')))),
  }
}

function parsePageMargins(node: XmlNode | undefined): SectionProps['pgMar'] | undefined {
  if (node === undefined) return undefined

  const pageMargins: SectionProps['pgMar'] = {
    ...(withValue('top', parseTwipAttr(node, 'w:top'))),
    ...(withValue('right', parseTwipAttr(node, 'w:right'))),
    ...(withValue('bottom', parseTwipAttr(node, 'w:bottom'))),
    ...(withValue('left', parseTwipAttr(node, 'w:left'))),
    ...(withValue('header', parseTwipAttr(node, 'w:header'))),
    ...(withValue('footer', parseTwipAttr(node, 'w:footer'))),
    ...(withValue('gutter', parseTwipAttr(node, 'w:gutter'))),
  }

  return hasKeys(pageMargins) ? pageMargins : undefined
}

function parseSectionColumns(node: XmlNode | undefined): SectionColumns | undefined {
  if (node === undefined) return undefined

  const col = getNodes(node, 'w:col').map((colNode) => ({
    ...(withValue('w', parseTwipAttr(colNode, 'w:w'))),
    ...(withValue('space', parseTwipAttr(colNode, 'w:space'))),
  }))

  const sectionColumns: SectionColumns = {
    ...(withValue('num', parseInteger(getAttr(node, 'w:num')))),
    ...(withValue('space', parseTwipAttr(node, 'w:space'))),
    ...(withValue('sep', parseOnOffAttr(getAttr(node, 'w:sep')))),
    ...(withValue('equalWidth', parseOnOffAttr(getAttr(node, 'w:equalWidth')))),
    col,
  }

  return hasKeys(sectionColumns) ? sectionColumns : undefined
}

function parsePageNumberType(node: XmlNode | undefined): SectionProps['pgNumType'] | undefined {
  if (node === undefined) return undefined

  const pageNumberType: SectionProps['pgNumType'] = {
    ...(withValue('start', parseInteger(getAttr(node, 'w:start')))),
    ...(withValue('fmt', getAttr(node, 'w:fmt'))),
  }

  return hasKeys(pageNumberType) ? pageNumberType : undefined
}

function parseHeaderReferences(node: XmlNode): SectionProps['headerReference'] | undefined {
  const headerReference = getNodes(node, 'w:headerReference')
    .map((item) => {
      const id = getAttr(item, 'r:id')
      const type = parseHeaderFooterReferenceType(getAttr(item, 'w:type'))
      if (id === undefined || type === undefined) return undefined
      return { id, type }
    })
    .filter(
      (
        item,
      ): item is NonNullable<SectionProps['headerReference']>[number] => item !== undefined,
    )

  return headerReference.length > 0 ? headerReference : undefined
}

function parseFooterReferences(node: XmlNode): SectionProps['footerReference'] | undefined {
  const footerReference = getNodes(node, 'w:footerReference')
    .map((item) => {
      const id = getAttr(item, 'r:id')
      const type = parseHeaderFooterReferenceType(getAttr(item, 'w:type'))
      if (id === undefined || type === undefined) return undefined
      return { id, type }
    })
    .filter(
      (
        item,
      ): item is NonNullable<SectionProps['footerReference']>[number] => item !== undefined,
    )

  return footerReference.length > 0 ? footerReference : undefined
}

function parseLineNumberType(node: XmlNode | undefined): SectionProps['lnNumType'] | undefined {
  if (node === undefined) return undefined

  const lineNumberType: SectionProps['lnNumType'] = {
    ...(withValue('countBy', parseInteger(getAttr(node, 'w:countBy')))),
    ...(withValue('start', parseInteger(getAttr(node, 'w:start')))),
    ...(withValue('distance', parseTwipAttr(node, 'w:distance'))),
    ...(withValue('restart', parseLineNumberRestart(getAttr(node, 'w:restart')))),
  }

  return hasKeys(lineNumberType) ? lineNumberType : undefined
}

function parseWidth(node: XmlNode | undefined): Width | undefined {
  if (node === undefined) return undefined

  const type = parseWidthType(getAttr(node, 'w:type'))
  if (type === undefined) return undefined

  const rawValue = parseInteger(getAttr(node, 'w:w'))
  const value =
    type === 'pct'
      ? rawValue !== undefined
        ? pct(rawValue)
        : undefined
      : type === 'dxa'
        ? rawValue !== undefined
          ? twip(rawValue)
          : undefined
        : undefined

  return {
    type,
    ...(withValue('value', value)),
  }
}

function parseInsetSet(node: XmlNode | undefined): InsetSet | undefined {
  if (node === undefined) return undefined

  const insetSet: InsetSet = {
    ...(withValue('top', parseWidth(getNode(node, 'w:top')))),
    ...(withValue('left', parseWidth(getNode(node, 'w:left')))),
    ...(withValue('bottom', parseWidth(getNode(node, 'w:bottom')))),
    ...(withValue('right', parseWidth(getNode(node, 'w:right')))),
    ...(withValue('start', parseWidth(getNode(node, 'w:start')))),
    ...(withValue('end', parseWidth(getNode(node, 'w:end')))),
  }

  return hasKeys(insetSet) ? insetSet : undefined
}

function parseTableLook(node: XmlNode | undefined): TableLook | undefined {
  if (node === undefined) return undefined

  const tableLook: TableLook = {
    ...(withValue('value', getAttr(node, 'w:val'))),
    ...(withValue('firstRow', parseOnOffAttr(getAttr(node, 'w:firstRow')))),
    ...(withValue('lastRow', parseOnOffAttr(getAttr(node, 'w:lastRow')))),
    ...(withValue('firstColumn', parseOnOffAttr(getAttr(node, 'w:firstColumn')))),
    ...(withValue('lastColumn', parseOnOffAttr(getAttr(node, 'w:lastColumn')))),
    ...(withValue('noHBand', parseOnOffAttr(getAttr(node, 'w:noHBand')))),
    ...(withValue('noVBand', parseOnOffAttr(getAttr(node, 'w:noVBand')))),
  }

  return hasKeys(tableLook) ? tableLook : undefined
}

function parseStyleType(value: string | undefined): StyleType | undefined {
  switch (value) {
    case 'paragraph':
    case 'character':
    case 'table':
    case 'numbering':
      return value
    default:
      return undefined
  }
}

function parseOnOffElement(value: XmlValue | undefined): OnOff | undefined {
  if (value === undefined) return undefined
  if (Array.isArray(value)) {
    return value.length === 0 ? true : parseOnOffElement(value[0])
  }
  if (isXmlNode(value)) {
    return parseOnOffAttr(getAttr(value, 'w:val')) ?? true
  }
  return parseOnOffAttr(toOptionalString(value)) ?? true
}

function parseOnOffAttr(value: string | undefined): OnOff | undefined {
  if (value === undefined) return undefined

  switch (value.toLowerCase()) {
    case '0':
    case 'false':
    case 'off':
      return false
    case '1':
    case 'true':
    case 'on':
    case '':
      return true
    default:
      return true
  }
}

function parseJustifyContent(value: string | undefined): JustifyContent | undefined {
  switch (value) {
    case 'left':
    case 'start':
      return 'start'
    case 'center':
      return 'center'
    case 'right':
    case 'end':
      return 'end'
    case 'both':
      return 'both'
    case 'distribute':
      return 'distribute'
    default:
      return undefined
  }
}

function parseTextAlignment(value: string | undefined): TextAlignment | undefined {
  switch (value) {
    case 'top':
    case 'center':
    case 'baseline':
    case 'auto':
    case 'bottom':
      return value
    default:
      return undefined
  }
}

function parseLineRule(value: string | undefined): LineRule | undefined {
  switch (value) {
    case 'auto':
    case 'exact':
    case 'atLeast':
      return value
    default:
      return undefined
  }
}

function parseUnderlineStyle(value: string | undefined): UnderlineStyle | undefined {
  switch (value) {
    case 'none':
    case 'single':
    case 'words':
    case 'double':
    case 'thick':
    case 'dotted':
    case 'dottedHeavy':
    case 'dash':
    case 'dashedHeavy':
    case 'dashLong':
    case 'dashLongHeavy':
    case 'dotDash':
    case 'dashDotHeavy':
    case 'dotDotDash':
    case 'dashDotDotHeavy':
    case 'wave':
    case 'wavyHeavy':
    case 'wavyDouble':
      return value
    default:
      return undefined
  }
}

function parseBorderStyle(value: string | undefined): Border['style'] | undefined {
  switch (value) {
    case 'nil':
    case 'none':
    case 'single':
    case 'thick':
    case 'double':
    case 'dotted':
    case 'dashed':
    case 'dotDash':
    case 'dotDotDash':
    case 'triple':
    case 'thinThickSmallGap':
    case 'thickThinSmallGap':
    case 'thinThickThinSmallGap':
    case 'thinThickMediumGap':
    case 'thickThinMediumGap':
    case 'thinThickThinMediumGap':
    case 'thinThickLargeGap':
    case 'thickThinLargeGap':
    case 'thinThickThinLargeGap':
    case 'wave':
    case 'doubleWave':
    case 'dashSmallGap':
    case 'dashDotStroked':
    case 'threeDEmboss':
    case 'threeDEngrave':
    case 'outset':
    case 'inset':
      return value
    default:
      return undefined
  }
}

function parseTabAlignment(value: string | undefined): TabAlignment | undefined {
  switch (value) {
    case 'left':
    case 'center':
    case 'right':
    case 'decimal':
    case 'bar':
    case 'clear':
    case 'start':
    case 'end':
    case 'num':
      return value
    default:
      return undefined
  }
}

function parseTabLeader(value: string | undefined): TabLeader | undefined {
  switch (value) {
    case 'none':
    case 'dot':
    case 'hyphen':
    case 'underscore':
    case 'heavy':
    case 'middleDot':
      return value
    default:
      return undefined
  }
}

function parseFontHint(value: string | undefined): FontHint | undefined {
  switch (value) {
    case 'default':
    case 'eastAsia':
    case 'cs':
    case 'hAnsi':
      return value
    default:
      return undefined
  }
}

function parseWidthType(value: string | undefined): WidthType | undefined {
  switch (value) {
    case 'auto':
    case 'dxa':
    case 'pct':
    case 'nil':
      return value
    default:
      return undefined
  }
}

function parseTableLayout(value: string | undefined): TableLayout | undefined {
  switch (value) {
    case 'fixed':
    case 'autofit':
      return value
    default:
      return undefined
  }
}

function parseFrameHorizontalAlign(
  value: string | undefined,
): FrameHorizontalAlign | undefined {
  switch (value) {
    case 'left':
    case 'center':
    case 'right':
    case 'inside':
    case 'outside':
      return value
    default:
      return undefined
  }
}

function parseFrameVerticalAlign(value: string | undefined): FrameVerticalAlign | undefined {
  switch (value) {
    case 'top':
    case 'center':
    case 'bottom':
    case 'inside':
    case 'outside':
      return value
    default:
      return undefined
  }
}

function parseFrameAnchor(value: string | undefined): FrameAnchor | undefined {
  switch (value) {
    case 'page':
    case 'margin':
    case 'text':
      return value
    default:
      return undefined
  }
}

function parseFrameWrap(value: string | undefined): FrameWrap | undefined {
  switch (value) {
    case 'none':
    case 'around':
    case 'notBeside':
    case 'through':
    case 'tight':
      return value
    default:
      return undefined
  }
}

function parseFrameDropCap(value: string | undefined): FrameDropCap | undefined {
  switch (value) {
    case 'none':
    case 'drop':
    case 'margin':
      return value
    default:
      return undefined
  }
}

function parsePageOrientation(value: string | undefined): NonNullable<SectionProps['pgSz']>['orient'] | undefined {
  switch (value) {
    case 'portrait':
    case 'landscape':
      return value
    default:
      return undefined
  }
}

function parseHeaderFooterReferenceType(
  value: string | undefined,
): NonNullable<SectionProps['headerReference']>[number]['type'] | undefined {
  switch (value) {
    case 'default':
    case 'first':
    case 'even':
      return value
    default:
      return undefined
  }
}

function parseSectionBreakType(
  value: string | undefined,
): SectionProps['type'] | undefined {
  switch (value) {
    case 'continuous':
    case 'nextPage':
    case 'nextColumn':
    case 'evenPage':
    case 'oddPage':
      return value
    default:
      return undefined
  }
}

function parseLineNumberRestart(
  value: string | undefined,
): NonNullable<SectionProps['lnNumType']>['restart'] | undefined {
  switch (value) {
    case 'continuous':
    case 'newPage':
    case 'newSection':
      return value
    default:
      return undefined
  }
}

function parseSectionVerticalAlign(
  value: string | undefined,
): SectionProps['vAlign'] | undefined {
  switch (value) {
    case 'top':
    case 'center':
    case 'both':
    case 'bottom':
      return value
    default:
      return undefined
  }
}

function parseColor(value: string | undefined): Color | undefined {
  if (value === undefined || value === '') return undefined
  if (value === 'auto') return 'auto'
  return hexColor(value)
}

function parseAliases(value: string | undefined): ReadonlyArray<string> | undefined {
  if (value === undefined || value.trim() === '') return undefined

  const aliases = value
    .split(',')
    .map((item) => item.trim())
    .filter((item) => item !== '')

  return aliases.length > 0 ? aliases : undefined
}

function getNodes(parent: XmlNode | undefined, key: string): XmlNode[] {
  return toArray(parent?.[key]).filter((value): value is XmlNode => isXmlNode(value))
}

function getNode(parent: XmlNode | undefined, key: string): XmlNode | undefined {
  return asXmlNode(parent?.[key])
}

function getValAttr(node: XmlNode | undefined): string | undefined {
  return getAttr(node, 'w:val')
}

function getAttr(node: XmlNode | undefined, key: string): string | undefined {
  if (node === undefined) return undefined
  return toOptionalString(node[`@_${key}`])
}

function parseInteger(value: string | undefined): number | undefined {
  if (value === undefined || value.trim() === '') return undefined

  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) ? parsed : undefined
}

function parseTwipAttr(node: XmlNode | undefined, key: string): ReturnType<typeof twip> | undefined {
  const value = parseInteger(getAttr(node, key))
  return value !== undefined ? twip(value) : undefined
}

function parseHalfPointAttr(
  node: XmlNode | undefined,
  key: string,
): ReturnType<typeof halfPoint> | undefined {
  const value = parseInteger(getAttr(node, key))
  return value !== undefined ? halfPoint(value) : undefined
}

function parseEighthPointAttr(
  node: XmlNode | undefined,
  key: string,
): ReturnType<typeof eighthPoint> | undefined {
  const value = parseInteger(getAttr(node, key))
  return value !== undefined ? eighthPoint(value) : undefined
}

function toArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? [...value] : [value]
}

function isXmlNode(value: XmlValue | unknown): value is XmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function asXmlNode(value: XmlValue | unknown): XmlNode | undefined {
  return isXmlNode(value) ? value : undefined
}

function toOptionalString(value: XmlValue | undefined): string | undefined {
  if (typeof value === 'string') return value
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  return undefined
}

function withValue<K extends string, V>(
  key: K,
  value: V | undefined,
) {
  if (value === undefined) return {}
  return { [key]: value }
}

function hasKeys(value: object): boolean {
  return Object.keys(value).length > 0
}
