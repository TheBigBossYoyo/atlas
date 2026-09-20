import { XMLBuilder } from 'fast-xml-parser'

import type {
  Border,
  BorderSet,
  FontSet,
  FrameProps,
  InsetSet,
  LanguageSet,
  OnOff,
  ParaProps,
  RunProps,
  SectionProps,
  Shading,
  Style,
  TableCellProps,
  TableConditionalFormat,
  TableConditionalFormatType,
  TableLook,
  TableRowProps,
  TableStyleProps,
  Width,
} from '../model'
import type { StylesPart } from '../parser/styles'

type XmlPrimitive = string | number | boolean
type XmlValue = XmlPrimitive | XmlNode | XmlValue[]

interface XmlNode {
  readonly [key: string]: XmlValue | undefined
}

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const RELATIONSHIP_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

export function writeStylesXml(stylesPart: StylesPart): string {
  const styles = Array.from(stylesPart.styles.values()).map(buildStyleXml)

  const root: XmlNode = {
    '@_xmlns:w': WORD_NAMESPACE,
    '@_xmlns:r': RELATIONSHIP_NAMESPACE,
    'w:docDefaults': buildDocDefaultsXml(stylesPart.docDefaults),
    // D19 / DXS-14: re-emit `<w:latentStyles>` verbatim (as parsed) rather
    // than silently dropping it — see `StylesPart.latentStyles`'s doc
    // comment in `parser/styles.ts`.
    ...(stylesPart.latentStyles !== undefined ? { 'w:latentStyles': stylesPart.latentStyles as XmlNode } : {}),
    ...(styles.length > 0 ? { 'w:style': styles } : {}),
  }

  return `${XML_DECLARATION}${xmlBuilder.build({ 'w:styles': root })}`
}

// Child order follows ECMA-376 `CT_PPrBase`/`CT_PPr` (§17.3.1.26), restricted
// to the members `ParaProps` models — see `documentWriter.ts`'s
// `buildParagraphPropertiesNode` doc comment for the violations an
// element-order review found in the equivalent (pre-fix) ordering there,
// which this duplicate builder shared.
export function buildParagraphPropertiesXml(
  props: ParaProps | undefined,
): Record<string, unknown> | undefined {
  if (props === undefined) return undefined

  const node: XmlNode = {
    ...buildValElement('w:pStyle', props.pStyle),
    ...withElement('w:keepNext', buildOnOffElement(props.keepNext)),
    ...withElement('w:keepLines', buildOnOffElement(props.keepLines)),
    ...withElement('w:pageBreakBefore', buildOnOffElement(props.pageBreakBefore)),
    ...withElement('w:framePr', buildFramePropsXml(props.framePr)),
    ...withElement('w:widowControl', buildOnOffElement(props.widowControl)),
    ...withElement('w:numPr', buildNumPrXml(props.numPr)),
    ...withElement('w:suppressLineNumbers', buildOnOffElement(props.suppressLineNumbers)),
    ...withElement('w:pBdr', buildBorderSetXml(props.pBdr)),
    ...withElement('w:shd', buildShadingXml(props.shd)),
    ...withElement('w:tabs', buildTabsXml(props.tabs)),
    ...withElement('w:suppressAutoHyphens', buildOnOffElement(props.suppressAutoHyphens)),
    ...withElement('w:bidi', buildOnOffElement(props.bidi)),
    ...withElement('w:spacing', buildSpacingXml(props.spacing)),
    ...withElement('w:ind', buildIndentXml(props.ind)),
    ...withElement('w:contextualSpacing', buildOnOffElement(props.contextualSpacing)),
    ...withElement('w:mirrorIndents', buildOnOffElement(props.mirrorIndents)),
    ...buildValElement('w:jc', props.jc),
    ...buildValElement('w:textAlignment', props.textAlignment),
    ...buildValElement('w:outlineLvl', props.outlineLvl),
    ...buildValElement('w:divId', props.divId),
    ...withElement('w:sectPr', buildSectionPropsXml(props.sectPr)),
  }

  return hasEntries(node) ? node : undefined
}

// Child order follows ECMA-376 `CT_RPr`/`EG_RPrBase` (§17.3.2.28), restricted
// to the members `RunProps` models — see `documentWriter.ts`'s
// `buildRunPropertiesNode` doc comment for the violations an element-order
// review found in the equivalent (pre-fix) ordering there, which this
// duplicate builder shared.
export function buildRunPropertiesXml(
  props: RunProps | undefined,
): Record<string, unknown> | undefined {
  if (props === undefined) return undefined

  const node: XmlNode = {
    ...buildValElement('w:rStyle', props.rStyle),
    ...withElement('w:rFonts', buildFontSetXml(props.rFonts)),
    ...withElement('w:b', buildOnOffElement(props.bold)),
    ...withElement('w:bCs', buildOnOffElement(props.boldCs)),
    ...withElement('w:i', buildOnOffElement(props.italic)),
    ...withElement('w:iCs', buildOnOffElement(props.italicCs)),
    ...withElement('w:caps', buildOnOffElement(props.caps)),
    ...withElement('w:smallCaps', buildOnOffElement(props.smallCaps)),
    ...withElement('w:strike', buildOnOffElement(props.strike)),
    ...withElement('w:dstrike', buildOnOffElement(props.dstrike)),
    // DOCX-12 — see `documentWriter.ts`'s `buildRunPropertiesNode` doc
    // comment for the same gap in the run-level builder this mirrors.
    ...withElement('w:outline', buildOnOffElement(props.outline)),
    ...withElement('w:emboss', buildOnOffElement(props.emboss)),
    ...withElement('w:imprint', buildOnOffElement(props.imprint)),
    ...withElement('w:vanish', buildOnOffElement(props.vanish)),
    ...withElement('w:webHidden', buildOnOffElement(props.webHidden)),
    ...buildValElement('w:color', props.color),
    ...buildValElement('w:spacing', props.spacing),
    ...buildValElement('w:w', props.charScale),
    ...buildValElement('w:kern', props.kern),
    ...buildValElement('w:position', props.position),
    ...buildValElement('w:sz', props.sz),
    ...buildValElement('w:szCs', props.szCs),
    ...buildValElement('w:highlight', props.highlight),
    ...withElement('w:u', buildUnderlineXml(props.underline)),
    ...withElement('w:bdr', buildBorderXml(props.bdr)),
    ...withElement('w:shd', buildShadingXml(props.shd)),
    ...buildValElement('w:vertAlign', props.vertAlign),
    ...withElement('w:rtl', buildOnOffElement(props.rtl)),
    ...buildValElement('w:em', props.em),
    ...withElement('w:lang', buildLanguageSetXml(props.lang)),
  }

  return hasEntries(node) ? node : undefined
}

function buildDocDefaultsXml(defaults: StylesPart['docDefaults']): XmlNode {
  return {
    ...withElement(
      'w:rPrDefault',
      defaults.rPr === undefined ? undefined : { 'w:rPr': buildRunPropertiesXml(defaults.rPr) ?? {} },
    ),
    ...withElement(
      'w:pPrDefault',
      defaults.pPr === undefined
        ? undefined
        : { 'w:pPr': buildParagraphPropertiesXml(defaults.pPr) ?? {} },
    ),
  }
}

function buildStyleXml(style: Style): XmlNode {
  const paragraph = buildStyleParagraphProps(style)
  const aliases = serializeAliases(style.aliases)
  const tblStylePr = buildTableConditionalFormatsXml(style.conditionalFormats)

  return {
    '@_w:type': style.type,
    '@_w:styleId': style.id,
    ...withAttribute('@_w:default', buildOnOffAttribute(style.isDefault)),
    ...withAttribute('@_w:customStyle', buildOnOffAttribute(style.custom)),
    ...buildValElement('w:name', style.name),
    ...buildValElement('w:aliases', aliases),
    ...buildValElement('w:basedOn', style.basedOn),
    ...buildValElement('w:next', style.next),
    ...buildValElement('w:link', style.linked),
    // CT_Style (§17.7.4.17) orders `hidden` before `uiPriority`.
    ...withElement('w:hidden', buildOnOffElement(style.hidden)),
    ...buildValElement('w:uiPriority', style.uiPriority),
    ...withElement('w:semiHidden', buildOnOffElement(style.semiHidden)),
    ...withElement('w:unhideWhenUsed', buildOnOffElement(style.unhideWhenUsed)),
    ...withElement('w:qFormat', buildOnOffElement(style.qFormat)),
    ...withElement('w:locked', buildOnOffElement(style.locked)),
    ...withElement('w:pPr', buildParagraphPropertiesXml(paragraph)),
    ...withElement('w:rPr', buildRunPropertiesXml(style.run)),
    ...withElement('w:tblPr', buildTableStylePropertiesXml(style.table)),
    ...(tblStylePr.length > 0 ? { 'w:tblStylePr': tblStylePr } : {}),
  }
}

/**
 * D7 / DXP-06, DXL-08, DXS-05: re-emits one `<w:tblStylePr>` per
 * conditional-formatting region the style carries — closes the "loses all
 * table formatting on save" data loss for any table style using Word's
 * Table Style Options (banded rows/columns, header/total row, first/last
 * column).
 */
function buildTableConditionalFormatsXml(
  conditionalFormats: ReadonlyMap<TableConditionalFormatType, TableConditionalFormat> | undefined,
): XmlNode[] {
  if (conditionalFormats === undefined) return []

  return Array.from(conditionalFormats.entries()).map(([type, format]) => ({
    '@_w:type': type,
    ...withElement('w:pPr', buildParagraphPropertiesXml(format.paragraph)),
    ...withElement('w:rPr', buildRunPropertiesXml(format.run)),
    ...withElement('w:tblPr', buildTableStylePropertiesXml(format.table)),
    ...withElement('w:trPr', buildTableConditionalRowPropertiesXml(format.row)),
    ...withElement('w:tcPr', buildTableConditionalCellPropertiesXml(format.cell)),
  }))
}

// Child order follows ECMA-376 `CT_TrPrBase` (§17.4.83): `cantSplit`
// precedes `trHeight`.
function buildTableConditionalRowPropertiesXml(row: TableRowProps | undefined): XmlNode | undefined {
  if (row === undefined) return undefined

  const node: XmlNode = {
    ...withElement('w:cantSplit', buildOnOffElement(row.cantSplit)),
    ...(row.trHeight !== undefined
      ? {
          'w:trHeight': {
            '@_w:val': toOptionalString(row.trHeight.val),
            ...withAttribute('@_w:hRule', row.trHeight.hRule),
          },
        }
      : {}),
    ...withElement('w:tblHeader', buildOnOffElement(row.tblHeader)),
    ...buildValElement('w:jc', row.jc),
  }

  return hasEntries(node) ? node : undefined
}

// Child order follows ECMA-376 `CT_TcPrBase` (§17.4.70): `noWrap` precedes
// `tcMar`/`vAlign`.
function buildTableConditionalCellPropertiesXml(cell: TableCellProps | undefined): XmlNode | undefined {
  if (cell === undefined) return undefined

  const node: XmlNode = {
    ...withElement('w:tcW', buildWidthXml(cell.tcW)),
    ...buildValElement('w:gridSpan', cell.gridSpan),
    ...(cell.vMerge !== undefined ? { 'w:vMerge': withAttribute('@_w:val', cell.vMerge) } : {}),
    ...withElement('w:tcBorders', buildBorderSetXml(cell.tcBorders)),
    ...withElement('w:shd', buildShadingXml(cell.shd)),
    ...withElement('w:noWrap', buildOnOffElement(cell.noWrap)),
    ...withElement('w:tcMar', buildInsetSetXml(cell.tcMar)),
    ...buildValElement('w:vAlign', cell.vAlign),
    ...withElement('w:hideMark', buildOnOffElement(cell.hideMark)),
  }

  return hasEntries(node) ? node : undefined
}

function buildStyleParagraphProps(style: Style): ParaProps | undefined {
  if (style.numbering === undefined) return style.paragraph

  const numPr = {
    ...(style.paragraph?.numPr ?? {}),
    ...(style.paragraph?.numPr?.numId === undefined ? { numId: style.numbering.numId } : {}),
    ...(style.paragraph?.numPr?.ilvl === undefined && style.numbering.ilvl !== undefined
      ? { ilvl: style.numbering.ilvl }
      : {}),
  }

  if (style.paragraph === undefined) {
    return { numPr }
  }

  return {
    ...style.paragraph,
    ...(hasEntries(numPr as unknown as XmlNode) ? { numPr } : {}),
  }
}

function buildNumPrXml(numPr: ParaProps['numPr']): XmlNode | undefined {
  if (numPr === undefined) return undefined

  const node: XmlNode = {
    ...buildValElement('w:ilvl', numPr.ilvl),
    ...buildValElement('w:numId', numPr.numId),
  }

  return hasEntries(node) ? node : undefined
}

function buildSpacingXml(spacing: ParaProps['spacing']): XmlNode | undefined {
  if (spacing === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:before', toOptionalString(spacing.before)),
    ...withAttribute('@_w:after', toOptionalString(spacing.after)),
    ...withAttribute('@_w:beforeAutospacing', buildOnOffAttribute(spacing.beforeAutospacing)),
    ...withAttribute('@_w:afterAutospacing', buildOnOffAttribute(spacing.afterAutospacing)),
    ...withAttribute('@_w:line', toOptionalString(spacing.line)),
    ...withAttribute('@_w:lineRule', spacing.lineRule),
  }

  return hasEntries(node) ? node : undefined
}

function buildIndentXml(indent: ParaProps['ind']): XmlNode | undefined {
  if (indent === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:left', toOptionalString(indent.left)),
    ...withAttribute('@_w:right', toOptionalString(indent.right)),
    ...withAttribute('@_w:firstLine', toOptionalString(indent.firstLine)),
    ...withAttribute('@_w:hanging', toOptionalString(indent.hanging)),
    ...withAttribute('@_w:start', toOptionalString(indent.start)),
    ...withAttribute('@_w:end', toOptionalString(indent.end)),
  }

  return hasEntries(node) ? node : undefined
}

function buildTabsXml(tabs: ParaProps['tabs']): XmlNode | undefined {
  if (tabs === undefined) return undefined

  const items = tabs.items.map((tab) => ({
    '@_w:pos': String(tab.position),
    ...withAttribute('@_w:val', tab.alignment),
    ...withAttribute('@_w:leader', tab.leader),
  }))

  const node: XmlNode = {
    ...withAttribute('@_w:defaultTabStop', toOptionalString(tabs.defaultTabStop)),
    ...(items.length > 0 ? { 'w:tab': items } : {}),
  }

  return hasEntries(node) ? node : undefined
}

function buildBorderSetXml(borderSet: BorderSet | undefined): XmlNode | undefined {
  if (borderSet === undefined) return undefined

  const node: XmlNode = {
    ...withElement('w:top', buildBorderXml(borderSet.top)),
    ...withElement('w:left', buildBorderXml(borderSet.left)),
    ...withElement('w:bottom', buildBorderXml(borderSet.bottom)),
    ...withElement('w:right', buildBorderXml(borderSet.right)),
    ...withElement('w:start', buildBorderXml(borderSet.start)),
    ...withElement('w:end', buildBorderXml(borderSet.end)),
    ...withElement('w:between', buildBorderXml(borderSet.between)),
    ...withElement('w:bar', buildBorderXml(borderSet.bar)),
    ...withElement('w:insideH', buildBorderXml(borderSet.insideH)),
    ...withElement('w:insideV', buildBorderXml(borderSet.insideV)),
  }

  return hasEntries(node) ? node : undefined
}

function buildBorderXml(border: Border | undefined): XmlNode | undefined {
  if (border === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:val', border.style),
    ...withAttribute('@_w:color', border.color),
    ...withAttribute('@_w:sz', toOptionalString(border.size)),
    ...withAttribute('@_w:space', toOptionalString(border.space)),
    ...withAttribute('@_w:shadow', buildOnOffAttribute(border.shadow)),
    ...withAttribute('@_w:frame', buildOnOffAttribute(border.frame)),
  }

  return hasEntries(node) ? node : undefined
}

function buildShadingXml(shading: Shading | undefined): XmlNode | undefined {
  if (shading === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:fill', shading.fill),
    ...withAttribute('@_w:color', shading.color),
    ...withAttribute('@_w:val', shading.pattern),
  }

  return hasEntries(node) ? node : undefined
}

function buildFramePropsXml(frameProps: FrameProps | undefined): XmlNode | undefined {
  if (frameProps === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:w', toOptionalString(frameProps.width)),
    ...withAttribute('@_w:h', toOptionalString(frameProps.height)),
    ...withAttribute('@_w:x', toOptionalString(frameProps.x)),
    ...withAttribute('@_w:y', toOptionalString(frameProps.y)),
    ...withAttribute('@_w:xAlign', frameProps.xAlign),
    ...withAttribute('@_w:yAlign', frameProps.yAlign),
    ...withAttribute('@_w:hAnchor', frameProps.hAnchor),
    ...withAttribute('@_w:vAnchor', frameProps.vAnchor),
    ...withAttribute('@_w:wrap', frameProps.wrap),
    ...withAttribute('@_w:lines', toOptionalString(frameProps.lines)),
    ...withAttribute('@_w:hSpace', toOptionalString(frameProps.hSpace)),
    ...withAttribute('@_w:vSpace', toOptionalString(frameProps.vSpace)),
    ...withAttribute('@_w:dropCap', frameProps.dropCap),
    ...withAttribute('@_w:lockAnchor', buildOnOffAttribute(frameProps.lockAnchor)),
  }

  return hasEntries(node) ? node : undefined
}

// Child order follows ECMA-376 `EG_SectPrContents` (§17.6.17), restricted to
// the members modeled here — see `documentWriter.ts`'s
// `buildSectionPropertiesNode` doc comment for the full member list and the
// violations an element-order review found in the equivalent (pre-fix)
// ordering there.
function buildSectionPropsXml(sectionProps: SectionProps | undefined): XmlNode | undefined {
  if (sectionProps === undefined) return undefined

  const headerReference = (sectionProps.headerReference ?? []).map((item) => ({
    '@_r:id': item.id,
    '@_w:type': item.type,
  }))

  const footerReference = (sectionProps.footerReference ?? []).map((item) => ({
    '@_r:id': item.id,
    '@_w:type': item.type,
  }))

  const node: XmlNode = {
    ...(headerReference.length > 0 ? { 'w:headerReference': headerReference } : {}),
    ...(footerReference.length > 0 ? { 'w:footerReference': footerReference } : {}),
    ...buildValElement('w:type', sectionProps.type),
    ...withElement('w:pgSz', buildPageSizeXml(sectionProps.pgSz)),
    ...withElement('w:pgMar', buildPageMarginsXml(sectionProps.pgMar)),
    ...withElement('w:lnNumType', buildLineNumberTypeXml(sectionProps.lnNumType)),
    ...withElement('w:pgNumType', buildPageNumberTypeXml(sectionProps.pgNumType)),
    ...withElement('w:cols', buildSectionColumnsXml(sectionProps.cols)),
    ...buildValElement('w:vAlign', sectionProps.vAlign),
    ...withElement('w:titlePg', buildOnOffElement(sectionProps.titlePg)),
  }

  return hasEntries(node) ? node : undefined
}

function buildPageSizeXml(pageSize: SectionProps['pgSz']): XmlNode | undefined {
  if (pageSize === undefined) return undefined

  return {
    '@_w:w': String(pageSize.w),
    '@_w:h': String(pageSize.h),
    ...withAttribute('@_w:orient', pageSize.orient),
  }
}

function buildPageMarginsXml(pageMargins: SectionProps['pgMar']): XmlNode | undefined {
  if (pageMargins === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:top', toOptionalString(pageMargins.top)),
    ...withAttribute('@_w:right', toOptionalString(pageMargins.right)),
    ...withAttribute('@_w:bottom', toOptionalString(pageMargins.bottom)),
    ...withAttribute('@_w:left', toOptionalString(pageMargins.left)),
    ...withAttribute('@_w:header', toOptionalString(pageMargins.header)),
    ...withAttribute('@_w:footer', toOptionalString(pageMargins.footer)),
    ...withAttribute('@_w:gutter', toOptionalString(pageMargins.gutter)),
  }

  return hasEntries(node) ? node : undefined
}

function buildSectionColumnsXml(columns: SectionProps['cols']): XmlNode | undefined {
  if (columns === undefined) return undefined

  const col = columns.col.map((item) => ({
    ...withAttribute('@_w:w', toOptionalString(item.w)),
    ...withAttribute('@_w:space', toOptionalString(item.space)),
  }))

  const node: XmlNode = {
    ...withAttribute('@_w:num', toOptionalString(columns.num)),
    ...withAttribute('@_w:space', toOptionalString(columns.space)),
    ...withAttribute('@_w:sep', buildOnOffAttribute(columns.sep)),
    ...withAttribute('@_w:equalWidth', buildOnOffAttribute(columns.equalWidth)),
    ...(col.length > 0 ? { 'w:col': col } : {}),
  }

  return hasEntries(node) ? node : undefined
}

function buildPageNumberTypeXml(pageNumberType: SectionProps['pgNumType']): XmlNode | undefined {
  if (pageNumberType === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:start', toOptionalString(pageNumberType.start)),
    ...withAttribute('@_w:fmt', pageNumberType.fmt),
  }

  return hasEntries(node) ? node : undefined
}

function buildLineNumberTypeXml(lineNumberType: SectionProps['lnNumType']): XmlNode | undefined {
  if (lineNumberType === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:countBy', toOptionalString(lineNumberType.countBy)),
    ...withAttribute('@_w:start', toOptionalString(lineNumberType.start)),
    ...withAttribute('@_w:distance', toOptionalString(lineNumberType.distance)),
    ...withAttribute('@_w:restart', lineNumberType.restart),
  }

  return hasEntries(node) ? node : undefined
}

function buildUnderlineXml(underline: RunProps['underline']): XmlNode | undefined {
  if (underline === undefined) return undefined

  return {
    '@_w:val': underline.style,
    ...withAttribute('@_w:color', underline.color),
  }
}

// DOCX-1 — same gap as `documentWriter.ts`'s `buildFontSetElement` (see its
// doc comment): `parser/styles.ts`'s `parseFontSet` has always read the
// theme-reference attributes onto `FontSet`, but this builder (used for
// `w:docDefaults`/named-style/table-conditional-format `w:rFonts`) only ever
// emitted the literal ones, silently dropping a style's theme font — often
// Normal's own — on every save, and dropping `w:rFonts` entirely when it
// carried nothing but a theme reference.
function buildFontSetXml(fontSet: FontSet | undefined): XmlNode | undefined {
  if (fontSet === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:ascii', fontSet.ascii),
    ...withAttribute('@_w:hAnsi', fontSet.hAnsi),
    ...withAttribute('@_w:cs', fontSet.cs),
    ...withAttribute('@_w:eastAsia', fontSet.eastAsia),
    ...withAttribute('@_w:hint', fontSet.hint),
    ...withAttribute('@_w:asciiTheme', fontSet.asciiTheme),
    ...withAttribute('@_w:hAnsiTheme', fontSet.hAnsiTheme),
    ...withAttribute('@_w:cstheme', fontSet.csTheme),
    ...withAttribute('@_w:eastAsiaTheme', fontSet.eastAsiaTheme),
  }

  return hasEntries(node) ? node : undefined
}

function buildLanguageSetXml(languageSet: LanguageSet | undefined): XmlNode | undefined {
  if (languageSet === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:val', languageSet.value),
    ...withAttribute('@_w:eastAsia', languageSet.eastAsia),
    ...withAttribute('@_w:bidi', languageSet.bidi),
  }

  return hasEntries(node) ? node : undefined
}

// Child order follows ECMA-376 `CT_TblPrBase` (§17.4.60), restricted to the
// members `TableStyleProps` models.
function buildTableStylePropertiesXml(tableStyleProps: TableStyleProps | undefined): XmlNode | undefined {
  if (tableStyleProps === undefined) return undefined

  const node: XmlNode = {
    ...buildValElement('w:tblStyleRowBandSize', tableStyleProps.rowBandSize),
    ...buildValElement('w:tblStyleColBandSize', tableStyleProps.colBandSize),
    ...withElement('w:tblW', buildWidthXml(tableStyleProps.width)),
    ...buildValElement('w:jc', tableStyleProps.justification),
    ...withElement('w:tblInd', buildWidthXml(tableStyleProps.indent)),
    ...withElement('w:tblBorders', buildBorderSetXml(tableStyleProps.borders)),
    ...withElement('w:shd', buildShadingXml(tableStyleProps.shading)),
    ...buildValElement('w:tblLayout', tableStyleProps.layout),
    ...withElement('w:tblCellMar', buildInsetSetXml(tableStyleProps.cellMargin)),
    ...withElement('w:tblLook', buildTableLookXml(tableStyleProps.look)),
  }

  return hasEntries(node) ? node : undefined
}

function buildWidthXml(width: Width | undefined): XmlNode | undefined {
  if (width === undefined) return undefined

  return {
    '@_w:type': width.type,
    ...withAttribute('@_w:w', toOptionalString(width.value)),
  }
}

function buildInsetSetXml(insetSet: InsetSet | undefined): XmlNode | undefined {
  if (insetSet === undefined) return undefined

  const node: XmlNode = {
    ...withElement('w:top', buildWidthXml(insetSet.top)),
    ...withElement('w:left', buildWidthXml(insetSet.left)),
    ...withElement('w:bottom', buildWidthXml(insetSet.bottom)),
    ...withElement('w:right', buildWidthXml(insetSet.right)),
    ...withElement('w:start', buildWidthXml(insetSet.start)),
    ...withElement('w:end', buildWidthXml(insetSet.end)),
  }

  return hasEntries(node) ? node : undefined
}

function buildTableLookXml(tableLook: TableLook | undefined): XmlNode | undefined {
  if (tableLook === undefined) return undefined

  const node: XmlNode = {
    ...withAttribute('@_w:val', tableLook.value),
    ...withAttribute('@_w:firstRow', buildOnOffAttribute(tableLook.firstRow)),
    ...withAttribute('@_w:lastRow', buildOnOffAttribute(tableLook.lastRow)),
    ...withAttribute('@_w:firstColumn', buildOnOffAttribute(tableLook.firstColumn)),
    ...withAttribute('@_w:lastColumn', buildOnOffAttribute(tableLook.lastColumn)),
    ...withAttribute('@_w:noHBand', buildOnOffAttribute(tableLook.noHBand)),
    ...withAttribute('@_w:noVBand', buildOnOffAttribute(tableLook.noVBand)),
  }

  return hasEntries(node) ? node : undefined
}

function buildOnOffElement(value: OnOff | undefined): XmlNode | undefined {
  if (value === undefined) return undefined
  return value ? {} : { '@_w:val': '0' }
}

function buildOnOffAttribute(value: OnOff | undefined): string | undefined {
  if (value === undefined) return undefined
  return value ? '1' : '0'
}

function buildValElement(name: string, value: string | number | undefined): XmlNode {
  return value === undefined ? {} : { [name]: { '@_w:val': String(value) } }
}

function withElement(name: string, value: Record<string, unknown> | XmlNode | undefined): XmlNode {
  return value === undefined ? {} : { [name]: value as XmlNode }
}

function withAttribute(name: string, value: string | undefined): XmlNode {
  return value === undefined ? {} : { [name]: value }
}

function hasEntries(node: XmlNode): boolean {
  return Object.keys(node).length > 0
}

function serializeAliases(aliases: ReadonlyArray<string> | undefined): string | undefined {
  if (aliases === undefined) return undefined

  const values = aliases.map((alias) => alias.trim()).filter((alias) => alias.length > 0)
  return values.length > 0 ? values.join(',') : undefined
}

function toOptionalString(value: string | number | undefined): string | undefined {
  return value === undefined ? undefined : String(value)
}
