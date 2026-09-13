/**
 * Atlas — DOCX document parser (Wave A.3)
 *
 * Parses `word/document.xml` into the immutable AST defined in Wave A.2.
 * The parser keeps OOXML child order, extracts the supported block/inline
 * surface, and falls back to `UnknownNode` for unsupported structural nodes.
 */

import { XMLBuilder, XMLParser } from 'fast-xml-parser'

import {
  eighthPoint,
  halfPoint,
  hexColor,
  pct,
  twip,
  type Block,
  type Bookmark,
  type Border,
  type BorderSet,
  type BreakClear,
  type BreakNode,
  type BreakType,
  type Color,
  type Comment,
  type CommentRange,
  type CommentReference,
  type DelRevision,
  type Document as DocxDocument,
  type Drawing,
  type Endnote,
  type EndnoteReference,
  type Footer,
  type FooterReference,
  type FontSet,
  type Footnote,
  type FootnoteReference,
  type FrameProps,
  type Header,
  type HeaderReference,
  type HighlightColor,
  type Hyperlink,
  type HyperlinkChild,
  type Indent,
  type InsRevision,
  type InsetSet,
  type JustifyContent,
  type LanguageSet,
  type LineNumberType,
  type NumPr,
  type OnOff,
  type PageNumberType,
  type ParaProps,
  type ParagraphChild,
  type Run,
  type RunChild,
  type RunProps,
  type Section,
  type SectionColumn,
  type SectionProps,
  type Shading,
  type Spacing,
  type Table,
  type TableCell,
  type TableCellMerge,
  type TableCellProps,
  type TableLook,
  type TableProps,
  type TableRow,
  type TableRowChild,
  type TableRowHeight,
  type TableRowProps,
  type TextNode,
  type Underline,
  type UnknownNode,
  type VerticalAlign,
  type Width,
} from '../model'

interface OrderedXmlNode {
  readonly ':@'?: XmlAttributes
  readonly '#text'?: string
  readonly [name: string]: OrderedXmlNode[] | XmlAttributes | string | undefined
}

interface XmlAttributes {
  readonly [name: string]: string | undefined
}

type Mutable<T> = {
  -readonly [K in keyof T]: T[K]
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
})

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  suppressEmptyNode: true,
})

export function parseDocument(xml: string): DocxDocument {
  const raw = xmlParser.parse(xml) as OrderedXmlNode[]

  const documentElement = findElement(raw, 'w:document')
  const bodyElement = child(documentElement, 'w:body')
  const body = parseBody(bodyElement)

  return {
    kind: 'document',
    sections: buildSections(body.blocks, body.sectionProps),
    styles: Object.freeze(new Map()),
    numbering: Object.freeze(new Map()),
    comments: Object.freeze(new Map<string, Comment>()),
    footnotes: Object.freeze(new Map<string, Footnote>()),
    endnotes: Object.freeze(new Map<string, Endnote>()),
    headers: Object.freeze(new Map<string, Header>()),
    footers: Object.freeze(new Map<string, Footer>()),
  }
}

function parseBody(element: OrderedXmlNode | undefined): {
  readonly blocks: ReadonlyArray<Block>
  readonly sectionProps?: SectionProps
} {
  const blocks: Block[] = []
  let sectionProps: SectionProps | undefined

  for (const entry of nodeChildren(element)) {
    if (isIgnorableText(entry)) {
      continue
    }

    switch (nodeName(entry)) {
      case 'w:p':
        blocks.push(parseParagraph(entry))
        break
      case 'w:tbl':
        blocks.push(parseTable(entry))
        break
      case 'w:sectPr':
        sectionProps = parseSectionProps(entry)
        break
      default:
        blocks.push(parseUnknownNode(entry))
        break
    }
  }

  return { blocks, ...(sectionProps !== undefined ? { sectionProps } : {}) }
}

function buildSections(
  blocks: ReadonlyArray<Block>,
  trailingSectionProps: SectionProps | undefined,
): ReadonlyArray<Section> {
  const sections: Section[] = []
  let currentBlocks: Block[] = []

  for (const block of blocks) {
    currentBlocks.push(block)

    const sectionProps = block.kind === 'paragraph' ? block.props?.sectPr : undefined
    if (sectionProps !== undefined) {
      sections.push({
        kind: 'section',
        props: sectionProps,
        blocks: currentBlocks,
      })
      currentBlocks = []
    }
  }

  if (
    currentBlocks.length > 0 ||
    sections.length === 0 ||
    trailingSectionProps !== undefined
  ) {
    sections.push({
      kind: 'section',
      props: trailingSectionProps ?? {},
      blocks: currentBlocks,
    })
  }

  return sections
}

function parseParagraph(element: OrderedXmlNode): Block {
  const props = parseParaProps(child(element, 'w:pPr'))
  const children: ParagraphChild[] = []

  for (const entry of nodeChildren(element)) {
    if (isIgnorableText(entry)) {
      continue
    }

    const child = parseParagraphChild(entry)
    if (child !== null) {
      children.push(child)
    }
  }

  return {
    kind: 'paragraph',
    ...(props !== undefined ? { props } : {}),
    children,
  }
}

function parseParagraphChild(element: OrderedXmlNode): ParagraphChild | null {
  switch (nodeName(element)) {
    case 'w:pPr':
      return null
    case 'w:r':
      return parseRun(element)
    case 'w:hyperlink':
      return parseHyperlink(element)
    case 'w:bookmarkStart':
      return parseBookmark(element, 'start')
    case 'w:bookmarkEnd':
      return parseBookmark(element, 'end')
    case 'w:commentRangeStart':
      return parseCommentRange(element, 'start')
    case 'w:commentRangeEnd':
      return parseCommentRange(element, 'end')
    case 'w:commentReference':
      return parseCommentReference(element)
    case 'w:footnoteReference':
      return parseFootnoteReference(element)
    case 'w:endnoteReference':
      return parseEndnoteReference(element)
    case 'w:ins':
      return parseRevision(element, 'ins')
    case 'w:del':
      return parseRevision(element, 'del')
    default:
      return parseUnknownNode(element)
  }
}

function parseRun(element: OrderedXmlNode): Run {
  const props = parseRunProps(child(element, 'w:rPr'))
  const children: RunChild[] = []

  for (const entry of nodeChildren(element)) {
    if (isIgnorableText(entry) || nodeName(entry) === 'w:rPr') {
      continue
    }

    switch (nodeName(entry)) {
      case 'w:t':
        children.push(parseTextNode(entry))
        break
      case 'w:delText':
        children.push(parseTextNode(entry))
        break
      case 'w:tab':
        children.push({ kind: 'tab' })
        break
      case 'w:br':
        children.push(parseBreak(entry))
        break
      case 'w:drawing': {
        const drawing = parseDrawing(entry)
        children.push(drawing)
        break
      }
      case 'w:commentReference':
        children.push(parseCommentReference(entry))
        break
      case 'w:footnoteReference':
        children.push(parseFootnoteReference(entry))
        break
      case 'w:endnoteReference':
        children.push(parseEndnoteReference(entry))
        break
      default:
        children.push(parseUnknownNode(entry))
        break
    }
  }

  return {
    kind: 'run',
    ...(props !== undefined ? { props } : {}),
    children,
  }
}

function parseHyperlink(element: OrderedXmlNode): Hyperlink {
  const relationshipId = attr(element, 'r:id')
  const anchor = attr(element, 'w:anchor')
  const tooltip = attr(element, 'w:tooltip')
  const targetFrame = attr(element, 'w:tgtFrame')
  const history = parseOnOff(attr(element, 'w:history'))
  const children: HyperlinkChild[] = []

  for (const entry of nodeChildren(element)) {
    if (isIgnorableText(entry)) {
      continue
    }

    switch (nodeName(entry)) {
      case 'w:r':
        children.push(parseRun(entry))
        break
      case 'w:bookmarkStart':
        children.push(parseBookmark(entry, 'start'))
        break
      case 'w:bookmarkEnd':
        children.push(parseBookmark(entry, 'end'))
        break
      case 'w:commentRangeStart':
        children.push(parseCommentRange(entry, 'start'))
        break
      case 'w:commentRangeEnd':
        children.push(parseCommentRange(entry, 'end'))
        break
      case 'w:commentReference':
        children.push(parseCommentReference(entry))
        break
      case 'w:footnoteReference':
        children.push(parseFootnoteReference(entry))
        break
      case 'w:endnoteReference':
        children.push(parseEndnoteReference(entry))
        break
      default:
        children.push(parseUnknownNode(entry))
        break
    }
  }

  return {
    kind: 'hyperlink',
    ...(relationshipId !== undefined ? { relationshipId } : {}),
    ...(anchor !== undefined ? { anchor } : {}),
    ...(tooltip !== undefined ? { tooltip } : {}),
    ...(targetFrame !== undefined ? { targetFrame } : {}),
    ...(history !== undefined ? { history } : {}),
    children,
  }
}

function parseTextNode(element: OrderedXmlNode): TextNode {
  const preserveSpace = attr(element, 'xml:space')
  return {
    kind: 'text',
    value: textValue(element),
    ...(preserveSpace !== undefined
      ? { preserveSpace: preserveSpace === 'preserve' }
      : {}),
  }
}

function parseRevision(
  element: OrderedXmlNode,
  variant: 'ins',
): InsRevision
function parseRevision(
  element: OrderedXmlNode,
  variant: 'del',
): DelRevision
function parseRevision(
  element: OrderedXmlNode,
  variant: 'ins' | 'del',
): InsRevision | DelRevision {
  const id = attr(element, 'w:id') ?? ''
  const author = attr(element, 'w:author')
  const date = attr(element, 'w:date')
  const children: ParagraphChild[] = []

  for (const entry of nodeChildren(element)) {
    if (isIgnorableText(entry)) {
      continue
    }

    const child = parseParagraphChild(entry)
    if (child !== null) {
      children.push(child)
    }
  }

  if (variant === 'ins') {
    return {
      kind: 'ins-revision',
      id,
      ...(author !== undefined ? { author } : {}),
      ...(date !== undefined ? { date } : {}),
      children: freezeRevisionChildren(children),
    }
  }

  return {
    kind: 'del-revision',
    id,
    ...(author !== undefined ? { author } : {}),
    ...(date !== undefined ? { date } : {}),
    children: freezeRevisionChildren(children),
  }
}

function freezeRevisionChildren(
  children: ReadonlyArray<ParagraphChild>,
): ReadonlyArray<ParagraphChild> & ReadonlyArray<Run> {
  return Object.freeze(children.slice()) as ReadonlyArray<ParagraphChild> & ReadonlyArray<Run>
}

function parseBreak(element: OrderedXmlNode): BreakNode {
  const breakType = parseBreakType(attr(element, 'w:type'))
  const clear = parseBreakClear(attr(element, 'w:clear'))

  return {
    kind: 'break',
    ...(breakType !== undefined ? { breakType } : {}),
    ...(clear !== undefined ? { clear } : {}),
  }
}

function parseDrawing(element: OrderedXmlNode): Drawing | UnknownNode {
  const inline = child(element, 'wp:inline')
  const anchor = child(element, 'wp:anchor')
  const layoutElement = inline ?? anchor

  if (layoutElement === undefined) {
    return parseUnknownNode(element)
  }

  const docPr = findDescendant(layoutElement, 'wp:docPr')
  const blip = findDescendant(layoutElement, 'a:blip')
  const extentElement = findDescendant(layoutElement, 'wp:extent')
  const relationshipId =
    attr(blip, 'r:embed') ?? attr(blip, 'r:link')
  const title = attr(docPr, 'title')
  const description = attr(docPr, 'descr')
  const name = attr(docPr, 'name')
  const extent =
    extentElement !== undefined
      ? parseDrawingExtent(extentElement)
      : undefined

  return {
    kind: 'drawing',
    layout: inline !== undefined ? 'inline' : 'anchor',
    ...(relationshipId !== undefined ? { relationshipId } : {}),
    ...(title !== undefined ? { title } : {}),
    ...(description !== undefined ? { description } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(extent !== undefined ? { extent } : {}),
  }
}

function parseDrawingExtent(element: OrderedXmlNode): {
  readonly cx: number
  readonly cy: number
} | undefined {
  const cx = parseInteger(attr(element, 'cx'))
  const cy = parseInteger(attr(element, 'cy'))

  if (cx === undefined || cy === undefined) {
    return undefined
  }

  return { cx, cy }
}

function parseBookmark(
  element: OrderedXmlNode,
  boundary: 'start' | 'end',
): Bookmark {
  const name = attr(element, 'w:name')
  const colFirst = parseInteger(attr(element, 'w:colFirst'))
  const colLast = parseInteger(attr(element, 'w:colLast'))

  return {
    kind: 'bookmark',
    id: attr(element, 'w:id') ?? '',
    boundary,
    ...(name !== undefined ? { name } : {}),
    ...(colFirst !== undefined ? { colFirst } : {}),
    ...(colLast !== undefined ? { colLast } : {}),
  }
}

function parseCommentRange(
  element: OrderedXmlNode,
  boundary: 'start' | 'end',
): CommentRange {
  return {
    kind: 'comment-range',
    id: attr(element, 'w:id') ?? '',
    boundary,
  }
}

function parseCommentReference(element: OrderedXmlNode): CommentReference {
  return {
    kind: 'comment-reference',
    id: attr(element, 'w:id') ?? '',
  }
}

function parseFootnoteReference(element: OrderedXmlNode): FootnoteReference {
  const customMarkFollows = parseOnOff(attr(element, 'w:customMarkFollows'))

  return {
    kind: 'footnote-reference',
    id: attr(element, 'w:id') ?? '',
    ...(customMarkFollows !== undefined ? { customMarkFollows } : {}),
  }
}

function parseEndnoteReference(element: OrderedXmlNode): EndnoteReference {
  const customMarkFollows = parseOnOff(attr(element, 'w:customMarkFollows'))

  return {
    kind: 'endnote-reference',
    id: attr(element, 'w:id') ?? '',
    ...(customMarkFollows !== undefined ? { customMarkFollows } : {}),
  }
}

function parseTable(element: OrderedXmlNode): Table {
  const props = parseTableProps(child(element, 'w:tblPr'))
  const rows: Array<TableRow | UnknownNode> = []

  for (const entry of nodeChildren(element)) {
    if (
      isIgnorableText(entry) ||
      nodeName(entry) === 'w:tblPr' ||
      nodeName(entry) === 'w:tblGrid'
    ) {
      continue
    }

    switch (nodeName(entry)) {
      case 'w:tr':
        rows.push(parseTableRow(entry))
        break
      default:
        rows.push(parseUnknownNode(entry))
        break
    }
  }

  return {
    kind: 'table',
    ...(props !== undefined ? { props } : {}),
    rows,
  }
}

function parseTableRow(element: OrderedXmlNode): TableRow {
  const props = parseTableRowProps(child(element, 'w:trPr'))
  const cells: TableRowChild[] = []

  for (const entry of nodeChildren(element)) {
    if (isIgnorableText(entry) || nodeName(entry) === 'w:trPr') {
      continue
    }

    switch (nodeName(entry)) {
      case 'w:tc':
        cells.push(parseTableCell(entry))
        break
      default:
        cells.push(parseUnknownNode(entry))
        break
    }
  }

  return {
    kind: 'table-row',
    ...(props !== undefined ? { props } : {}),
    cells,
  }
}

function parseTableCell(element: OrderedXmlNode): TableCell {
  const props = parseTableCellProps(child(element, 'w:tcPr'))
  const blocks: Block[] = []

  for (const entry of nodeChildren(element)) {
    if (isIgnorableText(entry) || nodeName(entry) === 'w:tcPr') {
      continue
    }

    switch (nodeName(entry)) {
      case 'w:p':
        blocks.push(parseParagraph(entry))
        break
      case 'w:tbl':
        blocks.push(parseTable(entry))
        break
      default:
        blocks.push(parseUnknownNode(entry))
        break
    }
  }

  return {
    kind: 'table-cell',
    ...(props !== undefined ? { props } : {}),
    blocks,
  }
}

function parseRunProps(element: OrderedXmlNode | undefined): RunProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const props: Mutable<RunProps> = {}
  const rStyle = attr(child(element, 'w:rStyle'), 'w:val')
  const bold = parseToggleElement(child(element, 'w:b'))
  const italic = parseToggleElement(child(element, 'w:i'))
  const underline = parseUnderline(child(element, 'w:u'))
  const strike = parseToggleElement(child(element, 'w:strike'))
  const dstrike = parseToggleElement(child(element, 'w:dstrike'))
  const vertAlign = parseVerticalAlign(attr(child(element, 'w:vertAlign'), 'w:val'))
  const color = parseColor(attr(child(element, 'w:color'), 'w:val'))
  const highlight = parseHighlightColor(attr(child(element, 'w:highlight'), 'w:val'))
  const shd = parseShading(child(element, 'w:shd'))
  const sz = parseHalfPoint(attr(child(element, 'w:sz'), 'w:val'))
  const szCs = parseHalfPoint(attr(child(element, 'w:szCs'), 'w:val'))
  const rFonts = parseFontSet(child(element, 'w:rFonts'))
  const spacing = parseTwip(attr(child(element, 'w:spacing'), 'w:val'))
  const kern = parseHalfPoint(attr(child(element, 'w:kern'), 'w:val'))
  const position = parseHalfPoint(attr(child(element, 'w:position'), 'w:val'))
  const lang = parseLanguageSet(child(element, 'w:lang'))
  const caps = parseToggleElement(child(element, 'w:caps'))
  const smallCaps = parseToggleElement(child(element, 'w:smallCaps'))
  const vanish = parseToggleElement(child(element, 'w:vanish'))
  const webHidden = parseToggleElement(child(element, 'w:webHidden'))
  const rtl = parseToggleElement(child(element, 'w:rtl'))

  if (rStyle !== undefined) props.rStyle = rStyle
  if (bold !== undefined) props.bold = bold
  if (italic !== undefined) props.italic = italic
  if (underline !== undefined) props.underline = underline
  if (strike !== undefined) props.strike = strike
  if (dstrike !== undefined) props.dstrike = dstrike
  if (vertAlign !== undefined) props.vertAlign = vertAlign
  if (color !== undefined) props.color = color
  if (highlight !== undefined) props.highlight = highlight
  if (shd !== undefined) props.shd = shd
  if (sz !== undefined) props.sz = sz
  if (szCs !== undefined) props.szCs = szCs
  if (rFonts !== undefined) props.rFonts = rFonts
  if (spacing !== undefined) props.spacing = spacing
  if (kern !== undefined) props.kern = kern
  if (position !== undefined) props.position = position
  if (lang !== undefined) props.lang = lang
  if (caps !== undefined) props.caps = caps
  if (smallCaps !== undefined) props.smallCaps = smallCaps
  if (vanish !== undefined) props.vanish = vanish
  if (webHidden !== undefined) props.webHidden = webHidden
  if (rtl !== undefined) props.rtl = rtl

  return hasProps(props) ? props : undefined
}

function parseParaProps(element: OrderedXmlNode | undefined): ParaProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const props: Mutable<ParaProps> = {}
  const pStyle = attr(child(element, 'w:pStyle'), 'w:val')
  const numPr = parseNumPr(child(element, 'w:numPr'))
  const spacing = parseSpacing(child(element, 'w:spacing'))
  const ind = parseIndent(child(element, 'w:ind'))
  const jc = parseJustifyContent(attr(child(element, 'w:jc'), 'w:val'))
  const keepNext = parseToggleElement(child(element, 'w:keepNext'))
  const keepLines = parseToggleElement(child(element, 'w:keepLines'))
  const pageBreakBefore = parseToggleElement(child(element, 'w:pageBreakBefore'))
  const widowControl = parseToggleElement(child(element, 'w:widowControl'))
  const suppressLineNumbers = parseToggleElement(child(element, 'w:suppressLineNumbers'))
  const suppressAutoHyphens = parseToggleElement(child(element, 'w:suppressAutoHyphens'))
  const contextualSpacing = parseToggleElement(child(element, 'w:contextualSpacing'))
  const mirrorIndents = parseToggleElement(child(element, 'w:mirrorIndents'))
  const outlineLvl = parseInteger(attr(child(element, 'w:outlineLvl'), 'w:val'))
  const textAlignment = parseTextAlignment(attr(child(element, 'w:textAlignment'), 'w:val'))
  const tabs = parseTabs(child(element, 'w:tabs'))
  const pBdr = parseBorderSet(child(element, 'w:pBdr'))
  const shd = parseShading(child(element, 'w:shd'))
  const framePr = parseFrameProps(child(element, 'w:framePr'))
  const divId = parseInteger(attr(child(element, 'w:divId'), 'w:val'))
  const sectPr = parseSectionProps(child(element, 'w:sectPr'))

  if (pStyle !== undefined) props.pStyle = pStyle
  if (numPr !== undefined) props.numPr = numPr
  if (spacing !== undefined) props.spacing = spacing
  if (ind !== undefined) props.ind = ind
  if (jc !== undefined) props.jc = jc
  if (keepNext !== undefined) props.keepNext = keepNext
  if (keepLines !== undefined) props.keepLines = keepLines
  if (pageBreakBefore !== undefined) props.pageBreakBefore = pageBreakBefore
  if (widowControl !== undefined) props.widowControl = widowControl
  if (suppressLineNumbers !== undefined) props.suppressLineNumbers = suppressLineNumbers
  if (suppressAutoHyphens !== undefined) props.suppressAutoHyphens = suppressAutoHyphens
  if (contextualSpacing !== undefined) props.contextualSpacing = contextualSpacing
  if (mirrorIndents !== undefined) props.mirrorIndents = mirrorIndents
  if (outlineLvl !== undefined) props.outlineLvl = outlineLvl
  if (textAlignment !== undefined) props.textAlignment = textAlignment
  if (tabs !== undefined) props.tabs = tabs
  if (pBdr !== undefined) props.pBdr = pBdr
  if (shd !== undefined) props.shd = shd
  if (framePr !== undefined) props.framePr = framePr
  if (divId !== undefined) props.divId = divId
  if (sectPr !== undefined) props.sectPr = sectPr

  return hasProps(props) ? props : undefined
}

function parseTableProps(element: OrderedXmlNode | undefined): TableProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const props: Mutable<TableProps> = {}
  const tblStyle = attr(child(element, 'w:tblStyle'), 'w:val')
  const tblW = parseWidth(child(element, 'w:tblW'))
  const tblInd = parseWidth(child(element, 'w:tblInd'))
  const tblBorders = parseBorderSet(child(element, 'w:tblBorders'))
  const tblCellMar = parseInsetSet(child(element, 'w:tblCellMar'))
  const tblLayout = parseTableLayout(attr(child(element, 'w:tblLayout'), 'w:type'))
  const tblLook = parseTableLook(child(element, 'w:tblLook'))
  const jc = parseJustifyContent(attr(child(element, 'w:jc'), 'w:val'))
  const shd = parseShading(child(element, 'w:shd'))

  if (tblStyle !== undefined) props.tblStyle = tblStyle
  if (tblW !== undefined) props.tblW = tblW
  if (tblInd !== undefined) props.tblInd = tblInd
  if (tblBorders !== undefined) props.tblBorders = tblBorders
  if (tblCellMar !== undefined) props.tblCellMar = tblCellMar
  if (tblLayout !== undefined) props.tblLayout = tblLayout
  if (tblLook !== undefined) props.tblLook = tblLook
  if (jc !== undefined) props.jc = jc
  if (shd !== undefined) props.shd = shd

  return hasProps(props) ? props : undefined
}

function parseTableRowProps(element: OrderedXmlNode | undefined): TableRowProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const props: Mutable<TableRowProps> = {}
  const trHeight = parseTableRowHeight(child(element, 'w:trHeight'))
  const cantSplit = parseToggleElement(child(element, 'w:cantSplit'))
  const tblHeader = parseToggleElement(child(element, 'w:tblHeader'))
  const jc = parseJustifyContent(attr(child(element, 'w:jc'), 'w:val'))

  if (trHeight !== undefined) props.trHeight = trHeight
  if (cantSplit !== undefined) props.cantSplit = cantSplit
  if (tblHeader !== undefined) props.tblHeader = tblHeader
  if (jc !== undefined) props.jc = jc

  return hasProps(props) ? props : undefined
}

function parseTableCellProps(element: OrderedXmlNode | undefined): TableCellProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const props: Mutable<TableCellProps> = {}
  const tcW = parseWidth(child(element, 'w:tcW'))
  const gridSpan = parseInteger(attr(child(element, 'w:gridSpan'), 'w:val'))
  const vMerge = parseTableCellMerge(child(element, 'w:vMerge'))
  const tcBorders = parseBorderSet(child(element, 'w:tcBorders'))
  const shd = parseShading(child(element, 'w:shd'))
  const tcMar = parseInsetSet(child(element, 'w:tcMar'))
  const vAlign = parseTableCellVerticalAlign(attr(child(element, 'w:vAlign'), 'w:val'))
  const noWrap = parseToggleElement(child(element, 'w:noWrap'))
  const hideMark = parseToggleElement(child(element, 'w:hideMark'))

  if (tcW !== undefined) props.tcW = tcW
  if (gridSpan !== undefined) props.gridSpan = gridSpan
  if (vMerge !== undefined) props.vMerge = vMerge
  if (tcBorders !== undefined) props.tcBorders = tcBorders
  if (shd !== undefined) props.shd = shd
  if (tcMar !== undefined) props.tcMar = tcMar
  if (vAlign !== undefined) props.vAlign = vAlign
  if (noWrap !== undefined) props.noWrap = noWrap
  if (hideMark !== undefined) props.hideMark = hideMark

  return hasProps(props) ? props : undefined
}

function parseSectionProps(element: OrderedXmlNode | undefined): SectionProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const props: Mutable<SectionProps> = {}
  const pgSz = parsePageSize(child(element, 'w:pgSz'))
  const pgMar = parsePageMargins(child(element, 'w:pgMar'))
  const cols = parseSectionColumns(child(element, 'w:cols'))
  const pgNumType = parsePageNumberType(child(element, 'w:pgNumType'))
  const titlePg = parseToggleElement(child(element, 'w:titlePg'))
  const type = parseSectionBreakType(attr(child(element, 'w:type'), 'w:val'))
  const headerReference = children(element, 'w:headerReference')
    .map(parseHeaderReference)
    .filter(isDefined)
  const footerReference = children(element, 'w:footerReference')
    .map(parseFooterReference)
    .filter(isDefined)
  const lnNumType = parseLineNumberType(child(element, 'w:lnNumType'))
  const vAlign = parseSectionVerticalAlign(attr(child(element, 'w:vAlign'), 'w:val'))

  if (pgSz !== undefined) props.pgSz = pgSz
  if (pgMar !== undefined) props.pgMar = pgMar
  if (cols !== undefined) props.cols = cols
  if (pgNumType !== undefined) props.pgNumType = pgNumType
  if (titlePg !== undefined) props.titlePg = titlePg
  if (type !== undefined) props.type = type
  if (headerReference.length > 0) props.headerReference = headerReference
  if (footerReference.length > 0) props.footerReference = footerReference
  if (lnNumType !== undefined) props.lnNumType = lnNumType
  if (vAlign !== undefined) props.vAlign = vAlign

  return hasProps(props) ? props : undefined
}

function parseUnderline(element: OrderedXmlNode | undefined): Underline | undefined {
  if (element === undefined) {
    return undefined
  }

  const style = parseUnderlineStyle(attr(element, 'w:val')) ?? 'single'
  const color = parseColor(attr(element, 'w:color'))

  return {
    style,
    ...(color !== undefined ? { color } : {}),
  }
}

function parseFontSet(element: OrderedXmlNode | undefined): FontSet | undefined {
  if (element === undefined) {
    return undefined
  }

  const fonts: Mutable<FontSet> = {}
  const ascii = attr(element, 'w:ascii')
  const hAnsi = attr(element, 'w:hAnsi')
  const cs = attr(element, 'w:cs')
  const eastAsia = attr(element, 'w:eastAsia')
  const hint = parseFontHint(attr(element, 'w:hint'))
  const asciiTheme = parseFontThemeAttr(attr(element, 'w:asciiTheme'))
  const hAnsiTheme = parseFontThemeAttr(attr(element, 'w:hAnsiTheme'))
  const csTheme = parseFontThemeAttr(attr(element, 'w:cstheme'))
  const eastAsiaTheme = parseFontThemeAttr(attr(element, 'w:eastAsiaTheme'))

  if (ascii !== undefined) fonts.ascii = ascii
  if (hAnsi !== undefined) fonts.hAnsi = hAnsi
  if (cs !== undefined) fonts.cs = cs
  if (eastAsia !== undefined) fonts.eastAsia = eastAsia
  if (hint !== undefined) fonts.hint = hint
  if (asciiTheme !== undefined) fonts.asciiTheme = asciiTheme
  if (hAnsiTheme !== undefined) fonts.hAnsiTheme = hAnsiTheme
  if (csTheme !== undefined) fonts.csTheme = csTheme
  if (eastAsiaTheme !== undefined) fonts.eastAsiaTheme = eastAsiaTheme

  return hasProps(fonts) ? fonts : undefined
}

function parseFontThemeAttr(value: string | undefined): FontSet['asciiTheme'] | undefined {
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

function parseLanguageSet(element: OrderedXmlNode | undefined): LanguageSet | undefined {
  if (element === undefined) {
    return undefined
  }

  const lang: Mutable<LanguageSet> = {}
  const value = attr(element, 'w:val')
  const eastAsia = attr(element, 'w:eastAsia')
  const bidi = attr(element, 'w:bidi')

  if (value !== undefined) lang.value = value
  if (eastAsia !== undefined) lang.eastAsia = eastAsia
  if (bidi !== undefined) lang.bidi = bidi

  return hasProps(lang) ? lang : undefined
}

function parseSpacing(element: OrderedXmlNode | undefined): Spacing | undefined {
  if (element === undefined) {
    return undefined
  }

  const spacing: Mutable<Spacing> = {}
  const before = parseTwip(attr(element, 'w:before'))
  const after = parseTwip(attr(element, 'w:after'))
  const beforeAutospacing = parseOnOff(attr(element, 'w:beforeAutospacing'))
  const afterAutospacing = parseOnOff(attr(element, 'w:afterAutospacing'))
  const line = parseTwip(attr(element, 'w:line'))
  const lineRule = parseLineRule(attr(element, 'w:lineRule'))

  if (before !== undefined) spacing.before = before
  if (after !== undefined) spacing.after = after
  if (beforeAutospacing !== undefined) spacing.beforeAutospacing = beforeAutospacing
  if (afterAutospacing !== undefined) spacing.afterAutospacing = afterAutospacing
  if (line !== undefined) spacing.line = line
  if (lineRule !== undefined) spacing.lineRule = lineRule

  return hasProps(spacing) ? spacing : undefined
}

function parseIndent(element: OrderedXmlNode | undefined): Indent | undefined {
  if (element === undefined) {
    return undefined
  }

  const indent: Mutable<Indent> = {}
  const left = parseTwip(attr(element, 'w:left'))
  const right = parseTwip(attr(element, 'w:right'))
  const firstLine = parseTwip(attr(element, 'w:firstLine'))
  const hanging = parseTwip(attr(element, 'w:hanging'))
  const start = parseTwip(attr(element, 'w:start'))
  const end = parseTwip(attr(element, 'w:end'))

  if (left !== undefined) indent.left = left
  if (right !== undefined) indent.right = right
  if (firstLine !== undefined) indent.firstLine = firstLine
  if (hanging !== undefined) indent.hanging = hanging
  if (start !== undefined) indent.start = start
  if (end !== undefined) indent.end = end

  return hasProps(indent) ? indent : undefined
}

function parseTabs(element: OrderedXmlNode | undefined): {
  readonly items: ReadonlyArray<{
    readonly position: ReturnType<typeof twip>
    readonly alignment?:
      | 'left'
      | 'center'
      | 'right'
      | 'decimal'
      | 'bar'
      | 'clear'
      | 'start'
      | 'end'
      | 'num'
    readonly leader?:
      | 'none'
      | 'dot'
      | 'hyphen'
      | 'underscore'
      | 'heavy'
      | 'middleDot'
  }>
} | undefined {
  if (element === undefined) {
    return undefined
  }

  const items = children(element, 'w:tab')
    .map((tabElement) => {
      const position = parseTwip(attr(tabElement, 'w:pos'))
      if (position === undefined) {
        return undefined
      }

      const alignment = parseTabAlignment(attr(tabElement, 'w:val'))
      const leader = parseTabLeader(attr(tabElement, 'w:leader'))

      return {
        position,
        ...(alignment !== undefined ? { alignment } : {}),
        ...(leader !== undefined ? { leader } : {}),
      }
    })
    .filter(isDefined)

  return { items }
}

function parseBorderSet(element: OrderedXmlNode | undefined): BorderSet | undefined {
  if (element === undefined) {
    return undefined
  }

  const borderSet: Mutable<BorderSet> = {}
  const top = parseBorder(child(element, 'w:top'))
  const left = parseBorder(child(element, 'w:left'))
  const bottom = parseBorder(child(element, 'w:bottom'))
  const right = parseBorder(child(element, 'w:right'))
  const start = parseBorder(child(element, 'w:start'))
  const end = parseBorder(child(element, 'w:end'))
  const between = parseBorder(child(element, 'w:between'))
  const bar = parseBorder(child(element, 'w:bar'))
  const insideH = parseBorder(child(element, 'w:insideH'))
  const insideV = parseBorder(child(element, 'w:insideV'))

  if (top !== undefined) borderSet.top = top
  if (left !== undefined) borderSet.left = left
  if (bottom !== undefined) borderSet.bottom = bottom
  if (right !== undefined) borderSet.right = right
  if (start !== undefined) borderSet.start = start
  if (end !== undefined) borderSet.end = end
  if (between !== undefined) borderSet.between = between
  if (bar !== undefined) borderSet.bar = bar
  if (insideH !== undefined) borderSet.insideH = insideH
  if (insideV !== undefined) borderSet.insideV = insideV

  return hasProps(borderSet) ? borderSet : undefined
}

function parseBorder(element: OrderedXmlNode | undefined): Border | undefined {
  if (element === undefined) {
    return undefined
  }

  const border: Mutable<Border> = {}
  const style = parseBorderStyle(attr(element, 'w:val'))
  const color = parseColor(attr(element, 'w:color'))
  const size = parseEighthPoint(attr(element, 'w:sz'))
  const space = parseTwip(attr(element, 'w:space'))
  const shadow = parseOnOff(attr(element, 'w:shadow'))
  const frame = parseOnOff(attr(element, 'w:frame'))

  if (style !== undefined) border.style = style
  if (color !== undefined) border.color = color
  if (size !== undefined) border.size = size
  if (space !== undefined) border.space = space
  if (shadow !== undefined) border.shadow = shadow
  if (frame !== undefined) border.frame = frame

  return hasProps(border) ? border : undefined
}

function parseShading(element: OrderedXmlNode | undefined): Shading | undefined {
  if (element === undefined) {
    return undefined
  }

  const shd: Mutable<Shading> = {}
  const fill = parseColor(attr(element, 'w:fill'))
  const color = parseColor(attr(element, 'w:color'))
  const pattern = attr(element, 'w:val')

  if (fill !== undefined) shd.fill = fill
  if (color !== undefined) shd.color = color
  if (pattern !== undefined) shd.pattern = pattern

  return hasProps(shd) ? shd : undefined
}

function parseFrameProps(element: OrderedXmlNode | undefined): FrameProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const frame: Mutable<FrameProps> = {}
  const width = parseTwip(attr(element, 'w:w'))
  const height = parseTwip(attr(element, 'w:h'))
  const x = parseInteger(attr(element, 'w:x'))
  const y = parseInteger(attr(element, 'w:y'))
  const xAlign = parseFrameHorizontalAlign(attr(element, 'w:xAlign'))
  const yAlign = parseFrameVerticalAlign(attr(element, 'w:yAlign'))
  const hAnchor = parseFrameAnchor(attr(element, 'w:hAnchor'))
  const vAnchor = parseFrameAnchor(attr(element, 'w:vAnchor'))
  const wrap = parseFrameWrap(attr(element, 'w:wrap'))
  const lines = parseInteger(attr(element, 'w:lines'))
  const hSpace = parseTwip(attr(element, 'w:hSpace'))
  const vSpace = parseTwip(attr(element, 'w:vSpace'))
  const dropCap = parseFrameDropCap(attr(element, 'w:dropCap'))
  const lockAnchor = parseOnOff(attr(element, 'w:lockAnchor'))

  if (width !== undefined) frame.width = width
  if (height !== undefined) frame.height = height
  if (x !== undefined) frame.x = x
  if (y !== undefined) frame.y = y
  if (xAlign !== undefined) frame.xAlign = xAlign
  if (yAlign !== undefined) frame.yAlign = yAlign
  if (hAnchor !== undefined) frame.hAnchor = hAnchor
  if (vAnchor !== undefined) frame.vAnchor = vAnchor
  if (wrap !== undefined) frame.wrap = wrap
  if (lines !== undefined) frame.lines = lines
  if (hSpace !== undefined) frame.hSpace = hSpace
  if (vSpace !== undefined) frame.vSpace = vSpace
  if (dropCap !== undefined) frame.dropCap = dropCap
  if (lockAnchor !== undefined) frame.lockAnchor = lockAnchor

  return hasProps(frame) ? frame : undefined
}

function parseNumPr(element: OrderedXmlNode | undefined): NumPr | undefined {
  if (element === undefined) {
    return undefined
  }

  const numPr: Mutable<NumPr> = {}
  const ilvl = parseInteger(attr(child(element, 'w:ilvl'), 'w:val'))
  const numId = attr(child(element, 'w:numId'), 'w:val')

  if (ilvl !== undefined) numPr.ilvl = ilvl
  if (numId !== undefined) numPr.numId = numId

  return hasProps(numPr) ? numPr : undefined
}

function parseWidth(element: OrderedXmlNode | undefined): Width | undefined {
  if (element === undefined) {
    return undefined
  }

  const type = parseWidthType(attr(element, 'w:type'))
  if (type === undefined) {
    return undefined
  }

  const valueAttr = attr(element, 'w:w')
  const value =
    type === 'dxa'
      ? parseTwip(valueAttr)
      : type === 'pct'
        ? parsePct(valueAttr)
        : undefined

  return {
    type,
    ...(value !== undefined ? { value } : {}),
  }
}

function parseInsetSet(element: OrderedXmlNode | undefined): InsetSet | undefined {
  if (element === undefined) {
    return undefined
  }

  const insetSet: Mutable<InsetSet> = {}
  const top = parseWidth(child(element, 'w:top'))
  const left = parseWidth(child(element, 'w:left'))
  const bottom = parseWidth(child(element, 'w:bottom'))
  const right = parseWidth(child(element, 'w:right'))
  const start = parseWidth(child(element, 'w:start'))
  const end = parseWidth(child(element, 'w:end'))

  if (top !== undefined) insetSet.top = top
  if (left !== undefined) insetSet.left = left
  if (bottom !== undefined) insetSet.bottom = bottom
  if (right !== undefined) insetSet.right = right
  if (start !== undefined) insetSet.start = start
  if (end !== undefined) insetSet.end = end

  return hasProps(insetSet) ? insetSet : undefined
}

function parseTableLook(element: OrderedXmlNode | undefined): TableLook | undefined {
  if (element === undefined) {
    return undefined
  }

  const look: Mutable<TableLook> = {}
  const value = attr(element, 'w:val')
  const firstRow = parseOnOff(attr(element, 'w:firstRow'))
  const lastRow = parseOnOff(attr(element, 'w:lastRow'))
  const firstColumn = parseOnOff(attr(element, 'w:firstColumn'))
  const lastColumn = parseOnOff(attr(element, 'w:lastColumn'))
  const noHBand = parseOnOff(attr(element, 'w:noHBand'))
  const noVBand = parseOnOff(attr(element, 'w:noVBand'))

  if (value !== undefined) look.value = value
  if (firstRow !== undefined) look.firstRow = firstRow
  if (lastRow !== undefined) look.lastRow = lastRow
  if (firstColumn !== undefined) look.firstColumn = firstColumn
  if (lastColumn !== undefined) look.lastColumn = lastColumn
  if (noHBand !== undefined) look.noHBand = noHBand
  if (noVBand !== undefined) look.noVBand = noVBand

  return hasProps(look) ? look : undefined
}

function parseTableRowHeight(element: OrderedXmlNode | undefined): TableRowHeight | undefined {
  if (element === undefined) {
    return undefined
  }

  const val = parseTwip(attr(element, 'w:val'))
  if (val === undefined) {
    return undefined
  }

  const hRule = parseTableRowHeightRule(attr(element, 'w:hRule'))
  return {
    val,
    ...(hRule !== undefined ? { hRule } : {}),
  }
}

function parseTableCellMerge(element: OrderedXmlNode | undefined): TableCellMerge | undefined {
  if (element === undefined) {
    return undefined
  }

  const value = attr(element, 'w:val')
  if (value === 'restart') return 'restart'
  return 'continue'
}

function parsePageSize(element: OrderedXmlNode | undefined): {
  readonly w: ReturnType<typeof twip>
  readonly h: ReturnType<typeof twip>
  readonly orient?: 'portrait' | 'landscape'
} | undefined {
  if (element === undefined) {
    return undefined
  }

  const w = parseTwip(attr(element, 'w:w'))
  const h = parseTwip(attr(element, 'w:h'))
  const orient = parsePageOrientation(attr(element, 'w:orient'))

  if (w === undefined || h === undefined) {
    return undefined
  }

  return {
    w,
    h,
    ...(orient !== undefined ? { orient } : {}),
  }
}

function parsePageMargins(element: OrderedXmlNode | undefined): {
  readonly top?: ReturnType<typeof twip>
  readonly right?: ReturnType<typeof twip>
  readonly bottom?: ReturnType<typeof twip>
  readonly left?: ReturnType<typeof twip>
  readonly header?: ReturnType<typeof twip>
  readonly footer?: ReturnType<typeof twip>
  readonly gutter?: ReturnType<typeof twip>
} | undefined {
  if (element === undefined) {
    return undefined
  }

  const margins = {}
  const top = parseTwip(attr(element, 'w:top'))
  const right = parseTwip(attr(element, 'w:right'))
  const bottom = parseTwip(attr(element, 'w:bottom'))
  const left = parseTwip(attr(element, 'w:left'))
  const header = parseTwip(attr(element, 'w:header'))
  const footer = parseTwip(attr(element, 'w:footer'))
  const gutter = parseTwip(attr(element, 'w:gutter'))

  if (top !== undefined) {
    Object.assign(margins, { top })
  }
  if (right !== undefined) {
    Object.assign(margins, { right })
  }
  if (bottom !== undefined) {
    Object.assign(margins, { bottom })
  }
  if (left !== undefined) {
    Object.assign(margins, { left })
  }
  if (header !== undefined) {
    Object.assign(margins, { header })
  }
  if (footer !== undefined) {
    Object.assign(margins, { footer })
  }
  if (gutter !== undefined) {
    Object.assign(margins, { gutter })
  }

  return hasProps(margins) ? margins : undefined
}

function parseSectionColumns(element: OrderedXmlNode | undefined): {
  readonly num?: number
  readonly space?: ReturnType<typeof twip>
  readonly sep?: OnOff
  readonly equalWidth?: OnOff
  readonly col: ReadonlyArray<SectionColumn>
} | undefined {
  if (element === undefined) {
    return undefined
  }

  const num = parseInteger(attr(element, 'w:num'))
  const space = parseTwip(attr(element, 'w:space'))
  const sep = parseOnOff(attr(element, 'w:sep'))
  const equalWidth = parseOnOff(attr(element, 'w:equalWidth'))
  const col = children(element, 'w:col')
    .map(parseSectionColumn)
    .filter(isDefined)

  return {
    ...(num !== undefined ? { num } : {}),
    ...(space !== undefined ? { space } : {}),
    ...(sep !== undefined ? { sep } : {}),
    ...(equalWidth !== undefined ? { equalWidth } : {}),
    col,
  }
}

function parseSectionColumn(element: OrderedXmlNode): SectionColumn {
  const w = parseTwip(attr(element, 'w:w'))
  const space = parseTwip(attr(element, 'w:space'))

  return {
    ...(w !== undefined ? { w } : {}),
    ...(space !== undefined ? { space } : {}),
  }
}

function parsePageNumberType(element: OrderedXmlNode | undefined): PageNumberType | undefined {
  if (element === undefined) {
    return undefined
  }

  const pageNumberType: Mutable<PageNumberType> = {}
  const start = parseInteger(attr(element, 'w:start'))
  const fmt = attr(element, 'w:fmt')

  if (start !== undefined) pageNumberType.start = start
  if (fmt !== undefined) pageNumberType.fmt = fmt

  return hasProps(pageNumberType) ? pageNumberType : undefined
}

function parseHeaderReference(element: OrderedXmlNode): HeaderReference | undefined {
  const id = attr(element, 'r:id')
  const type = parseHeaderFooterReferenceType(attr(element, 'w:type'))

  if (id === undefined || type === undefined) {
    return undefined
  }

  return { id, type }
}

function parseFooterReference(element: OrderedXmlNode): FooterReference | undefined {
  const id = attr(element, 'r:id')
  const type = parseHeaderFooterReferenceType(attr(element, 'w:type'))

  if (id === undefined || type === undefined) {
    return undefined
  }

  return { id, type }
}

function parseLineNumberType(element: OrderedXmlNode | undefined): LineNumberType | undefined {
  if (element === undefined) {
    return undefined
  }

  const lineNumberType: Mutable<LineNumberType> = {}
  const countBy = parseInteger(attr(element, 'w:countBy'))
  const start = parseInteger(attr(element, 'w:start'))
  const distance = parseTwip(attr(element, 'w:distance'))
  const restart = parseLineNumberRestart(attr(element, 'w:restart'))

  if (countBy !== undefined) lineNumberType.countBy = countBy
  if (start !== undefined) lineNumberType.start = start
  if (distance !== undefined) lineNumberType.distance = distance
  if (restart !== undefined) lineNumberType.restart = restart

  return hasProps(lineNumberType) ? lineNumberType : undefined
}

function parseUnknownNode(element: OrderedXmlNode): UnknownNode {
  return {
    kind: 'unknown',
    xml: xmlBuilder.build([element]),
  }
}

function attr(element: OrderedXmlNode | undefined, name: string): string | undefined {
  return element?.[':@']?.[`@_${name}`]
}

function child(element: OrderedXmlNode | undefined, name: string): OrderedXmlNode | undefined {
  for (const entry of nodeChildren(element)) {
    if (nodeName(entry) === name) {
      return entry
    }
  }

  return undefined
}

function children(
  element: OrderedXmlNode | undefined,
  name: string,
): ReadonlyArray<OrderedXmlNode> {
  const matches: OrderedXmlNode[] = []

  for (const entry of nodeChildren(element)) {
    if (nodeName(entry) === name) {
      matches.push(entry)
    }
  }

  return matches
}

function findElement(
  entries: ReadonlyArray<OrderedXmlNode>,
  name: string,
): OrderedXmlNode | undefined {
  for (const entry of entries) {
    if (nodeName(entry) === name) {
      return entry
    }
  }

  return undefined
}

function findDescendant(
  element: OrderedXmlNode | undefined,
  name: string,
): OrderedXmlNode | undefined {
  if (element === undefined) {
    return undefined
  }

  for (const entry of nodeChildren(element)) {
    if (nodeName(entry) === name) {
      return entry
    }

    const nested = findDescendant(entry, name)
    if (nested !== undefined) {
      return nested
    }
  }

  return undefined
}

function nodeName(element: OrderedXmlNode): string | undefined {
  const keys = Object.keys(element)
  for (const key of keys) {
    if (key !== ':@') {
      return key
    }
  }

  return undefined
}

function nodeChildren(element: OrderedXmlNode | undefined): ReadonlyArray<OrderedXmlNode> {
  if (element === undefined) {
    return []
  }

  const name = nodeName(element)
  if (name === undefined) {
    return []
  }

  const value = element[name]
  return Array.isArray(value) ? value : []
}

function textValue(element: OrderedXmlNode): string {
  let value = ''

  for (const entry of nodeChildren(element)) {
    if (nodeName(entry) === '#text') {
      value += entry['#text'] ?? ''
    }
  }

  return value
}

function isIgnorableText(element: OrderedXmlNode): boolean {
  return nodeName(element) === '#text' && (element['#text'] ?? '').trim() === ''
}

function parseOnOff(value: string | undefined): OnOff | undefined {
  if (value === undefined) {
    return undefined
  }

  switch (value) {
    case '0':
    case 'false':
    case 'off':
      return false
    default:
      return true
  }
}

function parseToggleElement(element: OrderedXmlNode | undefined): OnOff | undefined {
  if (element === undefined) {
    return undefined
  }

  return parseOnOff(attr(element, 'w:val')) ?? true
}

function parseInteger(value: string | undefined): number | undefined {
  if (value === undefined || value === '') {
    return undefined
  }

  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

function parseTwip(value: string | undefined): ReturnType<typeof twip> | undefined {
  const parsed = parseInteger(value)
  return parsed !== undefined ? twip(parsed) : undefined
}

function parseHalfPoint(value: string | undefined): ReturnType<typeof halfPoint> | undefined {
  const parsed = parseInteger(value)
  return parsed !== undefined ? halfPoint(parsed) : undefined
}

function parseEighthPoint(value: string | undefined): ReturnType<typeof eighthPoint> | undefined {
  const parsed = parseInteger(value)
  return parsed !== undefined ? eighthPoint(parsed) : undefined
}

function parsePct(value: string | undefined): ReturnType<typeof pct> | undefined {
  const parsed = parseInteger(value)
  return parsed !== undefined ? pct(parsed) : undefined
}

function parseColor(value: string | undefined): Color | undefined {
  if (value === undefined || value === '') {
    return undefined
  }

  return value === 'auto' ? 'auto' : hexColor(value)
}

function parseBreakType(value: string | undefined): BreakType | undefined {
  switch (value) {
    case undefined:
    case 'textWrapping':
      return value === undefined ? 'line' : 'textWrapping'
    case 'page':
      return 'page'
    case 'column':
      return 'column'
    default:
      return undefined
  }
}

function parseBreakClear(value: string | undefined): BreakClear | undefined {
  switch (value) {
    case 'none':
      return 'none'
    case 'left':
      return 'left'
    case 'right':
      return 'right'
    case 'all':
      return 'all'
    default:
      return undefined
  }
}

function parseVerticalAlign(value: string | undefined): VerticalAlign | undefined {
  switch (value) {
    case 'baseline':
      return 'baseline'
    case 'superscript':
      return 'superscript'
    case 'subscript':
      return 'subscript'
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

function parseUnderlineStyle(value: string | undefined): Underline['style'] | undefined {
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

function parseLineRule(value: string | undefined): Spacing['lineRule'] | undefined {
  switch (value) {
    case 'auto':
    case 'exact':
    case 'atLeast':
      return value
    default:
      return undefined
  }
}

function parseJustifyContent(value: string | undefined): JustifyContent | undefined {
  switch (value) {
    case 'start':
    case 'left':
      return 'start'
    case 'center':
      return 'center'
    case 'end':
    case 'right':
      return 'end'
    case 'both':
    case 'justify':
      return 'both'
    case 'distribute':
      return 'distribute'
    default:
      return undefined
  }
}

function parseTextAlignment(
  value: string | undefined,
): ParaProps['textAlignment'] | undefined {
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

function parseTabAlignment(value: string | undefined):
  | 'left'
  | 'center'
  | 'right'
  | 'decimal'
  | 'bar'
  | 'clear'
  | 'start'
  | 'end'
  | 'num'
  | undefined {
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

function parseTabLeader(value: string | undefined):
  | 'none'
  | 'dot'
  | 'hyphen'
  | 'underscore'
  | 'heavy'
  | 'middleDot'
  | undefined {
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

function parseFontHint(value: string | undefined): FontSet['hint'] | undefined {
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

function parseTableLayout(value: string | undefined): TableProps['tblLayout'] | undefined {
  switch (value) {
    case 'fixed':
      return 'fixed'
    case 'autofit':
      return 'autofit'
    default:
      return undefined
  }
}

function parseWidthType(value: string | undefined): Width['type'] | undefined {
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

function parseTableCellVerticalAlign(
  value: string | undefined,
): TableCellProps['vAlign'] | undefined {
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

function parseTableRowHeightRule(
  value: string | undefined,
): TableRowHeight['hRule'] | undefined {
  switch (value) {
    case 'auto':
    case 'atLeast':
    case 'exact':
      return value
    default:
      return undefined
  }
}

function parseFrameDropCap(value: string | undefined): FrameProps['dropCap'] | undefined {
  switch (value) {
    case 'none':
    case 'drop':
    case 'margin':
      return value
    default:
      return undefined
  }
}

function parseFrameWrap(value: string | undefined): FrameProps['wrap'] | undefined {
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

function parseFrameHorizontalAlign(
  value: string | undefined,
): FrameProps['xAlign'] | undefined {
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

function parseFrameVerticalAlign(
  value: string | undefined,
): FrameProps['yAlign'] | undefined {
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

function parseFrameAnchor(value: string | undefined): FrameProps['hAnchor'] | undefined {
  switch (value) {
    case 'page':
    case 'margin':
    case 'text':
      return value
    default:
      return undefined
  }
}

function parsePageOrientation(
  value: string | undefined,
): 'portrait' | 'landscape' | undefined {
  switch (value) {
    case 'portrait':
      return 'portrait'
    case 'landscape':
      return 'landscape'
    default:
      return undefined
  }
}

function parseHeaderFooterReferenceType(
  value: string | undefined,
): HeaderReference['type'] | undefined {
  switch (value) {
    case 'default':
    case 'first':
    case 'even':
      return value
    default:
      return undefined
  }
}

function parseSectionBreakType(value: string | undefined): SectionProps['type'] | undefined {
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
): LineNumberType['restart'] | undefined {
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

function isDefined<T>(value: T | undefined): value is T {
  return value !== undefined
}

function hasProps(value: object): boolean {
  return Object.keys(value).length > 0
}
