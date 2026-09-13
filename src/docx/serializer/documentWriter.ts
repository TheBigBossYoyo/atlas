import { XMLBuilder } from 'fast-xml-parser'

import {
  assertNever,
  type Block,
  type Bookmark,
  type Border,
  type BorderSet,
  type BreakNode,
  type CommentRange,
  type CommentReference,
  type Document,
  type Drawing,
  type DrawingAnchorChild,
  type DrawingExtent,
  type EndnoteReference,
  type FooterReference,
  type FontSet,
  type FootnoteReference,
  type FrameProps,
  type HeaderReference,
  type Hyperlink,
  type HyperlinkChild,
  type InsRevision,
  type DelRevision,
  type InsetSet,
  type LineNumberType,
  type NumPr,
  type OnOff,
  type PageNumberType,
  type ParaProps,
  type Paragraph,
  type ParagraphChild,
  type Run,
  type RunChild,
  type RunProps,
  type Section,
  type SectionColumn,
  type SectionColumns,
  type SectionProps,
  type Shading,
  type Spacing,
  type Table,
  type TableCell,
  type TableCellProps,
  type TableChild,
  type TableLook,
  type TableProps,
  type TableRow,
  type TableRowChild,
  type TableRowHeight,
  type TableRowProps,
  type TextNode,
  type UnknownNode,
  type Width,
} from '../model'

interface XmlAttributes {
  readonly [name: string]: string | undefined
}

interface OrderedXmlNode {
  readonly ':@'?: XmlAttributes
  readonly '#text'?: string
  readonly [name: string]: OrderedXmlNode[] | XmlAttributes | string | undefined
}

interface SerializeState {
  readonly unknownXml: Map<string, string>
  nextUnknownId: number
}

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const UNKNOWN_PLACEHOLDER_TAG = 'atlas-raw-unknown'

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  suppressEmptyNode: false,
  format: false,
  processEntities: true,
  suppressBooleanAttributes: false,
  preserveOrder: true,
})

export function writeDocumentXml(doc: Document): string {
  const state = createSerializeState()
  const documentNode = createElement('w:document', [buildBodyNodeWithState(doc, state)], buildDocumentAttributes())
  const xml = `${XML_DECLARATION}${xmlBuilder.build([documentNode])}`
  return restoreUnknownXml(collapseEmptyElements(xml), state)
}

export function buildBodyNode(doc: Document): unknown {
  return buildBodyNodeWithState(doc, createSerializeState())
}

export function buildParagraph(paragraph: Paragraph): unknown {
  return buildParagraphWithState(paragraph, createSerializeState())
}

export function buildRun(run: Run): unknown {
  return buildRunWithState(run, createSerializeState())
}

export function buildTable(table: Table): unknown {
  return buildTableWithState(table, createSerializeState())
}

export function buildTableRow(row: TableRow): unknown {
  return buildTableRowWithState(row, createSerializeState())
}

export function buildTableCell(cell: TableCell): unknown {
  return buildTableCellWithState(cell, createSerializeState())
}

export function buildSectionProperties(section: Section | SectionProps): unknown {
  return buildSectionPropertiesNode('kind' in section ? section.props : section)
}

export function buildRunProperties(runProps: RunProps | undefined): unknown {
  return buildRunPropertiesNode(runProps)
}

export function buildParagraphProperties(paraProps: ParaProps | undefined): unknown {
  return buildParagraphPropertiesNode(paraProps)
}

function buildBodyNodeWithState(doc: Document, state: SerializeState): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const sections = doc.sections.length > 0 ? doc.sections : [emptySection()]

  for (let index = 0; index < sections.length; index += 1) {
    const section = sections[index]
    if (section === undefined) {
      continue
    }

    const isLast = index === sections.length - 1
    const sectionChildren = isLast
      ? buildFinalSectionChildren(section, state)
      : buildIntermediateSectionChildren(section, state)

    children.push(...sectionChildren)
  }

  return createElement('w:body', children)
}

function buildIntermediateSectionChildren(section: Section, state: SerializeState): OrderedXmlNode[] {
  const children: OrderedXmlNode[] = []
  const lastBlock = section.blocks[section.blocks.length - 1]

  if (lastBlock?.kind === 'paragraph') {
    for (let index = 0; index < section.blocks.length - 1; index += 1) {
      const block = section.blocks[index]
      if (block !== undefined) {
        children.push(buildBlockNode(block, state))
      }
    }

    children.push(buildParagraphWithState(withSectionProps(lastBlock, section.props), state))
    return children
  }

  for (const block of section.blocks) {
    children.push(buildBlockNode(block, state))
  }

  children.push(buildParagraphWithState(createSectionBoundaryParagraph(section.props), state))
  return children
}

function buildFinalSectionChildren(section: Section, state: SerializeState): OrderedXmlNode[] {
  const children: OrderedXmlNode[] = []
  const lastBlock = section.blocks[section.blocks.length - 1]
  const hasTrailingSectionParagraph =
    lastBlock?.kind === 'paragraph' && lastBlock.props?.sectPr !== undefined

  if (hasTrailingSectionParagraph && lastBlock.kind === 'paragraph') {
    for (let index = 0; index < section.blocks.length - 1; index += 1) {
      const block = section.blocks[index]
      if (block !== undefined) {
        children.push(buildBlockNode(block, state))
      }
    }

    children.push(buildParagraphWithState(withSectionProps(lastBlock, section.props), state))
    return children
  }

  for (const block of section.blocks) {
    children.push(buildBlockNode(block, state))
  }

  children.push(buildSectionPropertiesNode(section.props))
  return children
}

function buildBlockNode(block: Block, state: SerializeState): OrderedXmlNode {
  switch (block.kind) {
    case 'paragraph':
      return buildParagraphWithState(block, state)
    case 'table':
      return buildTableWithState(block, state)
    case 'unknown':
      return buildUnknownPlaceholder(block, state)
    default:
      return assertNever(block)
  }
}

function buildParagraphWithState(paragraph: Paragraph, state: SerializeState): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const props = buildParagraphPropertiesNode(paragraph.props)

  if (props !== undefined) {
    children.push(props)
  }

  for (const child of paragraph.children) {
    children.push(buildParagraphChildNode(child, state))
  }

  return createElement('w:p', children)
}

function buildParagraphChildNode(child: ParagraphChild, state: SerializeState): OrderedXmlNode {
  switch (child.kind) {
    case 'run':
      return buildRunWithState(child, state)
    case 'hyperlink':
      return buildHyperlinkNode(child, state)
    case 'bookmark':
      return buildBookmarkNode(child)
    case 'comment-range':
      return buildCommentRangeNode(child)
    case 'comment-reference':
      return buildCommentReferenceNode(child)
    case 'footnote-reference':
      return buildFootnoteReferenceNode(child)
    case 'endnote-reference':
      return buildEndnoteReferenceNode(child)
    case 'ins-revision':
      return buildRevisionNode(child, state)
    case 'del-revision':
      return buildRevisionNode(child, state)
    case 'unknown':
      return buildUnknownPlaceholder(child, state)
    default:
      return assertNever(child)
  }
}

function buildRunWithState(run: Run, state: SerializeState, asDel = false): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const props = buildRunPropertiesNode(run.props)

  if (props !== undefined) {
    children.push(props)
  }

  for (const child of run.children) {
    children.push(buildRunChildNode(child, state, asDel))
  }

  return createElement('w:r', children)
}

function buildRunChildNode(child: RunChild, state: SerializeState, asDel = false): OrderedXmlNode {
  switch (child.kind) {
    case 'text':
      return buildTextNode(child, asDel)
    case 'tab':
      return createElement('w:tab')
    case 'break':
      return buildBreakNode(child)
    case 'drawing':
      return buildDrawingNode(child, state)
    case 'comment-reference':
      return buildCommentReferenceNode(child)
    case 'footnote-reference':
      return buildFootnoteReferenceNode(child)
    case 'endnote-reference':
      return buildEndnoteReferenceNode(child)
    case 'unknown':
      return buildUnknownPlaceholder(child, state)
    default:
      return assertNever(child)
  }
}

function buildHyperlinkNode(
  hyperlink: Hyperlink,
  state: SerializeState,
  asDel = false,
): OrderedXmlNode {
  const attributes = createAttributes()
  appendAttribute(attributes, '@_r:id', hyperlink.relationshipId)
  appendAttribute(attributes, '@_w:anchor', hyperlink.anchor)
  appendAttribute(attributes, '@_w:tooltip', hyperlink.tooltip)
  appendAttribute(attributes, '@_w:tgtFrame', hyperlink.targetFrame)
  appendAttribute(attributes, '@_w:history', buildOnOffAttribute(hyperlink.history))

  const children: OrderedXmlNode[] = []
  for (const child of hyperlink.children) {
    children.push(buildHyperlinkChildNode(child, state, asDel))
  }

  return createElement('w:hyperlink', children, hasAttributes(attributes) ? attributes : undefined)
}

function buildHyperlinkChildNode(
  child: HyperlinkChild,
  state: SerializeState,
  asDel = false,
): OrderedXmlNode {
  switch (child.kind) {
    case 'run':
      return buildRunWithState(child, state, asDel)
    case 'bookmark':
      return buildBookmarkNode(child)
    case 'comment-range':
      return buildCommentRangeNode(child)
    case 'comment-reference':
      return buildCommentReferenceNode(child)
    case 'footnote-reference':
      return buildFootnoteReferenceNode(child)
    case 'endnote-reference':
      return buildEndnoteReferenceNode(child)
    case 'unknown':
      return buildUnknownPlaceholder(child, state)
    default:
      return assertNever(child)
  }
}

function buildTextNode(textNode: TextNode, asDel = false): OrderedXmlNode {
  return createElement(asDel ? 'w:delText' : 'w:t', [createText(textNode.value)], {
    '@_xml:space': 'preserve',
  })
}

function buildRevisionNode(
  revision: InsRevision | DelRevision,
  state: SerializeState,
): OrderedXmlNode {
  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:id', revision.id)
  appendAttribute(attributes, '@_w:author', revision.author)
  appendAttribute(attributes, '@_w:date', revision.date)

  const isDel = revision.kind === 'del-revision'
  const children = getRevisionChildren(revision).map((child) =>
    buildRevisionParagraphChildNode(child, state, isDel),
  )

  return createElement(
    isDel ? 'w:del' : 'w:ins',
    children,
    hasAttributes(attributes) ? attributes : undefined,
  )
}

function buildRevisionParagraphChildNode(
  child: ParagraphChild,
  state: SerializeState,
  asDel: boolean,
): OrderedXmlNode {
  switch (child.kind) {
    case 'run':
      return buildRunWithState(child, state, asDel)
    case 'hyperlink':
      return buildHyperlinkNode(child, state, asDel)
    case 'bookmark':
      return buildBookmarkNode(child)
    case 'comment-range':
      return buildCommentRangeNode(child)
    case 'comment-reference':
      return buildCommentReferenceNode(child)
    case 'footnote-reference':
      return buildFootnoteReferenceNode(child)
    case 'endnote-reference':
      return buildEndnoteReferenceNode(child)
    case 'ins-revision':
      return buildRevisionNode(child, state)
    case 'del-revision':
      return buildRevisionNode(child, state)
    case 'unknown':
      return buildUnknownPlaceholder(child, state)
    default:
      return assertNever(child)
  }
}

function getRevisionChildren(
  revision: InsRevision | DelRevision,
): ReadonlyArray<ParagraphChild> {
  return revision.children as ReadonlyArray<ParagraphChild>
}

function buildBreakNode(breakNode: BreakNode): OrderedXmlNode {
  const attributes = createAttributes()

  if (breakNode.breakType !== undefined && breakNode.breakType !== 'line') {
    attributes['@_w:type'] = breakNode.breakType
  }

  appendAttribute(attributes, '@_w:clear', breakNode.clear)
  return createElement('w:br', [], hasAttributes(attributes) ? attributes : undefined)
}

function buildDrawingNode(drawing: Drawing, state: SerializeState): OrderedXmlNode {
  const layoutName = drawing.layout === 'anchor' ? 'wp:anchor' : 'wp:inline'
  const layoutAttributes = drawing.layout === 'anchor' ? buildAnchorAttributes() : undefined
  const layoutChildren =
    drawing.layout === 'anchor' && drawing.anchorChildren !== undefined
      ? buildAnchorChildrenNodes(drawing, drawing.anchorChildren, state)
      : buildInlineDrawingChildren(drawing)

  return createElement('w:drawing', [createElement(layoutName, layoutChildren, layoutAttributes)])
}

function buildInlineDrawingChildren(drawing: Drawing): OrderedXmlNode[] {
  const children: OrderedXmlNode[] = []

  if (drawing.extent !== undefined) {
    children.push(buildExtentNode(drawing.extent))
  }

  children.push(buildDocPrNode(drawing))

  if (drawing.relationshipId !== undefined) {
    children.push(buildGraphicNode(drawing.relationshipId))
  }

  return children
}

/**
 * Rebuilds `wp:anchor`'s children from the original captured order (DXS-03):
 * the extent/docPr/graphic slots come from current model state (so an
 * in-app edit to e.g. alt text is reflected), while every other captured
 * child — position, wrap choice, effectExtent, cNvGraphicFramePr — is
 * unmodeled and replayed verbatim via the same raw-XML placeholder
 * mechanism `buildUnknownPlaceholder` uses elsewhere.
 */
function buildAnchorChildrenNodes(
  drawing: Drawing,
  anchorChildren: ReadonlyArray<DrawingAnchorChild>,
  state: SerializeState,
): OrderedXmlNode[] {
  const children: OrderedXmlNode[] = []

  for (const entry of anchorChildren) {
    if (entry.kind === 'unknown') {
      children.push(buildUnknownPlaceholder(entry, state))
      continue
    }

    switch (entry.slot) {
      case 'extent':
        if (drawing.extent !== undefined) {
          children.push(buildExtentNode(drawing.extent))
        }
        break
      case 'docPr':
        children.push(buildDocPrNode(drawing))
        break
      case 'graphic':
        if (drawing.relationshipId !== undefined) {
          children.push(buildGraphicNode(drawing.relationshipId))
        }
        break
      default:
        assertNever(entry.slot)
    }
  }

  return children
}

function buildExtentNode(extent: DrawingExtent): OrderedXmlNode {
  return createElement('wp:extent', [], {
    '@_cx': String(extent.cx),
    '@_cy': String(extent.cy),
  })
}

function buildDocPrNode(drawing: Drawing): OrderedXmlNode {
  const attributes = createAttributes()
  attributes['@_id'] = '1'
  appendAttribute(attributes, '@_name', drawing.name)
  appendAttribute(attributes, '@_title', drawing.title)
  appendAttribute(attributes, '@_descr', drawing.description)
  return createElement('wp:docPr', [], attributes)
}

function buildGraphicNode(relationshipId: string): OrderedXmlNode {
  return createElement('a:graphic', [
    createElement('a:graphicData', [
      createElement('pic:pic', [
        createElement('pic:blipFill', [
          createElement('a:blip', [], {
            '@_r:embed': relationshipId,
          }),
        ]),
      ]),
    ], {
      '@_uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }),
  ])
}

function buildAnchorAttributes(): XmlAttributes {
  const attributes = createAttributes()
  attributes['@_distT'] = '0'
  attributes['@_distB'] = '0'
  attributes['@_distL'] = '0'
  attributes['@_distR'] = '0'
  attributes['@_simplePos'] = '0'
  attributes['@_relativeHeight'] = '0'
  attributes['@_behindDoc'] = '0'
  attributes['@_locked'] = '0'
  attributes['@_layoutInCell'] = '1'
  attributes['@_allowOverlap'] = '1'
  return attributes
}

function buildBookmarkNode(bookmark: Bookmark): OrderedXmlNode {
  const attributes = createAttributes()
  attributes['@_w:id'] = bookmark.id
  appendAttribute(attributes, '@_w:name', bookmark.name)
  appendAttribute(attributes, '@_w:colFirst', stringifyNumber(bookmark.colFirst))
  appendAttribute(attributes, '@_w:colLast', stringifyNumber(bookmark.colLast))

  return createElement(
    bookmark.boundary === 'start' ? 'w:bookmarkStart' : 'w:bookmarkEnd',
    [],
    attributes,
  )
}

function buildCommentRangeNode(commentRange: CommentRange): OrderedXmlNode {
  return createElement(
    commentRange.boundary === 'start' ? 'w:commentRangeStart' : 'w:commentRangeEnd',
    [],
    {
      '@_w:id': commentRange.id,
    },
  )
}

function buildCommentReferenceNode(commentReference: CommentReference): OrderedXmlNode {
  return createElement('w:commentReference', [], {
    '@_w:id': commentReference.id,
  })
}

function buildFootnoteReferenceNode(footnoteReference: FootnoteReference): OrderedXmlNode {
  const attributes = createAttributes()
  attributes['@_w:id'] = footnoteReference.id
  appendAttribute(
    attributes,
    '@_w:customMarkFollows',
    buildOnOffAttribute(footnoteReference.customMarkFollows),
  )
  return createElement('w:footnoteReference', [], attributes)
}

function buildEndnoteReferenceNode(endnoteReference: EndnoteReference): OrderedXmlNode {
  const attributes = createAttributes()
  attributes['@_w:id'] = endnoteReference.id
  appendAttribute(
    attributes,
    '@_w:customMarkFollows',
    buildOnOffAttribute(endnoteReference.customMarkFollows),
  )
  return createElement('w:endnoteReference', [], attributes)
}

function buildTableWithState(table: Table, state: SerializeState): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const props = buildTablePropertiesNode(table.props)

  if (props !== undefined) {
    children.push(props)
  }

  pushIfDefined(children, buildTableGridNode(table.tblGrid))

  for (const row of table.rows) {
    children.push(buildTableChildNode(row, state))
  }

  return createElement('w:tbl', children)
}

function buildTableGridNode(tblGrid: Table['tblGrid']): OrderedXmlNode | undefined {
  if (tblGrid === undefined || tblGrid.length === 0) {
    return undefined
  }

  const columns = tblGrid.map((width) => createElement('w:gridCol', [], { '@_w:w': String(width) }))
  return createElement('w:tblGrid', columns)
}

function buildTableChildNode(child: TableChild, state: SerializeState): OrderedXmlNode {
  switch (child.kind) {
    case 'table-row':
      return buildTableRowWithState(child, state)
    case 'unknown':
      return buildUnknownPlaceholder(child, state)
    default:
      return assertNever(child)
  }
}

function buildTableRowWithState(row: TableRow, state: SerializeState): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const props = buildTableRowPropertiesNode(row.props)

  if (props !== undefined) {
    children.push(props)
  }

  for (const cell of row.cells) {
    children.push(buildTableRowChildNode(cell, state))
  }

  return createElement('w:tr', children)
}

function buildTableRowChildNode(child: TableRowChild, state: SerializeState): OrderedXmlNode {
  switch (child.kind) {
    case 'table-cell':
      return buildTableCellWithState(child, state)
    case 'unknown':
      return buildUnknownPlaceholder(child, state)
    default:
      return assertNever(child)
  }
}

function buildTableCellWithState(cell: TableCell, state: SerializeState): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const props = buildTableCellPropertiesNode(cell.props)

  if (props !== undefined) {
    children.push(props)
  }

  for (const block of cell.blocks) {
    children.push(buildBlockNode(block, state))
  }

  return createElement('w:tc', children)
}

function buildTablePropertiesNode(tableProps: TableProps | undefined): OrderedXmlNode | undefined {
  if (tableProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildValueElement('w:tblStyle', tableProps.tblStyle))
  pushIfDefined(children, buildWidthElement('w:tblW', tableProps.tblW))
  pushIfDefined(children, buildWidthElement('w:tblInd', tableProps.tblInd))
  pushIfDefined(children, buildBorderSetElement('w:tblBorders', tableProps.tblBorders))
  pushIfDefined(children, buildInsetSetElement('w:tblCellMar', tableProps.tblCellMar))
  pushIfDefined(children, buildTypeElement('w:tblLayout', tableProps.tblLayout))
  pushIfDefined(children, buildTableLookElement(tableProps.tblLook))
  pushIfDefined(children, buildValueElement('w:jc', tableProps.jc))
  pushIfDefined(children, buildShadingElement('w:shd', tableProps.shd))

  return children.length > 0 ? createElement('w:tblPr', children) : undefined
}

function buildTableRowPropertiesNode(rowProps: TableRowProps | undefined): OrderedXmlNode | undefined {
  if (rowProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildTableRowHeightElement(rowProps.trHeight))
  pushIfDefined(children, buildToggleElement('w:cantSplit', rowProps.cantSplit))
  pushIfDefined(children, buildToggleElement('w:tblHeader', rowProps.tblHeader))
  pushIfDefined(children, buildValueElement('w:jc', rowProps.jc))

  return children.length > 0 ? createElement('w:trPr', children) : undefined
}

function buildTableCellPropertiesNode(cellProps: TableCellProps | undefined): OrderedXmlNode | undefined {
  if (cellProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildWidthElement('w:tcW', cellProps.tcW))
  pushIfDefined(children, buildValueElement('w:gridSpan', cellProps.gridSpan))
  pushIfDefined(children, buildTableCellMergeElement(cellProps.vMerge))
  pushIfDefined(children, buildBorderSetElement('w:tcBorders', cellProps.tcBorders))
  pushIfDefined(children, buildShadingElement('w:shd', cellProps.shd))
  pushIfDefined(children, buildInsetSetElement('w:tcMar', cellProps.tcMar))
  pushIfDefined(children, buildValueElement('w:vAlign', cellProps.vAlign))
  pushIfDefined(children, buildToggleElement('w:noWrap', cellProps.noWrap))
  pushIfDefined(children, buildToggleElement('w:hideMark', cellProps.hideMark))

  return children.length > 0 ? createElement('w:tcPr', children) : undefined
}

function buildTableCellMergeElement(merge: TableCellProps['vMerge']): OrderedXmlNode | undefined {
  if (merge === undefined) {
    return undefined
  }

  if (merge === 'continue') {
    return createElement('w:vMerge')
  }

  return createElement('w:vMerge', [], {
    '@_w:val': merge,
  })
}

function buildTableLookElement(look: TableLook | undefined): OrderedXmlNode | undefined {
  if (look === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:val', look.value)
  appendAttribute(attributes, '@_w:firstRow', buildOnOffAttribute(look.firstRow))
  appendAttribute(attributes, '@_w:lastRow', buildOnOffAttribute(look.lastRow))
  appendAttribute(attributes, '@_w:firstColumn', buildOnOffAttribute(look.firstColumn))
  appendAttribute(attributes, '@_w:lastColumn', buildOnOffAttribute(look.lastColumn))
  appendAttribute(attributes, '@_w:noHBand', buildOnOffAttribute(look.noHBand))
  appendAttribute(attributes, '@_w:noVBand', buildOnOffAttribute(look.noVBand))

  return hasAttributes(attributes) ? createElement('w:tblLook', [], attributes) : undefined
}

function buildTableRowHeightElement(height: TableRowHeight | undefined): OrderedXmlNode | undefined {
  if (height === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  attributes['@_w:val'] = String(height.val)
  appendAttribute(attributes, '@_w:hRule', height.hRule)
  return createElement('w:trHeight', [], attributes)
}

function buildRunPropertiesNode(runProps: RunProps | undefined): OrderedXmlNode | undefined {
  if (runProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildValueElement('w:rStyle', runProps.rStyle))
  pushIfDefined(children, buildToggleElement('w:b', runProps.bold))
  pushIfDefined(children, buildToggleElement('w:i', runProps.italic))
  pushIfDefined(children, buildUnderlineElement(runProps.underline))
  pushIfDefined(children, buildToggleElement('w:strike', runProps.strike))
  pushIfDefined(children, buildToggleElement('w:dstrike', runProps.dstrike))
  pushIfDefined(children, buildValueElement('w:vertAlign', runProps.vertAlign))
  pushIfDefined(children, buildValueElement('w:color', runProps.color))
  pushIfDefined(children, buildValueElement('w:highlight', runProps.highlight))
  pushIfDefined(children, buildShadingElement('w:shd', runProps.shd))
  pushIfDefined(children, buildValueElement('w:sz', runProps.sz))
  pushIfDefined(children, buildValueElement('w:szCs', runProps.szCs))
  pushIfDefined(children, buildFontSetElement(runProps.rFonts))
  pushIfDefined(children, buildValueElement('w:spacing', runProps.spacing))
  pushIfDefined(children, buildValueElement('w:kern', runProps.kern))
  pushIfDefined(children, buildValueElement('w:position', runProps.position))
  pushIfDefined(children, buildLanguageSetElement(runProps.lang))
  pushIfDefined(children, buildToggleElement('w:caps', runProps.caps))
  pushIfDefined(children, buildToggleElement('w:smallCaps', runProps.smallCaps))
  pushIfDefined(children, buildToggleElement('w:vanish', runProps.vanish))
  pushIfDefined(children, buildToggleElement('w:webHidden', runProps.webHidden))
  pushIfDefined(children, buildToggleElement('w:rtl', runProps.rtl))

  return children.length > 0 ? createElement('w:rPr', children) : undefined
}

function buildParagraphPropertiesNode(paraProps: ParaProps | undefined): OrderedXmlNode | undefined {
  if (paraProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildValueElement('w:pStyle', paraProps.pStyle))
  pushIfDefined(children, buildNumPrElement(paraProps.numPr))
  pushIfDefined(children, buildSpacingElement(paraProps.spacing))
  pushIfDefined(children, buildIndentElement(paraProps.ind))
  pushIfDefined(children, buildValueElement('w:jc', paraProps.jc))
  pushIfDefined(children, buildToggleElement('w:keepNext', paraProps.keepNext))
  pushIfDefined(children, buildToggleElement('w:keepLines', paraProps.keepLines))
  pushIfDefined(children, buildToggleElement('w:pageBreakBefore', paraProps.pageBreakBefore))
  pushIfDefined(children, buildToggleElement('w:widowControl', paraProps.widowControl))
  pushIfDefined(children, buildToggleElement('w:suppressLineNumbers', paraProps.suppressLineNumbers))
  pushIfDefined(children, buildToggleElement('w:suppressAutoHyphens', paraProps.suppressAutoHyphens))
  pushIfDefined(children, buildToggleElement('w:contextualSpacing', paraProps.contextualSpacing))
  pushIfDefined(children, buildToggleElement('w:mirrorIndents', paraProps.mirrorIndents))
  pushIfDefined(children, buildValueElement('w:outlineLvl', paraProps.outlineLvl))
  pushIfDefined(children, buildValueElement('w:textAlignment', paraProps.textAlignment))
  pushIfDefined(children, buildTabsElement(paraProps.tabs))
  pushIfDefined(children, buildBorderSetElement('w:pBdr', paraProps.pBdr))
  pushIfDefined(children, buildShadingElement('w:shd', paraProps.shd))
  pushIfDefined(children, buildFramePropsElement(paraProps.framePr))
  pushIfDefined(children, buildValueElement('w:divId', paraProps.divId))

  if (paraProps.sectPr !== undefined) {
    children.push(buildSectionPropertiesNode(paraProps.sectPr))
  }

  return children.length > 0 ? createElement('w:pPr', children) : undefined
}

function buildSectionPropertiesNode(sectionProps: SectionProps): OrderedXmlNode {
  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildPageSizeElement(sectionProps.pgSz))
  pushIfDefined(children, buildPageMarginsElement(sectionProps.pgMar))
  pushIfDefined(children, buildSectionColumnsElement(sectionProps.cols))
  pushIfDefined(children, buildPageNumberTypeElement(sectionProps.pgNumType))
  pushIfDefined(children, buildToggleElement('w:titlePg', sectionProps.titlePg))
  pushIfDefined(children, buildValueElement('w:type', sectionProps.type))

  for (const headerReference of sectionProps.headerReference ?? []) {
    children.push(buildHeaderReferenceNode(headerReference))
  }

  for (const footerReference of sectionProps.footerReference ?? []) {
    children.push(buildFooterReferenceNode(footerReference))
  }

  pushIfDefined(children, buildLineNumberTypeElement(sectionProps.lnNumType))
  pushIfDefined(children, buildValueElement('w:vAlign', sectionProps.vAlign))

  return createElement('w:sectPr', children)
}

function buildUnderlineElement(underline: RunProps['underline']): OrderedXmlNode | undefined {
  if (underline === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  attributes['@_w:val'] = underline.style
  appendAttribute(attributes, '@_w:color', underline.color)
  return createElement('w:u', [], attributes)
}

function buildFontSetElement(fonts: FontSet | undefined): OrderedXmlNode | undefined {
  if (fonts === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:ascii', fonts.ascii)
  appendAttribute(attributes, '@_w:hAnsi', fonts.hAnsi)
  appendAttribute(attributes, '@_w:cs', fonts.cs)
  appendAttribute(attributes, '@_w:eastAsia', fonts.eastAsia)
  appendAttribute(attributes, '@_w:hint', fonts.hint)

  return hasAttributes(attributes) ? createElement('w:rFonts', [], attributes) : undefined
}

function buildLanguageSetElement(language: RunProps['lang']): OrderedXmlNode | undefined {
  if (language === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:val', language.value)
  appendAttribute(attributes, '@_w:eastAsia', language.eastAsia)
  appendAttribute(attributes, '@_w:bidi', language.bidi)

  return hasAttributes(attributes) ? createElement('w:lang', [], attributes) : undefined
}

function buildSpacingElement(spacing: Spacing | undefined): OrderedXmlNode | undefined {
  if (spacing === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:before', stringifyNumber(spacing.before))
  appendAttribute(attributes, '@_w:after', stringifyNumber(spacing.after))
  appendAttribute(
    attributes,
    '@_w:beforeAutospacing',
    buildOnOffAttribute(spacing.beforeAutospacing),
  )
  appendAttribute(
    attributes,
    '@_w:afterAutospacing',
    buildOnOffAttribute(spacing.afterAutospacing),
  )
  appendAttribute(attributes, '@_w:line', stringifyNumber(spacing.line))
  appendAttribute(attributes, '@_w:lineRule', spacing.lineRule)

  return hasAttributes(attributes) ? createElement('w:spacing', [], attributes) : undefined
}

function buildIndentElement(indent: ParaProps['ind']): OrderedXmlNode | undefined {
  if (indent === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:left', stringifyNumber(indent.left))
  appendAttribute(attributes, '@_w:right', stringifyNumber(indent.right))
  appendAttribute(attributes, '@_w:firstLine', stringifyNumber(indent.firstLine))
  appendAttribute(attributes, '@_w:hanging', stringifyNumber(indent.hanging))
  appendAttribute(attributes, '@_w:start', stringifyNumber(indent.start))
  appendAttribute(attributes, '@_w:end', stringifyNumber(indent.end))

  return hasAttributes(attributes) ? createElement('w:ind', [], attributes) : undefined
}

function buildFramePropsElement(frameProps: FrameProps | undefined): OrderedXmlNode | undefined {
  if (frameProps === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:w', stringifyNumber(frameProps.width))
  appendAttribute(attributes, '@_w:h', stringifyNumber(frameProps.height))
  appendAttribute(attributes, '@_w:x', stringifyNumber(frameProps.x))
  appendAttribute(attributes, '@_w:y', stringifyNumber(frameProps.y))
  appendAttribute(attributes, '@_w:xAlign', frameProps.xAlign)
  appendAttribute(attributes, '@_w:yAlign', frameProps.yAlign)
  appendAttribute(attributes, '@_w:hAnchor', frameProps.hAnchor)
  appendAttribute(attributes, '@_w:vAnchor', frameProps.vAnchor)
  appendAttribute(attributes, '@_w:wrap', frameProps.wrap)
  appendAttribute(attributes, '@_w:lines', stringifyNumber(frameProps.lines))
  appendAttribute(attributes, '@_w:hSpace', stringifyNumber(frameProps.hSpace))
  appendAttribute(attributes, '@_w:vSpace', stringifyNumber(frameProps.vSpace))
  appendAttribute(attributes, '@_w:dropCap', frameProps.dropCap)
  appendAttribute(attributes, '@_w:lockAnchor', buildOnOffAttribute(frameProps.lockAnchor))

  return hasAttributes(attributes) ? createElement('w:framePr', [], attributes) : undefined
}

function buildNumPrElement(numPr: NumPr | undefined): OrderedXmlNode | undefined {
  if (numPr === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []
  pushIfDefined(children, buildValueElement('w:ilvl', numPr.ilvl))
  pushIfDefined(children, buildValueElement('w:numId', numPr.numId))
  return children.length > 0 ? createElement('w:numPr', children) : undefined
}

function buildTabsElement(tabs: ParaProps['tabs']): OrderedXmlNode | undefined {
  if (tabs === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []
  for (const tab of tabs.items) {
    const attributes = createAttributes()
    appendAttribute(attributes, '@_w:val', tab.alignment)
    appendAttribute(attributes, '@_w:leader', tab.leader)
    attributes['@_w:pos'] = String(tab.position)
    children.push(createElement('w:tab', [], attributes))
  }

  return createElement('w:tabs', children)
}

function buildBorderSetElement(name: string, borderSet: BorderSet | undefined): OrderedXmlNode | undefined {
  if (borderSet === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildBorderElement('w:top', borderSet.top))
  pushIfDefined(children, buildBorderElement('w:left', borderSet.left))
  pushIfDefined(children, buildBorderElement('w:bottom', borderSet.bottom))
  pushIfDefined(children, buildBorderElement('w:right', borderSet.right))
  pushIfDefined(children, buildBorderElement('w:start', borderSet.start))
  pushIfDefined(children, buildBorderElement('w:end', borderSet.end))
  pushIfDefined(children, buildBorderElement('w:between', borderSet.between))
  pushIfDefined(children, buildBorderElement('w:bar', borderSet.bar))
  pushIfDefined(children, buildBorderElement('w:insideH', borderSet.insideH))
  pushIfDefined(children, buildBorderElement('w:insideV', borderSet.insideV))

  return createElement(name, children)
}

function buildBorderElement(name: string, border: Border | undefined): OrderedXmlNode | undefined {
  if (border === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:val', border.style)
  appendAttribute(attributes, '@_w:sz', stringifyNumber(border.size))
  appendAttribute(attributes, '@_w:space', stringifyNumber(border.space))
  appendAttribute(attributes, '@_w:color', border.color)
  appendAttribute(attributes, '@_w:shadow', buildOnOffAttribute(border.shadow))
  appendAttribute(attributes, '@_w:frame', buildOnOffAttribute(border.frame))

  return hasAttributes(attributes) ? createElement(name, [], attributes) : undefined
}

function buildShadingElement(name: string, shading: Shading | undefined): OrderedXmlNode | undefined {
  if (shading === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:val', shading.pattern)
  appendAttribute(attributes, '@_w:color', shading.color)
  appendAttribute(attributes, '@_w:fill', shading.fill)

  return hasAttributes(attributes) ? createElement(name, [], attributes) : undefined
}

function buildWidthElement(name: string, width: Width | undefined): OrderedXmlNode | undefined {
  if (width === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:w', stringifyNumber(width.value))
  attributes['@_w:type'] = width.type
  return createElement(name, [], attributes)
}

function buildInsetSetElement(name: string, insetSet: InsetSet | undefined): OrderedXmlNode | undefined {
  if (insetSet === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []
  pushIfDefined(children, buildWidthElement('w:top', insetSet.top))
  pushIfDefined(children, buildWidthElement('w:left', insetSet.left))
  pushIfDefined(children, buildWidthElement('w:bottom', insetSet.bottom))
  pushIfDefined(children, buildWidthElement('w:right', insetSet.right))
  pushIfDefined(children, buildWidthElement('w:start', insetSet.start))
  pushIfDefined(children, buildWidthElement('w:end', insetSet.end))
  return createElement(name, children)
}

function buildPageSizeElement(pageSize: SectionProps['pgSz']): OrderedXmlNode | undefined {
  if (pageSize === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  attributes['@_w:w'] = String(pageSize.w)
  attributes['@_w:h'] = String(pageSize.h)
  appendAttribute(attributes, '@_w:orient', pageSize.orient)
  return createElement('w:pgSz', [], attributes)
}

function buildPageMarginsElement(pageMargins: SectionProps['pgMar']): OrderedXmlNode | undefined {
  if (pageMargins === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:top', stringifyNumber(pageMargins.top))
  appendAttribute(attributes, '@_w:right', stringifyNumber(pageMargins.right))
  appendAttribute(attributes, '@_w:bottom', stringifyNumber(pageMargins.bottom))
  appendAttribute(attributes, '@_w:left', stringifyNumber(pageMargins.left))
  appendAttribute(attributes, '@_w:header', stringifyNumber(pageMargins.header))
  appendAttribute(attributes, '@_w:footer', stringifyNumber(pageMargins.footer))
  appendAttribute(attributes, '@_w:gutter', stringifyNumber(pageMargins.gutter))

  return hasAttributes(attributes) ? createElement('w:pgMar', [], attributes) : undefined
}

function buildSectionColumnsElement(columns: SectionColumns | undefined): OrderedXmlNode | undefined {
  if (columns === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:num', stringifyNumber(columns.num))
  appendAttribute(attributes, '@_w:space', stringifyNumber(columns.space))
  appendAttribute(attributes, '@_w:sep', buildOnOffAttribute(columns.sep))
  appendAttribute(attributes, '@_w:equalWidth', buildOnOffAttribute(columns.equalWidth))

  const children: OrderedXmlNode[] = []
  for (const column of columns.col) {
    children.push(buildSectionColumnNode(column))
  }

  return createElement('w:cols', children, hasAttributes(attributes) ? attributes : undefined)
}

function buildSectionColumnNode(column: SectionColumn): OrderedXmlNode {
  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:w', stringifyNumber(column.w))
  appendAttribute(attributes, '@_w:space', stringifyNumber(column.space))
  return createElement('w:col', [], hasAttributes(attributes) ? attributes : undefined)
}

function buildPageNumberTypeElement(pageNumberType: PageNumberType | undefined): OrderedXmlNode | undefined {
  if (pageNumberType === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:start', stringifyNumber(pageNumberType.start))
  appendAttribute(attributes, '@_w:fmt', pageNumberType.fmt)
  return hasAttributes(attributes) ? createElement('w:pgNumType', [], attributes) : undefined
}

function buildHeaderReferenceNode(headerReference: HeaderReference): OrderedXmlNode {
  const attributes = createAttributes()
  attributes['@_w:type'] = headerReference.type
  attributes['@_r:id'] = headerReference.id
  return createElement('w:headerReference', [], attributes)
}

function buildFooterReferenceNode(footerReference: FooterReference): OrderedXmlNode {
  const attributes = createAttributes()
  attributes['@_w:type'] = footerReference.type
  attributes['@_r:id'] = footerReference.id
  return createElement('w:footerReference', [], attributes)
}

function buildLineNumberTypeElement(lineNumberType: LineNumberType | undefined): OrderedXmlNode | undefined {
  if (lineNumberType === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:countBy', stringifyNumber(lineNumberType.countBy))
  appendAttribute(attributes, '@_w:start', stringifyNumber(lineNumberType.start))
  appendAttribute(attributes, '@_w:distance', stringifyNumber(lineNumberType.distance))
  appendAttribute(attributes, '@_w:restart', lineNumberType.restart)
  return hasAttributes(attributes) ? createElement('w:lnNumType', [], attributes) : undefined
}

function buildToggleElement(name: string, value: OnOff | undefined): OrderedXmlNode | undefined {
  if (value === undefined) {
    return undefined
  }

  if (value) {
    return createElement(name)
  }

  return createElement(name, [], {
    '@_w:val': '0',
  })
}

function buildValueElement(name: string, value: string | number | undefined): OrderedXmlNode | undefined {
  if (value === undefined) {
    return undefined
  }

  return createElement(name, [], {
    '@_w:val': String(value),
  })
}

function buildTypeElement(name: string, value: string | undefined): OrderedXmlNode | undefined {
  if (value === undefined) {
    return undefined
  }

  return createElement(name, [], {
    '@_w:type': value,
  })
}

function buildUnknownPlaceholder(node: UnknownNode, state: SerializeState): OrderedXmlNode {
  const id = `unknown-${state.nextUnknownId}`
  state.nextUnknownId += 1
  state.unknownXml.set(id, node.xml)
  return createElement(UNKNOWN_PLACEHOLDER_TAG, [], {
    '@_data-id': id,
  })
}

function buildDocumentAttributes(): XmlAttributes {
  const attributes = createAttributes()
  attributes['@_xmlns:wpc'] = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
  attributes['@_xmlns:mc'] = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
  attributes['@_xmlns:o'] = 'urn:schemas-microsoft-com:office:office'
  attributes['@_xmlns:r'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  attributes['@_xmlns:m'] = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
  attributes['@_xmlns:v'] = 'urn:schemas-microsoft-com:vml'
  attributes['@_xmlns:wp14'] = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
  attributes['@_xmlns:wp'] = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
  attributes['@_xmlns:w10'] = 'urn:schemas-microsoft-com:office:word'
  attributes['@_xmlns:w'] = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  attributes['@_xmlns:w14'] = 'http://schemas.microsoft.com/office/word/2010/wordml'
  attributes['@_xmlns:w15'] = 'http://schemas.microsoft.com/office/word/2012/wordml'
  attributes['@_xmlns:w16cex'] = 'http://schemas.microsoft.com/office/word/2018/wordml/cex'
  attributes['@_xmlns:w16cid'] = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'
  attributes['@_xmlns:w16'] = 'http://schemas.microsoft.com/office/word/2018/wordml'
  attributes['@_xmlns:w16sdtdh'] = 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash'
  attributes['@_xmlns:w16se'] = 'http://schemas.microsoft.com/office/word/2015/wordml/symex'
  attributes['@_xmlns:wpg'] = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
  attributes['@_xmlns:wpi'] = 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
  attributes['@_xmlns:wne'] = 'http://schemas.microsoft.com/office/word/2006/wordml'
  attributes['@_xmlns:wps'] = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
  attributes['@_xmlns:a'] = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  attributes['@_xmlns:pic'] = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  attributes['@_mc:Ignorable'] = 'w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14'
  return attributes
}

function createSerializeState(): SerializeState {
  return {
    unknownXml: new Map(),
    nextUnknownId: 0,
  }
}

function emptySection(): Section {
  return {
    kind: 'section',
    props: {},
    blocks: [],
  }
}

function createSectionBoundaryParagraph(sectionProps: SectionProps): Paragraph {
  return {
    kind: 'paragraph',
    props: {
      sectPr: sectionProps,
    },
    children: [],
  }
}

function withSectionProps(paragraph: Paragraph, sectionProps: SectionProps): Paragraph {
  return {
    ...paragraph,
    props: {
      ...(paragraph.props ?? {}),
      sectPr: sectionProps,
    },
  }
}

function createElement(name: string, children: ReadonlyArray<OrderedXmlNode> = [], attributes?: XmlAttributes): OrderedXmlNode {
  const element: Record<string, OrderedXmlNode[] | XmlAttributes | string | undefined> = {
    [name]: [...children],
  }

  if (attributes !== undefined && hasAttributes(attributes)) {
    element[':@'] = attributes
  }

  return element as OrderedXmlNode
}

function createText(value: string): OrderedXmlNode {
  return {
    '#text': value,
  }
}

function createAttributes(): Record<string, string | undefined> {
  return {}
}

function appendAttribute(
  attributes: Record<string, string | undefined>,
  name: string,
  value: string | undefined,
): void {
  if (value !== undefined) {
    attributes[name] = value
  }
}

function pushIfDefined(children: OrderedXmlNode[], child: OrderedXmlNode | undefined): void {
  if (child !== undefined) {
    children.push(child)
  }
}

function hasAttributes(attributes: XmlAttributes): boolean {
  return Object.keys(attributes).length > 0
}

function stringifyNumber(value: number | undefined): string | undefined {
  return value !== undefined ? String(value) : undefined
}

function buildOnOffAttribute(value: OnOff | undefined): string | undefined {
  if (value === undefined) {
    return undefined
  }

  return value ? '1' : '0'
}

function collapseEmptyElements(xml: string): string {
  return xml.replace(/<([A-Za-z_][\w.:-]*)([^>]*)><\/\1>/g, '<$1$2/>')
}

function restoreUnknownXml(xml: string, state: SerializeState): string {
  let restored = xml

  for (const [id, rawXml] of state.unknownXml.entries()) {
    const openClose = `<${UNKNOWN_PLACEHOLDER_TAG} data-id="${id}"></${UNKNOWN_PLACEHOLDER_TAG}>`
    const selfClosing = `<${UNKNOWN_PLACEHOLDER_TAG} data-id="${id}"/>`
    restored = restored.split(openClose).join(rawXml)
    restored = restored.split(selfClosing).join(rawXml)
  }

  return restored
}
