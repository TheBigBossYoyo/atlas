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
  type DrawingCrop,
  type DrawingEffectExtent,
  type DrawingExtent,
  type DrawingPositionH,
  type DrawingPositionV,
  type DrawingWrap,
  type DocGrid,
  type EndnoteReference,
  type Field,
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
  type WrapperPassthrough,
} from '../model'

interface XmlAttributes {
  readonly [name: string]: string | undefined
}

interface OrderedXmlNode {
  readonly ':@'?: XmlAttributes
  readonly '#text'?: string
  readonly [name: string]: OrderedXmlNode[] | XmlAttributes | string | undefined
}

/**
 * Exported (DXS-08 follow-up fix) so `partWriterSupport.ts` can share one
 * state across an entire standalone part (header/footer/footnotes/comments)
 * and run the same `restoreUnknownXml` substitution pass those parts'
 * content needs whenever it contains a node type this serializer doesn't
 * model (e.g. `w:proofErr`, ubiquitous in real Word-authored documents) —
 * see `buildUnknownPlaceholder`'s doc comment for why a placeholder alone,
 * left un-restored, is a literal `<atlas-raw-unknown>` tag leaking into the
 * saved XML.
 */
export interface SerializeState {
  readonly unknownXml: Map<string, string>
  nextUnknownId: number
  /**
   * `w:sdt`/`mc:AlternateContent` wrapper regions captured while parsing
   * the source document (round-trip fidelity audit, DXS round 2 follow-up)
   * — see `WrapperPassthrough`'s doc comment on `../model/document.ts`.
   * Empty for any part other than `word/document.xml` (headers/footers/
   * comments/footnotes each create their own state with none), which is
   * exactly "never matches, always falls back to normal serialization" —
   * the pre-existing behavior for those parts.
   */
  readonly wrapperRegions: ReadonlyArray<WrapperPassthrough>
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
  const state = createSerializeState(doc.wrappers)
  const documentNode = createElement('w:document', [buildBodyNodeWithState(doc, state)], buildDocumentAttributes(doc))
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
  return buildRunPropertiesNode(runProps, createSerializeState())
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
  const lastBlock = section.blocks[section.blocks.length - 1]

  if (lastBlock?.kind === 'paragraph') {
    const children = buildBlockNodes(section.blocks.slice(0, -1), state)
    children.push(buildParagraphWithState(withSectionProps(lastBlock, section.props), state))
    return children
  }

  const children = buildBlockNodes(section.blocks, state)
  children.push(buildParagraphWithState(createSectionBoundaryParagraph(section.props), state))
  return children
}

function buildFinalSectionChildren(section: Section, state: SerializeState): OrderedXmlNode[] {
  const lastBlock = section.blocks[section.blocks.length - 1]
  const hasTrailingSectionParagraph =
    lastBlock?.kind === 'paragraph' && lastBlock.props?.sectPr !== undefined

  if (hasTrailingSectionParagraph && lastBlock.kind === 'paragraph') {
    const children = buildBlockNodes(section.blocks.slice(0, -1), state)
    children.push(buildParagraphWithState(withSectionProps(lastBlock, section.props), state))
    return children
  }

  const children = buildBlockNodes(section.blocks, state)
  children.push(buildSectionPropertiesNode(section.props))
  return children
}

/**
 * Builds a sibling run of blocks, splicing in an unedited `w:sdt`/
 * `mc:AlternateContent` wrapper's exact source bytes in place of the
 * block(s) it covers wherever `state.wrapperRegions` proves none of them
 * were touched since parse (round-trip fidelity audit, DXS round 2
 * follow-up) — see `findWrapperRegionAt`/`WrapperPassthrough`'s doc
 * comments. Falls back to the normal per-block build for everything else,
 * identical to the plain `blocks.map((b) => buildBlockNode(b, state))`
 * this replaces.
 */
function buildBlockNodes(blocks: ReadonlyArray<Block>, state: SerializeState): OrderedXmlNode[] {
  const nodes: OrderedXmlNode[] = []
  let index = 0
  while (index < blocks.length) {
    const region = findWrapperRegionAt(state.wrapperRegions, blocks, index)
    if (region !== undefined) {
      nodes.push(buildRawPassthroughPlaceholder(region.raw, state))
      index += region.content.length
      continue
    }

    const block = blocks[index]
    if (block !== undefined) {
      nodes.push(buildBlockNode(block, state))
    }
    index += 1
  }
  return nodes
}

/**
 * Finds a recorded wrapper-passthrough region whose captured content
 * starts exactly at `items[index]` (round-trip fidelity audit, DXS round 2
 * follow-up) — see `WrapperPassthrough`'s doc comment on
 * `../model/document.ts`. Reference equality only: every one of the
 * region's captured objects must still be the SAME object, in the same
 * position, as when it was captured at parse time — proof that nothing
 * inside the wrapper (and nothing about the wrapper's position among its
 * siblings) has changed since. `items`/`index` are generic over the three
 * sibling-list shapes a wrapper can sit in (a section's `Block[]`, a
 * paragraph's `ParagraphChild[]`, a run's `RunChild[]`).
 */
function findWrapperRegionAt<T>(
  regions: ReadonlyArray<WrapperPassthrough>,
  items: ReadonlyArray<T>,
  index: number,
): WrapperPassthrough | undefined {
  for (const region of regions) {
    const content = region.content as ReadonlyArray<unknown>
    const length = content.length
    if (length === 0 || index + length > items.length) {
      continue
    }

    let matches = true
    for (let offset = 0; offset < length; offset += 1) {
      if ((items[index + offset] as unknown) !== content[offset]) {
        matches = false
        break
      }
    }
    if (matches) {
      return region
    }
  }
  return undefined
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

export function buildParagraphWithState(paragraph: Paragraph, state: SerializeState): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const props = buildParagraphPropertiesNode(paragraph.props)

  if (props !== undefined) {
    children.push(props)
  }

  // Round-trip fidelity audit, DXS round 2 follow-up: same wrapper-region
  // passthrough `buildBlockNodes` does for a section's blocks — see
  // `findWrapperRegionAt`'s doc comment.
  const paragraphChildren = paragraph.children
  let index = 0
  while (index < paragraphChildren.length) {
    const region = findWrapperRegionAt(state.wrapperRegions, paragraphChildren, index)
    if (region !== undefined) {
      children.push(buildRawPassthroughPlaceholder(region.raw, state))
      index += region.content.length
      continue
    }

    const child = paragraphChildren[index]
    if (child !== undefined) {
      children.push(...buildParagraphChildNodes(child, state))
    }
    index += 1
  }

  return createElement('w:p', children, buildParagraphAttributes(paragraph))
}

/**
 * DXS-10: re-emit `w14:paraId`/`w14:textId`/`w:rsid*` verbatim when the
 * source paragraph carried them, instead of silently dropping them on every
 * save. Atlas never generates these for a paragraph that never had them —
 * only Word itself assigns fresh values, and `paraId` in particular is the
 * join key `commentsExtended.xml` uses to correlate a comment's
 * resolved/done state (D16), so inventing one here could point a comment at
 * the wrong paragraph.
 */
function buildParagraphAttributes(paragraph: Paragraph): XmlAttributes | undefined {
  const attributes = createAttributes()
  appendAttribute(attributes, '@_w14:paraId', paragraph.paraId)
  appendAttribute(attributes, '@_w14:textId', paragraph.textId)
  appendAttribute(attributes, '@_w:rsidR', paragraph.rsidR)
  appendAttribute(attributes, '@_w:rsidRDefault', paragraph.rsidRDefault)
  appendAttribute(attributes, '@_w:rsidP', paragraph.rsidP)
  appendAttribute(attributes, '@_w:rsidRPr', paragraph.rsidRPr)
  return hasAttributes(attributes) ? attributes : undefined
}

/**
 * A `Field` is the one `ParagraphChild` that can expand to zero, one, or
 * many sibling nodes (a complex field regenerated by "Update field(s)"
 * needs 5+ actual `<w:r>` siblings — begin/instrText/separate/result(s)/end
 * — see `buildFieldNodes`) — impossible to express through this function's
 * one-child-in-one-node-out shape, so every call site dispatches a `Field`
 * through `buildFieldNodes` (via this wrapper) before ever reaching the
 * single-node switch below.
 */
function buildParagraphChildNodes(
  child: ParagraphChild,
  state: SerializeState,
): ReadonlyArray<OrderedXmlNode> {
  return child.kind === 'field' ? buildFieldNodes(child, state) : [buildParagraphChildNode(child, state)]
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
    case 'field':
      // Every real call site goes through `buildParagraphChildNodes` (see
      // its doc comment) — reaching this branch would mean a new call site
      // forgot to.
      throw new Error('buildParagraphChildNode: Field must be built via buildFieldNodes, not this function')
    default:
      return assertNever(child)
  }
}

function buildRunWithState(run: Run, state: SerializeState, asDel = false): OrderedXmlNode {
  const children: OrderedXmlNode[] = []
  const props = buildRunPropertiesNode(run.props, state)

  if (props !== undefined) {
    children.push(props)
  }

  // Round-trip fidelity audit, DXS round 2 follow-up: same wrapper-region
  // passthrough as `buildParagraphWithState`'s — this is the level real
  // Word documents put `mc:AlternateContent` at (a shape/text box's
  // modern-vs-legacy-VML pair sits directly inside its "containing" run).
  const runChildren = run.children
  let index = 0
  while (index < runChildren.length) {
    const region = findWrapperRegionAt(state.wrapperRegions, runChildren, index)
    if (region !== undefined) {
      children.push(buildRawPassthroughPlaceholder(region.raw, state))
      index += region.content.length
      continue
    }

    const child = runChildren[index]
    if (child !== undefined) {
      children.push(buildRunChildNode(child, state, asDel))
    }
    index += 1
  }

  return createElement('w:r', children, buildRunAttributes(run))
}

/** DXS-10: re-emit a run's `w:rsid*` bookkeeping attributes verbatim. */
function buildRunAttributes(run: Run): XmlAttributes | undefined {
  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:rsidR', run.rsidR)
  appendAttribute(attributes, '@_w:rsidRPr', run.rsidRPr)
  appendAttribute(attributes, '@_w:rsidDel', run.rsidDel)
  return hasAttributes(attributes) ? attributes : undefined
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
    children.push(...buildHyperlinkChildNodes(child, state, asDel))
  }

  return createElement('w:hyperlink', children, hasAttributes(attributes) ? attributes : undefined)
}

/** See `buildParagraphChildNodes`'s doc comment — same reason this wraps rather than being one function. */
function buildHyperlinkChildNodes(
  child: HyperlinkChild,
  state: SerializeState,
  asDel: boolean,
): ReadonlyArray<OrderedXmlNode> {
  return child.kind === 'field'
    ? buildFieldNodes(child, state)
    : [buildHyperlinkChildNode(child, state, asDel)]
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
    case 'field':
      throw new Error('buildHyperlinkChildNode: Field must be built via buildFieldNodes, not this function')
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
  const children = getRevisionChildren(revision).flatMap((child) =>
    buildRevisionParagraphChildNodes(child, state, isDel),
  )

  return createElement(
    isDel ? 'w:del' : 'w:ins',
    children,
    hasAttributes(attributes) ? attributes : undefined,
  )
}

/** See `buildParagraphChildNodes`'s doc comment — same reason this wraps rather than being one function. */
function buildRevisionParagraphChildNodes(
  child: ParagraphChild,
  state: SerializeState,
  asDel: boolean,
): ReadonlyArray<OrderedXmlNode> {
  return child.kind === 'field'
    ? buildFieldNodes(child, state)
    : [buildRevisionParagraphChildNode(child, state, asDel)]
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
    case 'field':
      throw new Error('buildRevisionParagraphChildNode: Field must be built via buildFieldNodes, not this function')
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
  const layoutAttributes = drawing.layout === 'anchor' ? buildAnchorAttributes(drawing) : undefined
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

  if (drawing.effectExtent !== undefined) {
    children.push(buildEffectExtentNode(drawing.effectExtent))
  }

  children.push(buildDocPrNode(drawing))
  // Round-trip fidelity audit (DXS round 2): `wp:cNvGraphicFramePr`'s
  // `a:graphicFrameLocks noChangeAspect="1"` (the same aspect-ratio-lock
  // hint every picture-inserting tool sets) was previously dropped for
  // every `wp:inline` drawing on every save — `wp:anchor`'s equivalent
  // survives via `anchorChildren`'s raw passthrough (see
  // `buildAnchorChildrenNodes`'s doc comment), but `wp:inline` has no such
  // passthrough slot, so this is emitted as the same universal boilerplate
  // `buildGraphicNode`'s `pic:cNvPicPr/a:picLocks` now is. No local
  // `xmlns:a` needed — `w:document`'s root already declares it (see the
  // baseline namespace map this module builds the root element from).
  children.push(
    createElement('wp:cNvGraphicFramePr', [createElement('a:graphicFrameLocks', [], { '@_noChangeAspect': '1' })]),
  )

  if (drawing.relationshipId !== undefined) {
    children.push(buildGraphicNode(drawing))
  }

  return children
}

/**
 * Rebuilds `wp:anchor`'s children from the original captured order (DXS-03):
 * the extent/effectExtent/positionH/positionV/wrap/docPr/graphic slots come
 * from current model state (so an in-app edit — alt text, or a future
 * reposition/rewrap command — is reflected instead of replaying stale
 * captured XML; DXP-09 moved position/wrap from raw capture to these typed
 * slots), and every other captured child — `wp:simplePos`,
 * `wp:cNvGraphicFramePr` — is unmodeled and replayed verbatim via the same
 * raw-XML placeholder mechanism `buildUnknownPlaceholder` uses elsewhere.
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
      case 'effectExtent':
        if (drawing.effectExtent !== undefined) {
          children.push(buildEffectExtentNode(drawing.effectExtent))
        }
        break
      case 'positionH':
        if (drawing.positionH !== undefined) {
          children.push(buildPositionHNode(drawing.positionH))
        }
        break
      case 'positionV':
        if (drawing.positionV !== undefined) {
          children.push(buildPositionVNode(drawing.positionV))
        }
        break
      case 'wrap':
        children.push(buildWrapNode(drawing.wrap))
        break
      case 'docPr':
        children.push(buildDocPrNode(drawing))
        break
      case 'graphic':
        if (drawing.relationshipId !== undefined) {
          children.push(buildGraphicNode(drawing))
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

function buildEffectExtentNode(effectExtent: DrawingEffectExtent): OrderedXmlNode {
  return createElement('wp:effectExtent', [], {
    '@_l': String(effectExtent.l),
    '@_t': String(effectExtent.t),
    '@_r': String(effectExtent.r),
    '@_b': String(effectExtent.b),
  })
}

function buildPositionHNode(positionH: DrawingPositionH): OrderedXmlNode {
  return createElement(
    'wp:positionH',
    [buildPositionChild(positionH.align, positionH.offsetEmu)],
    { '@_relativeFrom': positionH.relativeFrom },
  )
}

function buildPositionVNode(positionV: DrawingPositionV): OrderedXmlNode {
  return createElement(
    'wp:positionV',
    [buildPositionChild(positionV.align, positionV.offsetEmu)],
    { '@_relativeFrom': positionV.relativeFrom },
  )
}

/**
 * `wp:positionH`/`wp:positionV` each carry exactly one child: `wp:align`
 * (text content, e.g. "center") or `wp:posOffset` (text content, EMU). A
 * schema-valid document always has one; if a hand-built `Drawing` somehow
 * has neither, default to a zero offset rather than emitting an
 * (invalid) childless position element.
 */
function buildPositionChild(align: string | undefined, offsetEmu: number | undefined): OrderedXmlNode {
  if (align !== undefined) {
    return createElement('wp:align', [createText(align)])
  }
  return createElement('wp:posOffset', [createText(String(offsetEmu ?? 0))])
}

/** Maps a `DrawingWrap.mode` back to its `wp:wrap*` element name. */
const WRAP_MODE_ELEMENT_NAMES: Readonly<Record<DrawingWrap['mode'], string>> = {
  none: 'wp:wrapNone',
  square: 'wp:wrapSquare',
  tight: 'wp:wrapTight',
  through: 'wp:wrapThrough',
  topAndBottom: 'wp:wrapTopAndBottom',
}

/**
 * Rebuilds the anchor's required wrap-choice element. `wrap` is normally
 * always present (a schema-valid `wp:anchor` requires one) — the `undefined`
 * case only guards a hand-built `Drawing`, falling back to `wrapNone` (no
 * wrap, matching the safest "just show it, don't blank out text" default).
 */
function buildWrapNode(wrap: DrawingWrap | undefined): OrderedXmlNode {
  const mode = wrap?.mode ?? 'none'
  const elementName = WRAP_MODE_ELEMENT_NAMES[mode]
  const attributes = createAttributes()
  appendAttribute(attributes, '@_wrapText', wrap?.side)
  appendAttribute(attributes, '@_distT', stringifyNumber(wrap?.distTEmu))
  appendAttribute(attributes, '@_distB', stringifyNumber(wrap?.distBEmu))
  appendAttribute(attributes, '@_distL', stringifyNumber(wrap?.distLEmu))
  appendAttribute(attributes, '@_distR', stringifyNumber(wrap?.distREmu))
  return createElement(elementName, [], attributes)
}

function buildDocPrNode(drawing: Drawing): OrderedXmlNode {
  const attributes = createAttributes()
  attributes['@_id'] = '1'
  appendAttribute(attributes, '@_name', drawing.name)
  appendAttribute(attributes, '@_title', drawing.title)
  appendAttribute(attributes, '@_descr', drawing.description)
  return createElement('wp:docPr', [], attributes)
}

function buildGraphicNode(drawing: Drawing): OrderedXmlNode {
  if (drawing.relationshipId === undefined) {
    throw new Error('buildGraphicNode requires a relationshipId')
  }

  const blipFillChildren: OrderedXmlNode[] = [
    createElement('a:blip', [], { '@_r:embed': drawing.relationshipId }),
  ]
  if (drawing.crop !== undefined) {
    blipFillChildren.push(buildSrcRectNode(drawing.crop))
  }
  // Round-trip fidelity audit (DXS round 2): every real `pic:blipFill`
  // (Word's own output, and `docx`-generated fixtures alike) fills the
  // frame by stretching the source rect — `<a:stretch><a:fillRect/></a:
  // stretch>` — but this was previously omitted unconditionally, on every
  // picture, every save.
  blipFillChildren.push(createElement('a:stretch', [createElement('a:fillRect', [])]))

  const picChildren: OrderedXmlNode[] = [
    // `pic:nvPicPr` (non-visual picture properties) carries no rendering-
    // relevant data of its own — the real accessibility name/description
    // live on `wp:docPr` (see `buildDocPrNode`, built from `drawing.name`/
    // `title`/`description`) — but `pic:cNvPicPr/a:picLocks` is the
    // "lock aspect ratio" editing hint every picture-inserting tool sets.
    // Previously the whole element was dropped on every save.
    createElement('pic:nvPicPr', [
      createElement('pic:cNvPr', [], { '@_id': '0', '@_name': '', '@_descr': '' }),
      createElement('pic:cNvPicPr', [
        createElement('a:picLocks', [], { '@_noChangeAspect': '1', '@_noChangeArrowheads': '1' }),
      ]),
    ]),
    createElement('pic:blipFill', blipFillChildren),
    buildPictureShapePropertiesNode(drawing),
  ]

  return createElement('a:graphic', [
    createElement('a:graphicData', [createElement('pic:pic', picChildren)], {
      '@_uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }),
  ])
}

/**
 * `pic:spPr` — previously emitted only when a rotation/flip transform was
 * present (and even then, missing the shape geometry every picture needs).
 * A plain rectangular picture's `a:xfrm` always carries the frame's own
 * off/ext (redundant with, but not identical in kind to, the wrapper
 * `wp:extent`/`wp:anchor` position — Word regenerates it either way, but
 * omitting it entirely was previously flagged as a difference from every
 * real-world `.docx`), and `a:prstGeom prst="rect"` is universal for an
 * unshaped picture.
 */
function buildPictureShapePropertiesNode(drawing: Drawing): OrderedXmlNode {
  const xfrmAttributes = createAttributes()
  appendAttribute(xfrmAttributes, '@_rot', stringifyNumber(drawing.transform?.rotation))
  appendAttribute(xfrmAttributes, '@_flipH', buildOnOffAttribute(drawing.transform?.flipH))
  appendAttribute(xfrmAttributes, '@_flipV', buildOnOffAttribute(drawing.transform?.flipV))

  const xfrmChildren: OrderedXmlNode[] = [
    createElement('a:off', [], { '@_x': '0', '@_y': '0' }),
  ]
  if (drawing.extent !== undefined) {
    xfrmChildren.push(
      createElement('a:ext', [], { '@_cx': String(drawing.extent.cx), '@_cy': String(drawing.extent.cy) }),
    )
  }

  return createElement('pic:spPr', [
    createElement('a:xfrm', xfrmChildren, xfrmAttributes),
    createElement('a:prstGeom', [createElement('a:avLst', [])], { '@_prst': 'rect' }),
  ])
}

function buildSrcRectNode(crop: DrawingCrop): OrderedXmlNode {
  const attributes = createAttributes()
  appendAttribute(attributes, '@_l', stringifyNumber(crop.l))
  appendAttribute(attributes, '@_t', stringifyNumber(crop.t))
  appendAttribute(attributes, '@_r', stringifyNumber(crop.r))
  appendAttribute(attributes, '@_b', stringifyNumber(crop.b))
  return createElement('a:srcRect', [], attributes)
}

function buildAnchorAttributes(drawing: Drawing): XmlAttributes {
  const attributes = createAttributes()
  attributes['@_distT'] = '0'
  attributes['@_distB'] = '0'
  attributes['@_distL'] = '0'
  attributes['@_distR'] = '0'
  attributes['@_simplePos'] = '0'
  attributes['@_relativeHeight'] = '0'
  attributes['@_behindDoc'] = buildOnOffAttribute(drawing.behindDoc) ?? '0'
  attributes['@_locked'] = '0'
  attributes['@_layoutInCell'] = '1'
  attributes['@_allowOverlap'] = buildOnOffAttribute(drawing.allowOverlap) ?? '1'
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

export function buildTableWithState(table: Table, state: SerializeState): OrderedXmlNode {
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

// Child order follows ECMA-376 `CT_TblPrBase` (§17.4.60), restricted to the
// members `TableProps` models. An element-order review found `jc` and `shd`
// emitted last instead of in-sequence, and `tblCellMar`/`tblLayout` swapped.
function buildTablePropertiesNode(tableProps: TableProps | undefined): OrderedXmlNode | undefined {
  if (tableProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildValueElement('w:tblStyle', tableProps.tblStyle))
  pushIfDefined(children, buildWidthElement('w:tblW', tableProps.tblW))
  pushIfDefined(children, buildValueElement('w:jc', tableProps.jc))
  pushIfDefined(children, buildWidthElement('w:tblInd', tableProps.tblInd))
  pushIfDefined(children, buildBorderSetElement('w:tblBorders', tableProps.tblBorders))
  pushIfDefined(children, buildShadingElement('w:shd', tableProps.shd))
  pushIfDefined(children, buildTypeElement('w:tblLayout', tableProps.tblLayout))
  pushIfDefined(children, buildInsetSetElement('w:tblCellMar', tableProps.tblCellMar))
  pushIfDefined(children, buildTableLookElement(tableProps.tblLook))

  return children.length > 0 ? createElement('w:tblPr', children) : undefined
}

// Child order follows ECMA-376 `CT_TrPrBase` (§17.4.83): `cantSplit`
// precedes `trHeight` — an element-order review found them swapped.
function buildTableRowPropertiesNode(rowProps: TableRowProps | undefined): OrderedXmlNode | undefined {
  if (rowProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildToggleElement('w:cantSplit', rowProps.cantSplit))
  pushIfDefined(children, buildTableRowHeightElement(rowProps.trHeight))
  pushIfDefined(children, buildToggleElement('w:tblHeader', rowProps.tblHeader))
  pushIfDefined(children, buildValueElement('w:jc', rowProps.jc))

  return children.length > 0 ? createElement('w:trPr', children) : undefined
}

// Child order follows ECMA-376 `CT_TcPrBase` (§17.4.70): `noWrap` precedes
// `tcMar`/`vAlign` — an element-order review found `noWrap` emitted last
// instead of right after `shd`.
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
  pushIfDefined(children, buildToggleElement('w:noWrap', cellProps.noWrap))
  pushIfDefined(children, buildInsetSetElement('w:tcMar', cellProps.tcMar))
  pushIfDefined(children, buildValueElement('w:vAlign', cellProps.vAlign))
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

// Child order follows ECMA-376 `CT_RPr`/`EG_RPrBase` (§17.3.2.28), restricted
// to the members `RunProps` models. An element-order review found this
// emitting something close to insertion order rather than schema order —
// e.g. `rFonts` last instead of second, `u`/`vertAlign` far too early, `caps`/
// `smallCaps`/`vanish`/`webHidden` far too late, and `lang` before `rtl`
// instead of after.
//
// DOCX-12 — `outline`/`emboss`/`imprint` (right after `dstrike`), `w`/
// `charScale` (right after `spacing`), `bdr` (right after `u`, before
// `shd`), and `em` (right after `rtl`, before `lang`) fill in gaps this
// same sequence had never modeled; `runProps.rPrUnknown` (any further
// sequence member — or genuinely unrecognized element — this model still
// doesn't carry a field for) is spliced into the finished list by
// `insertRPrUnknownChildren`, positioned relative to whichever of the
// above it sat next to in the source, rather than appended at the end.
function buildRunPropertiesNode(runProps: RunProps | undefined, state: SerializeState): OrderedXmlNode | undefined {
  if (runProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildValueElement('w:rStyle', runProps.rStyle))
  pushIfDefined(children, buildFontSetElement(runProps.rFonts))
  pushIfDefined(children, buildToggleElement('w:b', runProps.bold))
  pushIfDefined(children, buildToggleElement('w:bCs', runProps.boldCs))
  pushIfDefined(children, buildToggleElement('w:i', runProps.italic))
  pushIfDefined(children, buildToggleElement('w:iCs', runProps.italicCs))
  pushIfDefined(children, buildToggleElement('w:caps', runProps.caps))
  pushIfDefined(children, buildToggleElement('w:smallCaps', runProps.smallCaps))
  pushIfDefined(children, buildToggleElement('w:strike', runProps.strike))
  pushIfDefined(children, buildToggleElement('w:dstrike', runProps.dstrike))
  pushIfDefined(children, buildToggleElement('w:outline', runProps.outline))
  pushIfDefined(children, buildToggleElement('w:emboss', runProps.emboss))
  pushIfDefined(children, buildToggleElement('w:imprint', runProps.imprint))
  pushIfDefined(children, buildToggleElement('w:vanish', runProps.vanish))
  pushIfDefined(children, buildToggleElement('w:webHidden', runProps.webHidden))
  pushIfDefined(children, buildValueElement('w:color', runProps.color))
  pushIfDefined(children, buildValueElement('w:spacing', runProps.spacing))
  pushIfDefined(children, buildValueElement('w:w', runProps.charScale))
  pushIfDefined(children, buildValueElement('w:kern', runProps.kern))
  pushIfDefined(children, buildValueElement('w:position', runProps.position))
  pushIfDefined(children, buildValueElement('w:sz', runProps.sz))
  pushIfDefined(children, buildValueElement('w:szCs', runProps.szCs))
  pushIfDefined(children, buildValueElement('w:highlight', runProps.highlight))
  pushIfDefined(children, buildUnderlineElement(runProps.underline))
  pushIfDefined(children, buildBorderElement('w:bdr', runProps.bdr))
  pushIfDefined(children, buildShadingElement('w:shd', runProps.shd))
  pushIfDefined(children, buildValueElement('w:vertAlign', runProps.vertAlign))
  pushIfDefined(children, buildToggleElement('w:rtl', runProps.rtl))
  pushIfDefined(children, buildValueElement('w:em', runProps.em))
  pushIfDefined(children, buildLanguageSetElement(runProps.lang))

  insertRPrUnknownChildren(children, runProps.rPrUnknown, state)

  return children.length > 0 ? createElement('w:rPr', children) : undefined
}

/**
 * DOCX-12 — splices each captured `RPrUnknownChild` (see its doc comment on
 * `model/styles.ts`) into `children` — the `w:rPr` children
 * `buildRunPropertiesNode` has already built in schema order — immediately
 * ahead of the modeled sibling named by `before`, so the rebuilt sequence
 * stays schema-ordered instead of dumping every unmodeled child at the end.
 * Falls back to appending when `before` is `undefined` (this child was last
 * in the source) or when that sibling isn't actually present in `children`
 * (it was parsed but the value driving it has since gone missing — not
 * reachable via this file's own read-then-write round trip, but a graceful
 * fallback rather than a thrown/lost node either way).
 */
function insertRPrUnknownChildren(
  children: OrderedXmlNode[],
  unknownChildren: RunProps['rPrUnknown'],
  state: SerializeState,
): void {
  if (unknownChildren === undefined) {
    return
  }

  for (const entry of unknownChildren) {
    const placeholder = buildRawPassthroughPlaceholder(entry.xml, state)
    const anchorIndex =
      entry.before === undefined ? -1 : children.findIndex((node) => elementTagName(node) === entry.before)

    if (anchorIndex === -1) {
      children.push(placeholder)
    } else {
      children.splice(anchorIndex, 0, placeholder)
    }
  }
}

function elementTagName(node: OrderedXmlNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key !== ':@') {
      return key
    }
  }

  return undefined
}

// Child order follows ECMA-376 `CT_PPrBase`/`CT_PPr` (§17.3.1.26), restricted
// to the members `ParaProps` models, with `sectPr` last (the `CT_PPr`
// extension's own trailing member, after `rPr`/`sectPr`/`pPrChange` — Atlas
// doesn't model a paragraph-mark `rPr` or `pPrChange` here). An element-order
// review found `numPr`/`spacing`/`ind`/`jc` emitted right after `pStyle`
// instead of deep in the sequence, `keepNext`/`keepLines`/`pageBreakBefore`/
// `framePr`/`widowControl` pushed to the middle instead of right after
// `pStyle`, and `w:bidi` placed last (immediately before `sectPr`) although
// `spacing`/`ind`/`jc`/`textAlignment`/`outlineLvl`/`divId` all follow it in
// the real sequence.
function buildParagraphPropertiesNode(paraProps: ParaProps | undefined): OrderedXmlNode | undefined {
  if (paraProps === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []

  pushIfDefined(children, buildValueElement('w:pStyle', paraProps.pStyle))
  pushIfDefined(children, buildToggleElement('w:keepNext', paraProps.keepNext))
  pushIfDefined(children, buildToggleElement('w:keepLines', paraProps.keepLines))
  pushIfDefined(children, buildToggleElement('w:pageBreakBefore', paraProps.pageBreakBefore))
  pushIfDefined(children, buildFramePropsElement(paraProps.framePr))
  pushIfDefined(children, buildToggleElement('w:widowControl', paraProps.widowControl))
  pushIfDefined(children, buildNumPrElement(paraProps.numPr))
  pushIfDefined(children, buildToggleElement('w:suppressLineNumbers', paraProps.suppressLineNumbers))
  pushIfDefined(children, buildBorderSetElement('w:pBdr', paraProps.pBdr))
  pushIfDefined(children, buildShadingElement('w:shd', paraProps.shd))
  pushIfDefined(children, buildTabsElement(paraProps.tabs))
  pushIfDefined(children, buildToggleElement('w:suppressAutoHyphens', paraProps.suppressAutoHyphens))
  // Round-trip fidelity audit (DXS round 2): `w:bidi` was parsed onto
  // `ParaProps.bidi` (see parser/document.ts) but never written back here —
  // every right-to-left paragraph silently reverted to left-to-right on
  // save.
  pushIfDefined(children, buildToggleElement('w:bidi', paraProps.bidi))
  pushIfDefined(children, buildSpacingElement(paraProps.spacing))
  pushIfDefined(children, buildIndentElement(paraProps.ind))
  pushIfDefined(children, buildToggleElement('w:contextualSpacing', paraProps.contextualSpacing))
  pushIfDefined(children, buildToggleElement('w:mirrorIndents', paraProps.mirrorIndents))
  pushIfDefined(children, buildValueElement('w:jc', paraProps.jc))
  pushIfDefined(children, buildValueElement('w:textAlignment', paraProps.textAlignment))
  pushIfDefined(children, buildValueElement('w:outlineLvl', paraProps.outlineLvl))
  pushIfDefined(children, buildValueElement('w:divId', paraProps.divId))

  if (paraProps.sectPr !== undefined) {
    children.push(buildSectionPropertiesNode(paraProps.sectPr))
  }

  return children.length > 0 ? createElement('w:pPr', children) : undefined
}

// Child order follows ECMA-376 Part 1 §17.6.17's `EG_SectPrContents` group
// (shared by `CT_SectPrBase`/`CT_SectPr`), restricted to the members Atlas
// models: headerReference*, footerReference*, [footnotePr], [endnotePr],
// type, pgSz, pgMar, [paperSrc], pgBorders, lnNumType, pgNumType, cols,
// formProt, vAlign, noEndnote, titlePg, textDirection, bidi, rtlGutter,
// docGrid, [printerSettings], [sectPrChange] — bracketed members aren't
// modeled by `SectionProps` and are omitted rather than reordered. A
// structural-fidelity review (element-order pass) found the header/footer
// references emitted *last* instead of first, and `titlePg`/`pgNumType`/
// `cols` out of their relative sequence — the classic cause of Word's
// "found unreadable content" repair prompt, since `w:sectPr`'s content
// model is `xsd:sequence`, not a bag.
function buildSectionPropertiesNode(sectionProps: SectionProps): OrderedXmlNode {
  const children: OrderedXmlNode[] = []

  for (const headerReference of sectionProps.headerReference ?? []) {
    children.push(buildHeaderReferenceNode(headerReference))
  }

  for (const footerReference of sectionProps.footerReference ?? []) {
    children.push(buildFooterReferenceNode(footerReference))
  }

  pushIfDefined(children, buildValueElement('w:type', sectionProps.type))
  pushIfDefined(children, buildPageSizeElement(sectionProps.pgSz))
  pushIfDefined(children, buildPageMarginsElement(sectionProps.pgMar))
  // Round-trip fidelity audit (DXS round 2): `w:pgBorders` was previously
  // unmodeled entirely (not parsed, not passed through) — silently dropped
  // on every save. See `SectionProps.pgBorders`'s doc comment.
  pushIfDefined(children, buildPageBordersElement(sectionProps))
  pushIfDefined(children, buildLineNumberTypeElement(sectionProps.lnNumType))
  pushIfDefined(children, buildPageNumberTypeElement(sectionProps.pgNumType))
  pushIfDefined(children, buildSectionColumnsElement(sectionProps.cols))
  // The rest of `CT_SectPrBase` found unmodeled during the round-trip
  // fidelity audit (DXS round 2) — see `SectionProps`'s doc comments.
  pushIfDefined(children, buildToggleElement('w:formProt', sectionProps.formProt))
  pushIfDefined(children, buildValueElement('w:vAlign', sectionProps.vAlign))
  pushIfDefined(children, buildToggleElement('w:noEndnote', sectionProps.noEndnote))
  pushIfDefined(children, buildToggleElement('w:titlePg', sectionProps.titlePg))
  pushIfDefined(children, buildValueElement('w:textDirection', sectionProps.textDirection))
  // `w:bidi` on `w:sectPr` itself (section reads right-to-left) — distinct
  // from a paragraph's own `w:bidi` (see buildParagraphPropertiesNode).
  pushIfDefined(children, buildToggleElement('w:bidi', sectionProps.bidi))
  pushIfDefined(children, buildToggleElement('w:rtlGutter', sectionProps.rtlGutter))
  pushIfDefined(children, buildDocGridElement(sectionProps.docGrid))

  return createElement('w:sectPr', children)
}

function buildDocGridElement(docGrid: DocGrid | undefined): OrderedXmlNode | undefined {
  if (docGrid === undefined) {
    return undefined
  }

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:type', docGrid.type)
  appendAttribute(attributes, '@_w:linePitch', stringifyNumber(docGrid.linePitch))
  appendAttribute(attributes, '@_w:charSpace', stringifyNumber(docGrid.charSpace))

  return hasAttributes(attributes) ? createElement('w:docGrid', [], attributes) : undefined
}

function buildPageBordersElement(sectionProps: SectionProps): OrderedXmlNode | undefined {
  if (sectionProps.pgBorders === undefined) {
    return undefined
  }

  const children: OrderedXmlNode[] = []
  pushIfDefined(children, buildBorderElement('w:top', sectionProps.pgBorders.top))
  pushIfDefined(children, buildBorderElement('w:left', sectionProps.pgBorders.left))
  pushIfDefined(children, buildBorderElement('w:bottom', sectionProps.pgBorders.bottom))
  pushIfDefined(children, buildBorderElement('w:right', sectionProps.pgBorders.right))

  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:display', sectionProps.pgBorderDisplay)
  appendAttribute(attributes, '@_w:offsetFrom', sectionProps.pgBorderOffsetFrom)
  appendAttribute(attributes, '@_w:zOrder', sectionProps.pgBorderZOrder)

  return createElement('w:pgBorders', children, hasAttributes(attributes) ? attributes : undefined)
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

// DOCX-1 — Word has specified a style/run's font by theme reference (rather
// than a literal name) since Office 2007, Normal's default font most
// commonly of all. `parseFontSet` (parser/document.ts) has always read
// `w:asciiTheme`/`w:hAnsiTheme`/`w:cstheme`/`w:eastAsiaTheme` onto `FontSet`,
// but this builder only ever emitted the literal attributes — so the theme
// reference silently vanished on every save, and where no literal
// `w:ascii`/etc. accompanied it (the common case for a style that only sets
// a theme font), `hasAttributes` saw nothing at all and dropped the whole
// `w:rFonts` element.
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
  appendAttribute(attributes, '@_w:asciiTheme', fonts.asciiTheme)
  appendAttribute(attributes, '@_w:hAnsiTheme', fonts.hAnsiTheme)
  appendAttribute(attributes, '@_w:cstheme', fonts.csTheme)
  appendAttribute(attributes, '@_w:eastAsiaTheme', fonts.eastAsiaTheme)

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
  return buildRawPassthroughPlaceholder(node.xml, state)
}

/**
 * Splices an arbitrary raw XML string (one or several sibling elements —
 * the placeholder substitution is a plain string replace, so it doesn't
 * care) into the tree at this exact spot, the same mechanism
 * `buildUnknownPlaceholder` uses for `UnknownNode`. Also used for an
 * unmodified `Field`'s `raw` (D19/DXS-20 — see `buildFieldNodes`):
 * byte-faithful passthrough for a field the user hasn't run "Update
 * field(s)" on is the exact same problem as passthrough for any other
 * unsupported/unmodified source content.
 */
function buildRawPassthroughPlaceholder(xml: string, state: SerializeState): OrderedXmlNode {
  const id = `unknown-${state.nextUnknownId}`
  state.nextUnknownId += 1
  state.unknownXml.set(id, xml)
  return createElement(UNKNOWN_PLACEHOLDER_TAG, [], {
    '@_data-id': id,
  })
}

/**
 * DEFER-5 / DXS-20 — builds the XML for one field child (`w:fldSimple`, or
 * the run sequence for a complex field). An unmodified field (`raw` still
 * present — the overwhelming common case, since "Update field(s)" is an
 * explicit user action) round-trips via the exact captured source text
 * (`buildRawPassthroughPlaceholder`), the same byte-faithful passthrough
 * `UnknownNode` uses. Only a field the user has explicitly regenerated
 * (`raw` cleared — see `src/docx/fields`) gets structurally rebuilt from
 * `instruction`/`result` here.
 *
 * Returns an array (rather than the single `OrderedXmlNode` every other
 * `ParagraphChild` builder returns) because a regenerated complex field
 * fundamentally needs several sibling `<w:r>` elements — begin/instrText/
 * separate/result(s)/end — not one element with children; see
 * `buildParagraphChildNodes`'s doc comment for how callers handle that.
 *
 * Takes no `asDel` parameter, unlike its sibling `build*ChildNodes`
 * wrappers: a regenerated field's result comes from Atlas's own evaluators
 * (`src/docx/fields`), which never produce tracked-deletion content, so
 * there is no `asDel` context to thread through here even when the field
 * itself sits inside a `w:del` (its `raw`-passthrough form, taken above,
 * doesn't need one either — it's spliced in as literal text).
 */
function buildFieldNodes(field: Field, state: SerializeState): ReadonlyArray<OrderedXmlNode> {
  if (field.raw !== undefined) {
    return [buildRawPassthroughPlaceholder(field.raw, state)]
  }

  return field.simple === true
    ? [buildSimpleFieldNode(field, state)]
    : buildComplexFieldNodes(field, state)
}

function buildSimpleFieldNode(field: Field, state: SerializeState): OrderedXmlNode {
  const attributes = createAttributes()
  appendAttribute(attributes, '@_w:instr', field.instruction)
  appendAttribute(attributes, '@_w:fldLock', buildOnOffAttribute(field.locked))
  appendAttribute(attributes, '@_w:dirty', buildOnOffAttribute(field.dirty))

  const children: OrderedXmlNode[] = []
  for (const child of field.result) {
    children.push(...buildParagraphChildNodes(child, state))
  }

  return createElement('w:fldSimple', children, hasAttributes(attributes) ? attributes : undefined)
}

function buildComplexFieldNodes(field: Field, state: SerializeState): ReadonlyArray<OrderedXmlNode> {
  const beginAttributes = createAttributes()
  appendAttribute(beginAttributes, '@_w:fldCharType', 'begin')
  appendAttribute(beginAttributes, '@_w:fldLock', buildOnOffAttribute(field.locked))
  appendAttribute(beginAttributes, '@_w:dirty', buildOnOffAttribute(field.dirty))

  const beginRun = createElement('w:r', [createElement('w:fldChar', [], beginAttributes)])
  const instrRun = createElement('w:r', [
    createElement('w:instrText', [createText(field.instruction)], { '@_xml:space': 'preserve' }),
  ])
  const separateRun = createElement('w:r', [
    createElement('w:fldChar', [], { '@_w:fldCharType': 'separate' }),
  ])
  const endRun = createElement('w:r', [createElement('w:fldChar', [], { '@_w:fldCharType': 'end' })])

  const resultNodes: OrderedXmlNode[] = []
  for (const child of field.result) {
    resultNodes.push(...buildParagraphChildNodes(child, state))
  }

  return [beginRun, instrRun, separateRun, ...resultNodes, endRun]
}

/**
 * The namespace prefixes Atlas's own serializer relies on somewhere in its
 * output (directly, or via unknown-node passthrough content that commonly
 * uses one of these). Shared with `partWriterSupport.ts` (DXS-08) so
 * standalone parts (header/footer/footnote/endnote/comment) declare the
 * same set — previously they declared only `xmlns:w` (+ `xmlns:r`), so a
 * drawing/revision/etc. inside one of those parts emitted an undeclared
 * namespace prefix, which is an XML well-formedness violation Word may
 * reject or "repair" on open.
 */
export const STANDARD_NAMESPACE_URIS: Readonly<Record<string, string>> = {
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  o: 'urn:schemas-microsoft-com:office:office',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  w10: 'urn:schemas-microsoft-com:office:word',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}

const BASELINE_MC_IGNORABLE = 'w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14'

/** `{prefix: uri}` -> `{'@_xmlns:prefix': uri}`, for splicing into an XmlAttributes record. */
export function buildNamespaceDeclarationAttributes(
  namespaces: Readonly<Record<string, string>>,
): XmlAttributes {
  const attributes = createAttributes()
  for (const [prefix, uri] of Object.entries(namespaces)) {
    attributes[`@_xmlns:${prefix}`] = uri
  }
  return attributes
}

/**
 * Unions two `mc:Ignorable` token lists (space-separated namespace
 * prefixes), deduplicated, preserving baseline order first.
 */
function mergeIgnorableTokens(baseline: string, extra: string | undefined): string {
  const tokens: string[] = []
  const seen = new Set<string>()

  for (const token of `${baseline} ${extra ?? ''}`.split(/\s+/)) {
    if (token === '' || seen.has(token)) {
      continue
    }
    seen.add(token)
    tokens.push(token)
  }

  return tokens.join(' ')
}

/**
 * DXS-15: the document root's namespace declarations previously came from
 * a fixed, hardcoded set regardless of what the source document actually
 * declared. Now unions that baseline with whatever extra namespace
 * prefixes the source root declared (`doc.rootNamespaces`, captured by the
 * parser) — the baseline wins on the prefixes Atlas's own serializer
 * depends on (so e.g. `xmlns:w` always resolves to the wordprocessingml
 * URI regardless of what a source document did), while any additional,
 * unrecognized prefix the source declared — most relevantly one used only
 * by unknown-node passthrough content this serializer doesn't otherwise
 * understand — survives instead of being silently dropped. Likewise
 * `mc:Ignorable` unions the baseline token list with the source's.
 */
function buildDocumentAttributes(doc: Document): XmlAttributes {
  const attributes: Record<string, string | undefined> = {
    ...(doc.rootNamespaces !== undefined
      ? buildNamespaceDeclarationAttributes(Object.fromEntries(doc.rootNamespaces))
      : {}),
    ...buildNamespaceDeclarationAttributes(STANDARD_NAMESPACE_URIS),
  }
  attributes['@_mc:Ignorable'] = mergeIgnorableTokens(BASELINE_MC_IGNORABLE, doc.mcIgnorable)
  return attributes
}

export function createSerializeState(wrapperRegions: ReadonlyArray<WrapperPassthrough> = []): SerializeState {
  return {
    unknownXml: new Map(),
    nextUnknownId: 0,
    wrapperRegions,
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

export function restoreUnknownXml(xml: string, state: SerializeState): string {
  let restored = xml

  for (const [id, rawXml] of state.unknownXml.entries()) {
    const openClose = `<${UNKNOWN_PLACEHOLDER_TAG} data-id="${id}"></${UNKNOWN_PLACEHOLDER_TAG}>`
    const selfClosing = `<${UNKNOWN_PLACEHOLDER_TAG} data-id="${id}"/>`
    restored = restored.split(openClose).join(rawXml)
    restored = restored.split(selfClosing).join(rawXml)
  }

  return restored
}
