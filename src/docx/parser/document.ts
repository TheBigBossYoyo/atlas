/**
 * Atlas — DOCX document parser (Wave A.3)
 *
 * Parses `word/document.xml` into the immutable AST defined in Wave A.2.
 * The parser keeps OOXML child order, extracts the supported block/inline
 * surface, and falls back to `UnknownNode` for unsupported structural nodes.
 */

import { XMLBuilder, XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
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
  type DrawingAnchorChild,
  type DrawingCrop,
  type DrawingEffectExtent,
  type DrawingHorizontalAlign,
  type DrawingPositionH,
  type DrawingPositionV,
  type DrawingRelativeFromH,
  type DrawingRelativeFromV,
  type DrawingTransform,
  type DrawingVerticalAlign,
  type DrawingWrap,
  type DrawingWrapMode,
  type DrawingWrapSide,
  type EmphasisMark,
  type Endnote,
  type EndnoteReference,
  type Field,
  type FieldType,
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
  type DocGrid,
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
  type RPrUnknownChild,
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
  type WrapperPassthrough,
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

/**
 * D19 / DXS-16 — raw-substring context for the document XML currently being
 * parsed, consulted by `parseUnknownNode` so an unsupported element's
 * captured `xml` is the exact source text (byte-faithful, including exotic
 * formatting fast-xml-parser's builder wouldn't reproduce) rather than a
 * tree-rebuilt approximation, with any namespace prefix it relies on that's
 * declared by a non-root ancestor re-declared locally so the fragment stays
 * valid wherever the serializer splices it back in (see `computeRawNodeInfo`'s
 * doc comment for why that specific gap is the one worth closing here).
 *
 * Module-level rather than threaded as a parameter through every one of
 * `parseUnknownNode`'s ~10 call sites: `parseDocument` is this module's only
 * export and is never reentered mid-parse (no `await`, and nothing else in
 * this file calls back into it), so there is exactly one active parse at a
 * time — set at entry, read throughout, cleared in `finally` so a thrown
 * parse error never leaves a stale context (holding a potentially large
 * string) for a later call to accidentally reuse.
 */
let sourceRangeContext: SourceRangeContext | undefined

/**
 * `w:sdt`/`mc:AlternateContent` wrapper regions captured while parsing the
 * body of the document currently being parsed (round-trip fidelity audit,
 * DXS round 2 follow-up to D8/DXP-08) — see `WrapperPassthrough`'s doc
 * comment on `../model/document.ts`. Module-level for the same reason
 * `sourceRangeContext` is (see its own doc comment just above): exactly one
 * parse active at a time, set at entry, appended to throughout, cleared in
 * `finally`.
 */
let wrapperRegions: WrapperPassthrough[] | undefined

export function parseDocument(xml: string): DocxDocument {
  assertXmlPartSizeWithinLimit(xml, 'word/document.xml')
  let raw: OrderedXmlNode[]
  try {
    raw = xmlParser.parse(xml) as OrderedXmlNode[]
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse document XML: ${msg}`)
  }

  sourceRangeContext = computeSourceRangeContext(xml, raw)
  wrapperRegions = []
  try {
    return parseDocumentTree(raw)
  } finally {
    sourceRangeContext = undefined
    wrapperRegions = undefined
  }
}

function parseDocumentTree(raw: ReadonlyArray<OrderedXmlNode>): DocxDocument {
  const documentElement = findElement(raw, 'w:document')
  const bodyElement = child(documentElement, 'w:body')
  const body = parseBody(bodyElement)
  const rootNamespaces = parseRootNamespaces(documentElement)
  const mcIgnorable = attr(documentElement, 'mc:Ignorable')

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
    ...(rootNamespaces.size > 0 ? { rootNamespaces } : {}),
    ...(mcIgnorable !== undefined ? { mcIgnorable } : {}),
    ...(wrapperRegions !== undefined && wrapperRegions.length > 0 ? { wrappers: wrapperRegions } : {}),
  }
}

/**
 * D19 / DXS-15: captures every `xmlns:*` declaration the source
 * `<w:document>` root actually carried, so `documentWriter.ts` can union
 * it with Atlas's own required baseline set instead of emitting a fixed
 * hardcoded namespace list regardless of what the source declared.
 */
function parseRootNamespaces(element: OrderedXmlNode | undefined): ReadonlyMap<string, string> {
  const namespaces = new Map<string, string>()
  const attributes = element?.[':@']
  if (attributes === undefined) {
    return namespaces
  }

  const XMLNS_PREFIX = '@_xmlns:'
  for (const [key, value] of Object.entries(attributes)) {
    if (key.startsWith(XMLNS_PREFIX) && value !== undefined) {
      namespaces.set(key.slice(XMLNS_PREFIX.length), value)
    }
  }

  return namespaces
}

// ---------------------------------------------------------------------------
// D8 / DXP-08 — w:sdt / mc:AlternateContent unwrapping
//
// Word wraps content in `w:sdt` (a "structured document tag" — content
// controls, TOC/repeating-section fields, and similar) and in
// `mc:AlternateContent` (the markup-compatibility mechanism Word uses to
// offer both a modern and a legacy representation of the same content,
// most commonly shapes/drawings) far more pervasively than the parser
// previously accounted for. Both were entirely unhandled: every sibling
// list iterated below fell through to `parseUnknownNode`, so a document
// using either — a TOC built from a content control, a shape wrapped for
// backward compatibility — rendered nothing where that content should be.
//
// Neither actually needs its own model type: unwrapping to the real inner
// content at parse time and re-dispatching it through the same per-kind
// switch the caller already has recovers the content "for free". This
// necessarily means round-tripping a document through Atlas discards the
// `w:sdt`/`mc:AlternateContent` wrapper itself (its content control
// metadata, or its rejected alternate branch) — an accepted, documented
// trade-off for making the content visible at all, consistent with this
// plan's freeze on new structural editing surface.
//
// Round-trip fidelity audit, DXS round 2 follow-up: at the three call
// sites where a wrapper can sit as a direct sibling of body-level blocks,
// paragraph-level children, or run-level children (`parseBody`,
// `parseParagraph`, `parseRun` — the shapes real Word documents
// overwhelmingly use both wrappers in), `captureWrapperRaw`/
// `recordWrapperRegion` below additionally remember the wrapper's exact
// source bytes alongside the very model object(s) its content unwrapped
// to, as a `WrapperPassthrough` region on the document. This doesn't
// change the "no new structural editing surface" trade-off above at all —
// the visible model is identical either way — it only lets
// `documentWriter.ts` opportunistically re-emit the original wrapper
// byte-for-byte on save when an identity check proves none of its content
// was touched, rather than always discarding it. See `WrapperPassthrough`'s
// doc comment on `../model/document.ts` for how that identity check works
// and why it needs no cooperation from the edit pipeline.
// ---------------------------------------------------------------------------

/**
 * Expands `w:sdt` and `mc:AlternateContent` wrapper elements in a sibling
 * list into their real inner children, recursively (a wrapper can contain
 * another wrapper), so every call site below can iterate the result with
 * its normal per-kind switch as if the wrappers were never there.
 */
function expandWrapperNodes(entries: ReadonlyArray<OrderedXmlNode>): ReadonlyArray<OrderedXmlNode> {
  const expanded: OrderedXmlNode[] = []

  for (const entry of entries) {
    switch (nodeName(entry)) {
      case 'w:sdt':
        expanded.push(...expandWrapperNodes(nodeChildren(child(entry, 'w:sdtContent'))))
        break
      case 'mc:AlternateContent':
        expanded.push(...expandWrapperNodes(resolveAlternateContentChildren(entry)))
        break
      default:
        expanded.push(entry)
        break
    }
  }

  return expanded
}

/**
 * Resolves which branch of an `mc:AlternateContent` wrapper to keep.
 *
 * Prefers the first `mc:Choice`: in real-world Word output this is
 * overwhelmingly the modern DrawingML/wordprocessingShape representation,
 * which Atlas's existing drawing parser already understands at least for
 * plain pictures, whereas `mc:Fallback` overwhelmingly carries VML
 * (`v:shape`/`w:pict`) markup Atlas has no model for at all — preferring
 * Fallback would frequently still show nothing, defeating the point. Falls
 * back to `mc:Fallback` only when there is no `mc:Choice` element at all.
 */
function resolveAlternateContentChildren(element: OrderedXmlNode): ReadonlyArray<OrderedXmlNode> {
  const choice = child(element, 'mc:Choice')
  if (choice !== undefined) {
    return nodeChildren(choice)
  }

  const fallback = child(element, 'mc:Fallback')
  if (fallback !== undefined) {
    return nodeChildren(fallback)
  }

  return []
}

/** `entry`'s tag name if it's a wrapper this module knows how to passthrough-capture, else `undefined`. */
function wrapperKindOf(entry: OrderedXmlNode): 'w:sdt' | 'mc:AlternateContent' | undefined {
  const name = nodeName(entry)
  return name === 'w:sdt' || name === 'mc:AlternateContent' ? name : undefined
}

/**
 * The exact source XML for one `w:sdt`/`mc:AlternateContent` element,
 * captured the same way `captureRawSpan`/`parseUnknownNode` do.
 * `undefined` only for a lookup miss (no active `sourceRangeContext` — a
 * unit test calling this file's parse helpers directly rather than through
 * `parseDocument`).
 */
function captureWrapperRaw(entry: OrderedXmlNode): string | undefined {
  const info = sourceRangeContext?.info.get(entry)
  return info !== undefined ? sourceRangeContext!.xml.slice(info.start, info.end) : undefined
}

/** Appends one wrapper-passthrough region to this parse's accumulator, if there is one (see `wrapperRegions`'s doc comment). */
function recordWrapperRegion(region: WrapperPassthrough): void {
  wrapperRegions?.push(region)
}

function parseBody(element: OrderedXmlNode | undefined): {
  readonly blocks: ReadonlyArray<Block>
  readonly sectionProps?: SectionProps
} {
  const blocks: Block[] = []
  let sectionProps: SectionProps | undefined

  const pushEntry = (entry: OrderedXmlNode): void => {
    if (isIgnorableText(entry)) {
      return
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

  // Round-trip fidelity audit, DXS round 2 follow-up: a body-level
  // `w:sdt`/`mc:AlternateContent` (a content control or shape wrapping one
  // or more whole paragraphs/tables) is expanded one entry at a time here,
  // rather than pre-flattening the whole sibling list the way the old
  // single `expandWrapperNodes(nodeChildren(element))` loop did, so its
  // exact source span and the block(s) it expanded to can be captured
  // together as a `WrapperPassthrough` region — see `captureWrapperRaw`/
  // `recordWrapperRegion`'s doc comments. A non-wrapper entry goes straight
  // through `pushEntry` unchanged, identical to the old behavior.
  for (const entry of nodeChildren(element)) {
    const wrapper = wrapperKindOf(entry)
    if (wrapper === undefined) {
      pushEntry(entry)
      continue
    }

    const raw = captureWrapperRaw(entry)
    const before = blocks.length
    for (const inner of expandWrapperNodes([entry])) {
      pushEntry(inner)
    }
    if (raw !== undefined) {
      const content = blocks.slice(before)
      if (content.length > 0) {
        recordWrapperRegion({ wrapper, raw, content })
      }
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

  const consumeGroup = (entries: ReadonlyArray<OrderedXmlNode>): void => {
    for (const entry of groupComplexFieldRuns(entries)) {
      if (isComplexFieldGroup(entry)) {
        children.push(parseComplexField(entry.runs))
        continue
      }

      if (isIgnorableText(entry)) {
        continue
      }

      const parsedChild = parseParagraphChild(entry)
      if (parsedChild !== null) {
        children.push(parsedChild)
      }
    }
  }

  // Round-trip fidelity audit, DXS round 2 follow-up: same one-entry-at-a-
  // time restructuring as `parseBody`'s, for the same reason — a paragraph-
  // level wrapper (most commonly a content control around inline text)
  // needs its exact span and resulting children captured together. Runs of
  // consecutive non-wrapper siblings are buffered into `pending` and run
  // through the original `groupComplexFieldRuns(expandWrapperNodes(...))`
  // pipeline together (`consumeGroup`), so a complex field's begin/.../end
  // `w:r` siblings still group correctly as long as nothing wraps only PART
  // of that span — the same assumption real, non-pathological Word output
  // already satisfies.
  let pending: OrderedXmlNode[] = []
  for (const entry of nodeChildren(element)) {
    const wrapper = wrapperKindOf(entry)
    if (wrapper === undefined) {
      pending.push(entry)
      continue
    }

    if (pending.length > 0) {
      consumeGroup(expandWrapperNodes(pending))
      pending = []
    }

    const raw = captureWrapperRaw(entry)
    const before = children.length
    consumeGroup(expandWrapperNodes([entry]))
    if (raw !== undefined) {
      const content = children.slice(before)
      if (content.length > 0) {
        recordWrapperRegion({ wrapper, raw, content })
      }
    }
  }
  if (pending.length > 0) {
    consumeGroup(expandWrapperNodes(pending))
  }

  return {
    kind: 'paragraph',
    ...(props !== undefined ? { props } : {}),
    children,
    ...withAttrValue('paraId', attr(element, 'w14:paraId')),
    ...withAttrValue('textId', attr(element, 'w14:textId')),
    ...withAttrValue('rsidR', attr(element, 'w:rsidR')),
    ...withAttrValue('rsidRDefault', attr(element, 'w:rsidRDefault')),
    ...withAttrValue('rsidP', attr(element, 'w:rsidP')),
    ...withAttrValue('rsidRPr', attr(element, 'w:rsidRPr')),
  }
}

function withAttrValue<K extends string>(
  key: K,
  value: string | undefined,
): { [P in K]?: string } {
  return value !== undefined ? ({ [key]: value } as { [P in K]?: string }) : {}
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
    case 'w:fldSimple':
      return parseSimpleField(element)
    default:
      return parseUnknownNode(element)
  }
}

function parseRun(element: OrderedXmlNode): Run {
  const props = parseRunProps(child(element, 'w:rPr'))
  const children: RunChild[] = []

  const pushEntry = (entry: OrderedXmlNode): void => {
    if (isIgnorableText(entry) || nodeName(entry) === 'w:rPr') {
      return
    }

    switch (nodeName(entry)) {
      case 'w:t':
        children.push(parseTextNode(entry))
        break
      case 'w:delText':
        children.push(parseTextNode(entry))
        break
      case 'w:sym':
        children.push(parseSym(entry))
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

  // Round-trip fidelity audit, DXS round 2 follow-up: same one-entry-at-a-
  // time restructuring as `parseBody`'s/`parseParagraph`'s — this is where
  // real Word documents put `mc:AlternateContent` (a shape/text box's
  // modern-vs-legacy-VML fallback pair sits directly inside the run that
  // "contains" the drawing), so capturing its span here is what lets a
  // shape/text box survive a save untouched.
  for (const entry of nodeChildren(element)) {
    const wrapper = wrapperKindOf(entry)
    if (wrapper === undefined) {
      pushEntry(entry)
      continue
    }

    const raw = captureWrapperRaw(entry)
    const before = children.length
    for (const inner of expandWrapperNodes([entry])) {
      pushEntry(inner)
    }
    if (raw !== undefined) {
      const content = children.slice(before)
      if (content.length > 0) {
        recordWrapperRegion({ wrapper, raw, content })
      }
    }
  }

  return {
    kind: 'run',
    ...(props !== undefined ? { props } : {}),
    children,
    ...withAttrValue('rsidR', attr(element, 'w:rsidR')),
    ...withAttrValue('rsidRPr', attr(element, 'w:rsidRPr')),
    ...withAttrValue('rsidDel', attr(element, 'w:rsidDel')),
  }
}

function parseHyperlink(element: OrderedXmlNode): Hyperlink {
  const relationshipId = attr(element, 'r:id')
  const anchor = attr(element, 'w:anchor')
  const tooltip = attr(element, 'w:tooltip')
  const targetFrame = attr(element, 'w:tgtFrame')
  const history = parseOnOff(attr(element, 'w:history'))
  const children: HyperlinkChild[] = []

  for (const entry of groupComplexFieldRuns(expandWrapperNodes(nodeChildren(element)))) {
    if (isComplexFieldGroup(entry)) {
      children.push(parseComplexField(entry.runs))
      continue
    }

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
      case 'w:fldSimple':
        children.push(parseSimpleField(entry))
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

// ---------------------------------------------------------------------------
// DEFER-5 / DXS-20 — field parsing (w:fldSimple, and complex fields built
// from w:fldChar begin/separate/end + w:instrText)
// ---------------------------------------------------------------------------

/**
 * A complex field's begin/instrText/separate/result/end pieces are SIBLING
 * `w:r` elements at the paragraph/hyperlink level, not nested inside one —
 * unlike every other `ParagraphChild`, one `Field` model node has to be
 * assembled from a whole RUN of sibling entries rather than a single one.
 * `groupComplexFieldRuns` does that grouping as a preprocessing pass over
 * an already `expandWrapperNodes`-flattened sibling list, so
 * `parseParagraph`/`parseHyperlink`'s loops can treat a whole field like
 * any other single child once grouped.
 */
interface ComplexFieldGroup {
  readonly kind: 'complex-field-group'
  /** Only the `w:r` entries from the begin run through the end run, inclusive — see `groupComplexFieldRuns`'s doc comment on why non-run siblings in between are dropped from this list (though not from the field's raw byte span). */
  readonly runs: ReadonlyArray<OrderedXmlNode>
}

function isComplexFieldGroup(
  entry: OrderedXmlNode | ComplexFieldGroup,
): entry is ComplexFieldGroup {
  return (entry as ComplexFieldGroup).kind === 'complex-field-group'
}

/**
 * Scans `entries` for `w:r`-begin ... `w:r`-end runs (a "complex field")
 * and replaces each whole span with one `ComplexFieldGroup`, tracking
 * nesting depth so a field whose instruction or result legitimately
 * contains another field (rare, but valid — e.g. an `IF` field nesting a
 * `REF`) doesn't have its outer span cut short at the FIRST `end` it finds.
 * Every other entry (including a truncated field with no matching `end`,
 * left as ordinary `w:r` content) passes through unchanged.
 */
function groupComplexFieldRuns(
  entries: ReadonlyArray<OrderedXmlNode>,
): ReadonlyArray<OrderedXmlNode | ComplexFieldGroup> {
  const result: Array<OrderedXmlNode | ComplexFieldGroup> = []
  let index = 0

  while (index < entries.length) {
    const entry = entries[index]

    if (entry !== undefined && nodeName(entry) === 'w:r' && fldCharTypeOf(entry) === 'begin') {
      const group = collectComplexFieldGroup(entries, index)
      if (group !== undefined) {
        result.push(group.value)
        index = group.nextIndex
        continue
      }
    }

    if (entry !== undefined) {
      result.push(entry)
    }
    index += 1
  }

  return result
}

function collectComplexFieldGroup(
  entries: ReadonlyArray<OrderedXmlNode>,
  beginIndex: number,
): { readonly value: ComplexFieldGroup; readonly nextIndex: number } | undefined {
  const beginEntry = entries[beginIndex]
  if (beginEntry === undefined) {
    return undefined
  }

  const runs: OrderedXmlNode[] = [beginEntry]
  let depth = 1

  for (let i = beginIndex + 1; i < entries.length; i += 1) {
    const entry = entries[i]
    if (entry === undefined || nodeName(entry) !== 'w:r') {
      // Non-run content inside a field's span (a bookmark, stray whitespace
      // text, ...) still belongs to the field's raw byte range (captured
      // separately via `captureFieldRawSpan`'s begin/end anchors) but has no
      // structural role in `Field.result` — skip it from `runs` without
      // abandoning the scan.
      continue
    }

    runs.push(entry)
    const type = fldCharTypeOf(entry)
    if (type === 'begin') {
      depth += 1
    } else if (type === 'end') {
      depth -= 1
      if (depth === 0) {
        return { value: { kind: 'complex-field-group', runs }, nextIndex: i + 1 }
      }
    }
  }

  // No matching end (truncated/malformed field): let the caller fall back
  // to treating the begin run as ordinary content instead of silently
  // consuming the rest of the paragraph looking for an end that isn't there.
  return undefined
}

function fldCharTypeOf(runEntry: OrderedXmlNode): string | undefined {
  const fldChar = child(runEntry, 'w:fldChar')
  return fldChar !== undefined ? attr(fldChar, 'w:fldCharType') : undefined
}

const KNOWN_FIELD_TYPES: ReadonlySet<string> = new Set([
  'DATE',
  'TIME',
  'AUTHOR',
  'TITLE',
  'REF',
  'PAGEREF',
  'SEQ',
  'NUMPAGES',
  'PAGE',
  'HYPERLINK',
  'TOC',
])

function parseFieldInstructionType(instruction: string): FieldType {
  const token = /^([A-Za-z]+)/.exec(instruction)?.[1]?.toUpperCase()
  return token !== undefined && KNOWN_FIELD_TYPES.has(token) ? (token as FieldType) : 'unknown'
}

/**
 * Assembles a `Field` from a `ComplexFieldGroup`'s `w:r` runs: everything
 * before the `separate` fldChar is instruction text (`w:instrText`
 * content, concatenated across every run — Word sometimes splits a long
 * instruction across several runs); everything after is the cached
 * display content, parsed as ordinary runs. A field with no `separate` at
 * all (legal — an unresolved field can be just begin/instrText/end) has no
 * cached result runs, matching Word's own "not yet calculated" state.
 */
function parseComplexField(runs: ReadonlyArray<OrderedXmlNode>): Field {
  let phase: 'instruction' | 'result' = 'instruction'
  let instruction = ''
  const result: ParagraphChild[] = []
  let locked: OnOff | undefined
  let dirty: OnOff | undefined

  for (const run of runs) {
    const fldChar = child(run, 'w:fldChar')
    if (fldChar !== undefined) {
      const type = attr(fldChar, 'w:fldCharType')
      if (type === 'begin') {
        locked = parseOnOff(attr(fldChar, 'w:fldLock'))
        dirty = parseOnOff(attr(fldChar, 'w:dirty'))
      } else if (type === 'separate') {
        phase = 'result'
      }
      continue
    }

    if (phase === 'instruction') {
      instruction += extractInstrText(run)
    } else {
      result.push(parseRun(run))
    }
  }

  const trimmedInstruction = instruction.trim()
  const firstRun = runs[0]
  const lastRun = runs[runs.length - 1]
  const raw = firstRun !== undefined && lastRun !== undefined
    ? captureRawSpan(firstRun, lastRun)
    : undefined

  return {
    kind: 'field',
    fieldType: parseFieldInstructionType(trimmedInstruction),
    instruction: trimmedInstruction,
    result,
    ...(locked !== undefined ? { locked } : {}),
    ...(dirty !== undefined ? { dirty } : {}),
    ...(raw !== undefined ? { raw } : {}),
  }
}

function extractInstrText(run: OrderedXmlNode): string {
  let text = ''
  for (const entry of nodeChildren(run)) {
    if (nodeName(entry) === 'w:instrText') {
      text += textValue(entry)
    }
  }
  return text
}

/**
 * Byte-faithful passthrough for an unmodified field (D19 / DXS-16's
 * infrastructure, reused here): slices the exact source text spanning
 * `startNode` through `endNode` inclusive, from the same raw-range index
 * `parseUnknownNode` uses. `undefined` only when one of the anchors has no
 * indexed range (lookup miss, or this ran outside a `parseDocument` call,
 * as unit tests calling this file's helpers directly do) — the field is
 * still fully usable (`instruction`/`result` are always populated), it
 * just gets structurally rebuilt on save instead of passed through
 * verbatim (same fallback shape as `parseUnknownNode`'s).
 */
function captureRawSpan(startNode: OrderedXmlNode, endNode: OrderedXmlNode): string | undefined {
  const startInfo = sourceRangeContext?.info.get(startNode)
  const endInfo = sourceRangeContext?.info.get(endNode)
  if (startInfo === undefined || endInfo === undefined) {
    return undefined
  }
  return sourceRangeContext!.xml.slice(startInfo.start, endInfo.end)
}

/**
 * `w:fldSimple` is the "simple field" form: a single element carrying the
 * instruction as its own `w:instr` attribute, wrapping its cached result
 * directly as normal paragraph-child content (most commonly one or more
 * `w:r` runs) — no `fldChar`/`instrText` machinery needed. Dispatches its
 * children through the same `parseParagraphChild` every paragraph uses, so
 * a simple field's result can contain anything a paragraph can (bookmarks,
 * even a nested field).
 */
function parseSimpleField(element: OrderedXmlNode): Field {
  const instruction = (attr(element, 'w:instr') ?? '').trim()
  const result: ParagraphChild[] = []

  for (const entry of expandWrapperNodes(nodeChildren(element))) {
    if (isIgnorableText(entry)) {
      continue
    }
    const parsedChild = parseParagraphChild(entry)
    if (parsedChild !== null) {
      result.push(parsedChild)
    }
  }

  const info = sourceRangeContext?.info.get(element)
  const raw = info !== undefined ? sourceRangeContext!.xml.slice(info.start, info.end) : undefined
  const locked = parseOnOff(attr(element, 'w:fldLock'))
  const dirty = parseOnOff(attr(element, 'w:dirty'))

  return {
    kind: 'field',
    fieldType: parseFieldInstructionType(instruction),
    instruction,
    result,
    simple: true,
    ...(locked !== undefined ? { locked } : {}),
    ...(dirty !== undefined ? { dirty } : {}),
    ...(raw !== undefined ? { raw } : {}),
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

/**
 * D8 / DXP-08: `<w:sym w:font="Wingdings" w:char="F0E0"/>` represents a
 * single symbol-font glyph by its position in that font's own encoding —
 * unlike everything else in a run, it's not a Unicode text run at all.
 * Atlas has no model field for "this text is in font X, positioned by raw
 * code point" (nor any of the symbol fonts themselves bundled), so this
 * decodes `w:char` as a plain UTF-16 code unit into a synthetic text node —
 * shows *something* in the position the character belongs, rather than
 * nothing, even though the bundled substitute fonts can't reproduce the
 * intended glyph. Documented as a best-effort approximation, not a
 * faithful rendering.
 */
function parseSym(element: OrderedXmlNode): TextNode {
  const charAttr = attr(element, 'w:char')
  const codeUnit = charAttr !== undefined ? Number.parseInt(charAttr, 16) : Number.NaN
  const value = Number.isFinite(codeUnit) && codeUnit > 0 ? String.fromCharCode(codeUnit) : ''

  return {
    kind: 'text',
    value,
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

  for (const entry of expandWrapperNodes(nodeChildren(element))) {
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

  if (blip === undefined) {
    // Charts, SmartArt, and other non-picture graphicFrame content have no
    // a:blip. Atlas doesn't model any of that yet, so preserve the whole
    // w:drawing verbatim rather than building a lossy, schema-invalid
    // partial Drawing with extent/docPr but no graphic content at all
    // (DXS-04) — the same fallback already used when there's no
    // wp:inline/wp:anchor at all.
    return parseUnknownNode(element)
  }

  const extentElement = findDescendant(layoutElement, 'wp:extent')
  const effectExtentElement = child(layoutElement, 'wp:effectExtent')
  const relationshipId =
    attr(blip, 'r:embed') ?? attr(blip, 'r:link')
  const title = attr(docPr, 'title')
  const description = attr(docPr, 'descr')
  const name = attr(docPr, 'name')
  const extent =
    extentElement !== undefined
      ? parseDrawingExtent(extentElement)
      : undefined
  const effectExtent =
    effectExtentElement !== undefined ? parseDrawingEffectExtent(effectExtentElement) : undefined
  // a:srcRect (crop) and a:xfrm (rotation/flip) live inside the pic:pic
  // subtree regardless of inline/anchor layout (DXS-09).
  const crop = parseDrawingCrop(findDescendant(layoutElement, 'a:srcRect'))
  const transform = parseDrawingTransform(findDescendant(layoutElement, 'a:xfrm'))
  const anchorPosition = anchor !== undefined ? parseAnchorPosition(anchor) : undefined
  const anchorChildren = anchor !== undefined ? parseAnchorChildren(anchor) : undefined

  return {
    kind: 'drawing',
    layout: inline !== undefined ? 'inline' : 'anchor',
    ...(relationshipId !== undefined ? { relationshipId } : {}),
    ...(title !== undefined ? { title } : {}),
    ...(description !== undefined ? { description } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(extent !== undefined ? { extent } : {}),
    ...(effectExtent !== undefined ? { effectExtent } : {}),
    ...(crop !== undefined ? { crop } : {}),
    ...(transform !== undefined ? { transform } : {}),
    ...(anchorPosition ?? {}),
    ...(anchorChildren !== undefined ? { anchorChildren } : {}),
  }
}

/**
 * Parses `wp:anchor`'s own attributes (`behindDoc`/`allowOverlap`) and its
 * `wp:positionH`/`wp:positionV`/wrap-choice children into typed `Drawing`
 * fields (DXP-09/DXL-03) — the model the layout engine (`floats.ts`)
 * consumes to place a floating image and compute text-wrap exclusions.
 * Returns only the fields found; `parseDrawing` spreads the result onto the
 * `Drawing` it builds.
 */
function parseAnchorPosition(
  anchorElement: OrderedXmlNode,
): Pick<Drawing, 'behindDoc' | 'allowOverlap' | 'positionH' | 'positionV' | 'wrap'> {
  const behindDoc = parseOnOff(attr(anchorElement, 'behindDoc'))
  const allowOverlap = parseOnOff(attr(anchorElement, 'allowOverlap'))
  const positionH = parseDrawingPositionH(child(anchorElement, 'wp:positionH'))
  const positionV = parseDrawingPositionV(child(anchorElement, 'wp:positionV'))
  const wrap = parseDrawingWrap(anchorElement)

  return {
    ...(behindDoc !== undefined ? { behindDoc } : {}),
    ...(allowOverlap !== undefined ? { allowOverlap } : {}),
    ...(positionH !== undefined ? { positionH } : {}),
    ...(positionV !== undefined ? { positionV } : {}),
    ...(wrap !== undefined ? { wrap } : {}),
  }
}

const DRAWING_RELATIVE_FROM_H: ReadonlySet<string> = new Set([
  'page',
  'margin',
  'column',
  'character',
  'leftMargin',
  'rightMargin',
  'insideMargin',
  'outsideMargin',
])

const DRAWING_RELATIVE_FROM_V: ReadonlySet<string> = new Set([
  'page',
  'margin',
  'paragraph',
  'line',
  'topMargin',
  'bottomMargin',
  'insideMargin',
  'outsideMargin',
])

const DRAWING_HORIZONTAL_ALIGN: ReadonlySet<string> = new Set([
  'left',
  'center',
  'right',
  'inside',
  'outside',
])

const DRAWING_VERTICAL_ALIGN: ReadonlySet<string> = new Set([
  'top',
  'center',
  'bottom',
  'inside',
  'outside',
])

function parseDrawingPositionH(element: OrderedXmlNode | undefined): DrawingPositionH | undefined {
  if (element === undefined) {
    return undefined
  }

  const relativeFromRaw = attr(element, 'relativeFrom')
  const relativeFrom: DrawingRelativeFromH = DRAWING_RELATIVE_FROM_H.has(relativeFromRaw ?? '')
    ? (relativeFromRaw as DrawingRelativeFromH)
    : 'page'
  const alignElement = child(element, 'wp:align')
  const offsetElement = child(element, 'wp:posOffset')
  const alignRaw = alignElement !== undefined ? textValue(alignElement) : undefined
  const align: DrawingHorizontalAlign | undefined = DRAWING_HORIZONTAL_ALIGN.has(alignRaw ?? '')
    ? (alignRaw as DrawingHorizontalAlign)
    : undefined
  const offsetEmu = offsetElement !== undefined ? parseInteger(textValue(offsetElement)) : undefined

  return {
    relativeFrom,
    ...(align !== undefined ? { align } : {}),
    ...(offsetEmu !== undefined ? { offsetEmu } : {}),
  }
}

function parseDrawingPositionV(element: OrderedXmlNode | undefined): DrawingPositionV | undefined {
  if (element === undefined) {
    return undefined
  }

  const relativeFromRaw = attr(element, 'relativeFrom')
  const relativeFrom: DrawingRelativeFromV = DRAWING_RELATIVE_FROM_V.has(relativeFromRaw ?? '')
    ? (relativeFromRaw as DrawingRelativeFromV)
    : 'page'
  const alignElement = child(element, 'wp:align')
  const offsetElement = child(element, 'wp:posOffset')
  const alignRaw = alignElement !== undefined ? textValue(alignElement) : undefined
  const align: DrawingVerticalAlign | undefined = DRAWING_VERTICAL_ALIGN.has(alignRaw ?? '')
    ? (alignRaw as DrawingVerticalAlign)
    : undefined
  const offsetEmu = offsetElement !== undefined ? parseInteger(textValue(offsetElement)) : undefined

  return {
    relativeFrom,
    ...(align !== undefined ? { align } : {}),
    ...(offsetEmu !== undefined ? { offsetEmu } : {}),
  }
}

const DRAWING_WRAP_ELEMENT_MODES: ReadonlyArray<readonly [string, DrawingWrapMode]> = [
  ['wp:wrapNone', 'none'],
  ['wp:wrapSquare', 'square'],
  ['wp:wrapTight', 'tight'],
  ['wp:wrapThrough', 'through'],
  ['wp:wrapTopAndBottom', 'topAndBottom'],
]

const DRAWING_WRAP_SIDES: ReadonlySet<string> = new Set(['bothSides', 'left', 'right', 'largest'])

/**
 * `wp:anchor`'s wrap choice is one required element out of five alternatives
 * (`wp:wrapNone`/`Square`/`Tight`/`Through`/`TopAndBottom`) — finds whichever
 * is present and normalizes it to one `DrawingWrap` shape.
 */
function parseDrawingWrap(anchorElement: OrderedXmlNode): DrawingWrap | undefined {
  for (const [elementName, mode] of DRAWING_WRAP_ELEMENT_MODES) {
    const wrapElement = child(anchorElement, elementName)
    if (wrapElement === undefined) {
      continue
    }

    const sideRaw = attr(wrapElement, 'wrapText')
    const side: DrawingWrapSide | undefined = DRAWING_WRAP_SIDES.has(sideRaw ?? '')
      ? (sideRaw as DrawingWrapSide)
      : undefined
    const distTEmu = parseInteger(attr(wrapElement, 'distT'))
    const distBEmu = parseInteger(attr(wrapElement, 'distB'))
    const distLEmu = parseInteger(attr(wrapElement, 'distL'))
    const distREmu = parseInteger(attr(wrapElement, 'distR'))

    return {
      mode,
      ...(side !== undefined ? { side } : {}),
      ...(distTEmu !== undefined ? { distTEmu } : {}),
      ...(distBEmu !== undefined ? { distBEmu } : {}),
      ...(distLEmu !== undefined ? { distLEmu } : {}),
      ...(distREmu !== undefined ? { distREmu } : {}),
    }
  }

  return undefined
}

/**
 * Captures `wp:anchor`'s direct children in original order: the Atlas-
 * modeled ones (`wp:extent`/`wp:effectExtent`/`wp:positionH`/`wp:positionV`/
 * the wrap choice/`wp:docPr`/`a:graphic`) become slot markers the serializer
 * rebuilds from live model state (DXP-09's position/wrap fields, parsed
 * above, join the pre-existing extent/docPr/graphic slots), and everything
 * else Atlas still doesn't model — `wp:simplePos`, `wp:cNvGraphicFramePr` —
 * is preserved as raw `UnknownNode` XML so a save emits schema-valid
 * `wp:anchor` output (DXS-03) instead of silently dropping it.
 */
function parseAnchorChildren(anchorElement: OrderedXmlNode): ReadonlyArray<DrawingAnchorChild> {
  const result: DrawingAnchorChild[] = []

  for (const entry of nodeChildren(anchorElement)) {
    if (isIgnorableText(entry)) {
      continue
    }

    switch (nodeName(entry)) {
      case 'wp:extent':
        result.push({ kind: 'anchor-slot', slot: 'extent' })
        break
      case 'wp:effectExtent':
        result.push({ kind: 'anchor-slot', slot: 'effectExtent' })
        break
      case 'wp:positionH':
        result.push({ kind: 'anchor-slot', slot: 'positionH' })
        break
      case 'wp:positionV':
        result.push({ kind: 'anchor-slot', slot: 'positionV' })
        break
      case 'wp:wrapNone':
      case 'wp:wrapSquare':
      case 'wp:wrapTight':
      case 'wp:wrapThrough':
      case 'wp:wrapTopAndBottom':
        result.push({ kind: 'anchor-slot', slot: 'wrap' })
        break
      case 'wp:docPr':
        result.push({ kind: 'anchor-slot', slot: 'docPr' })
        break
      case 'a:graphic':
        result.push({ kind: 'anchor-slot', slot: 'graphic' })
        break
      default:
        result.push(parseUnknownNode(entry))
        break
    }
  }

  return result
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

function parseDrawingEffectExtent(element: OrderedXmlNode): DrawingEffectExtent | undefined {
  const l = parseInteger(attr(element, 'l'))
  const t = parseInteger(attr(element, 't'))
  const r = parseInteger(attr(element, 'r'))
  const b = parseInteger(attr(element, 'b'))

  if (l === undefined || t === undefined || r === undefined || b === undefined) {
    return undefined
  }

  return { l, t, r, b }
}

function parseDrawingCrop(element: OrderedXmlNode | undefined): DrawingCrop | undefined {
  if (element === undefined) {
    return undefined
  }

  const l = parseInteger(attr(element, 'l'))
  const t = parseInteger(attr(element, 't'))
  const r = parseInteger(attr(element, 'r'))
  const b = parseInteger(attr(element, 'b'))

  return {
    ...(l !== undefined ? { l } : {}),
    ...(t !== undefined ? { t } : {}),
    ...(r !== undefined ? { r } : {}),
    ...(b !== undefined ? { b } : {}),
  }
}

function parseDrawingTransform(element: OrderedXmlNode | undefined): DrawingTransform | undefined {
  if (element === undefined) {
    return undefined
  }

  const rotation = parseInteger(attr(element, 'rot'))
  const flipH = parseOnOff(attr(element, 'flipH'))
  const flipV = parseOnOff(attr(element, 'flipV'))

  if (rotation === undefined && flipH === undefined && flipV === undefined) {
    // An `a:xfrm` present solely for its (unmodeled) `a:off`/`a:ext`
    // children carries no information Atlas's Drawing model represents —
    // treat it the same as no `a:xfrm` at all rather than emitting a
    // meaningless empty `transform: {}`.
    return undefined
  }

  return {
    ...(rotation !== undefined ? { rotation } : {}),
    ...(flipH !== undefined ? { flipH } : {}),
    ...(flipV !== undefined ? { flipV } : {}),
  }
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
  const tblGrid = parseTableGrid(child(element, 'w:tblGrid'))
  const rows: Array<TableRow | UnknownNode> = []

  for (const entry of expandWrapperNodes(nodeChildren(element))) {
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
    ...(tblGrid !== undefined ? { tblGrid } : {}),
    rows,
  }
}

function parseTableGrid(
  element: OrderedXmlNode | undefined,
): ReadonlyArray<ReturnType<typeof twip>> | undefined {
  if (element === undefined) {
    return undefined
  }

  const columns = children(element, 'w:gridCol').map(
    (column) => parseTwip(attr(column, 'w:w')) ?? twip(0),
  )

  return columns.length > 0 ? columns : undefined
}

function parseTableRow(element: OrderedXmlNode): TableRow {
  const props = parseTableRowProps(child(element, 'w:trPr'))
  const cells: TableRowChild[] = []

  for (const entry of expandWrapperNodes(nodeChildren(element))) {
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

  for (const entry of expandWrapperNodes(nodeChildren(element))) {
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

/**
 * DOCX-12 — every `w:rPr` child name `parseRunProps` below understands,
 * whether or not `RunProps` models it with a dedicated field (`w:cs` and
 * friends fall through to `rPrUnknown` passthrough same as a genuinely
 * invented tag would). `parseRPrUnknownChildren` consults this to decide
 * what counts as "known" — keep it in sync with the `child(element, ...)`
 * calls just below.
 */
const KNOWN_RPR_CHILD_NAMES: ReadonlySet<string> = new Set([
  'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps',
  'w:strike', 'w:dstrike', 'w:outline', 'w:emboss', 'w:imprint', 'w:vanish', 'w:webHidden',
  'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight',
  'w:u', 'w:bdr', 'w:shd', 'w:vertAlign', 'w:rtl', 'w:em', 'w:lang',
])

function parseRunProps(element: OrderedXmlNode | undefined): RunProps | undefined {
  if (element === undefined) {
    return undefined
  }

  const props: Mutable<RunProps> = {}
  const rStyle = attr(child(element, 'w:rStyle'), 'w:val')
  const bold = parseToggleElement(child(element, 'w:b'))
  const italic = parseToggleElement(child(element, 'w:i'))
  const boldCs = parseToggleElement(child(element, 'w:bCs'))
  const italicCs = parseToggleElement(child(element, 'w:iCs'))
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
  // DOCX-12 — a general character-effects/scaling round-trip pass: these
  // were parsed nowhere, so a run carrying any of them silently reverted on
  // save (outline/emboss/imprint are routine "text effects" formatting; `em`
  // — emphasis marks — is routine in CJK documents; `bdr`/`w` are rarer but
  // just as silently dropped previously).
  const outline = parseToggleElement(child(element, 'w:outline'))
  const emboss = parseToggleElement(child(element, 'w:emboss'))
  const imprint = parseToggleElement(child(element, 'w:imprint'))
  const em = parseEmphasisMark(attr(child(element, 'w:em'), 'w:val'))
  const bdr = parseBorder(child(element, 'w:bdr'))
  const charScale = parseCharScale(attr(child(element, 'w:w'), 'w:val'))
  const rPrUnknown = parseRPrUnknownChildren(element)

  if (rStyle !== undefined) props.rStyle = rStyle
  if (bold !== undefined) props.bold = bold
  if (italic !== undefined) props.italic = italic
  if (boldCs !== undefined) props.boldCs = boldCs
  if (italicCs !== undefined) props.italicCs = italicCs
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
  if (outline !== undefined) props.outline = outline
  if (emboss !== undefined) props.emboss = emboss
  if (imprint !== undefined) props.imprint = imprint
  if (em !== undefined) props.em = em
  if (bdr !== undefined) props.bdr = bdr
  if (charScale !== undefined) props.charScale = charScale
  if (rPrUnknown !== undefined) props.rPrUnknown = rPrUnknown

  return hasProps(props) ? props : undefined
}

function parseEmphasisMark(value: string | undefined): EmphasisMark | undefined {
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

/** `w:w`'s `w:val` (`ST_TextScale`) — Word always writes a plain integer percentage; a trailing `%` is tolerated defensively but not expected from real files. */
function parseCharScale(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }

  const normalized = value.endsWith('%') ? value.slice(0, -1) : value
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : undefined
}

/**
 * DOCX-12 — the general unknown-`w:rPr`-child passthrough. Walks `w:rPr`'s
 * actual source children in order and captures every one whose name isn't
 * in {@link KNOWN_RPR_CHILD_NAMES} via `parseUnknownNode` (byte-exact when
 * called through `parseDocument`, tree-rebuilt otherwise — see that
 * function's own doc comment), tagging each with the name of the nearest
 * *known* sibling that followed it in the source so
 * `buildRunPropertiesNode` can reinsert it at the same relative position
 * once it rebuilds the known children in schema order. Returns `undefined`
 * when `w:rPr` has no such children (the common case) rather than an empty
 * array, matching this file's usual "absent, not empty" convention.
 */
function parseRPrUnknownChildren(element: OrderedXmlNode): ReadonlyArray<RPrUnknownChild> | undefined {
  const entries = nodeChildren(element)
  let unknown: RPrUnknownChild[] | undefined

  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index]
    if (entry === undefined || isIgnorableText(entry)) {
      continue
    }

    const name = nodeName(entry)
    if (name === undefined || KNOWN_RPR_CHILD_NAMES.has(name)) {
      continue
    }

    let before: string | undefined
    for (let lookahead = index + 1; lookahead < entries.length; lookahead += 1) {
      const laterEntry = entries[lookahead]
      const laterName = laterEntry === undefined ? undefined : nodeName(laterEntry)
      if (laterName !== undefined && KNOWN_RPR_CHILD_NAMES.has(laterName)) {
        before = laterName
        break
      }
    }

    unknown ??= []
    unknown.push({
      xml: parseUnknownNode(entry).xml,
      ...(before !== undefined ? { before } : {}),
    })
  }

  return unknown
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
  const bidi = parseToggleElement(child(element, 'w:bidi'))
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
  if (bidi !== undefined) props.bidi = bidi
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
  // Round-trip fidelity audit (DXS round 2): `w:pgBorders` (a whole-page
  // border, or Word's watermark-adjacent "art" page frame) and `w:bidi`
  // (section reads right-to-left) were both previously unmodeled — parsed
  // nowhere, so silently dropped by a save that always rewrites
  // `word/document.xml` from this model.
  const pgBordersElement = child(element, 'w:pgBorders')
  const pgBorders = parseBorderSet(pgBordersElement)
  const pgBorderDisplay = parsePageBorderDisplay(attr(pgBordersElement, 'w:display'))
  const pgBorderOffsetFrom = parsePageBorderOffsetFrom(attr(pgBordersElement, 'w:offsetFrom'))
  const pgBorderZOrder = parsePageBorderZOrder(attr(pgBordersElement, 'w:zOrder'))
  const bidi = parseToggleElement(child(element, 'w:bidi'))
  const docGrid = parseDocGrid(child(element, 'w:docGrid'))
  const textDirection = attr(child(element, 'w:textDirection'), 'w:val')
  const rtlGutter = parseToggleElement(child(element, 'w:rtlGutter'))
  const formProt = parseToggleElement(child(element, 'w:formProt'))
  const noEndnote = parseToggleElement(child(element, 'w:noEndnote'))

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
  if (pgBorders !== undefined) props.pgBorders = pgBorders
  if (pgBorderDisplay !== undefined) props.pgBorderDisplay = pgBorderDisplay
  if (pgBorderOffsetFrom !== undefined) props.pgBorderOffsetFrom = pgBorderOffsetFrom
  if (pgBorderZOrder !== undefined) props.pgBorderZOrder = pgBorderZOrder
  if (bidi !== undefined) props.bidi = bidi
  if (docGrid !== undefined) props.docGrid = docGrid
  if (textDirection !== undefined) props.textDirection = textDirection
  if (rtlGutter !== undefined) props.rtlGutter = rtlGutter
  if (formProt !== undefined) props.formProt = formProt
  if (noEndnote !== undefined) props.noEndnote = noEndnote

  return hasProps(props) ? props : undefined
}

function parseDocGrid(element: OrderedXmlNode | undefined): DocGrid | undefined {
  if (element === undefined) {
    return undefined
  }

  const type = attr(element, 'w:type')
  const linePitch = parseInteger(attr(element, 'w:linePitch'))
  const charSpace = parseInteger(attr(element, 'w:charSpace'))

  const docGrid: Mutable<DocGrid> = {}
  if (type !== undefined) docGrid.type = type
  if (linePitch !== undefined) docGrid.linePitch = linePitch
  if (charSpace !== undefined) docGrid.charSpace = charSpace

  return hasProps(docGrid) ? docGrid : undefined
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

/**
 * D19 / DXS-16 — reconstructing an unsupported element from the parsed tree
 * (`xmlBuilder.build([element])`) only re-emits attributes that element
 * itself (or one of its own descendants) carries. A namespace prefix the
 * element or a descendant USES but that a non-root ANCESTOR declared — most
 * commonly `w:sdt`/`mc:AlternateContent`/`mc:Choice` unwrapping (see
 * `expandWrapperNodes`) discarding a wrapper that declared the extension
 * namespace its unwrapped content still refers to — never appears anywhere
 * in the rebuild, producing schema-invalid XML (an undeclared prefix) on
 * save. Slicing the exact source substring instead fixes formatting fidelity
 * (attribute order, self-closing style, entity spelling, etc.) but not this
 * specific gap by itself, since the substring doesn't include the ancestor's
 * declaration either — so `sourceRangeContext` additionally tracks, per
 * node, which namespaces a non-root ancestor made available, and
 * `injectMissingNamespaces` re-declares locally whichever of those the
 * captured fragment actually references and doesn't already redeclare
 * itself. A prefix declared on the document root itself is deliberately
 * excluded from this tracking: `rootNamespaces`/`mcIgnorable` (DXS-15)
 * already re-emit the source root's full namespace set unconditionally, so
 * every root-declared prefix is already in scope everywhere in the output
 * without any per-node help, and re-injecting it here on top would just be
 * redundant (and break byte-identical round-trip for the overwhelmingly
 * common case where an unknown node only ever uses namespaces the root
 * itself already declares, e.g. plain `w:*`).
 */
function parseUnknownNode(element: OrderedXmlNode): UnknownNode {
  const info = sourceRangeContext?.info.get(element)
  if (info === undefined) {
    // No raw range available (lookup miss — e.g. the source tag couldn't be
    // relocated in the text, or this ran outside a `parseDocument` call, as
    // every unit test that calls the file's other parse helpers directly
    // does): fall back to the previous tree-rebuild behavior rather than
    // producing no output at all.
    return {
      kind: 'unknown',
      xml: xmlBuilder.build([element]),
    }
  }

  const rawXml = sourceRangeContext!.xml.slice(info.start, info.end)
  return {
    kind: 'unknown',
    xml: injectMissingNamespaces(rawXml, info.inheritedNamespaces),
  }
}

// ---------------------------------------------------------------------------
// D19 / DXS-16 — raw source-substring range tracking for unknown nodes
// ---------------------------------------------------------------------------

interface RawNodeInfo {
  readonly start: number
  readonly end: number
  /** xmlns:prefix -> uri available from a non-root ancestor at this node's position. */
  readonly inheritedNamespaces: ReadonlyMap<string, string>
}

interface SourceRangeContext {
  readonly xml: string
  readonly info: WeakMap<OrderedXmlNode, RawNodeInfo>
}

const EMPTY_NAMESPACE_SCOPE: ReadonlyMap<string, string> = new Map()

/**
 * Builds the raw-range/namespace-scope index for one `parseDocument` call.
 * Walks `raw` (the exact tree `xmlParser.parse(xml)` produced) depth-first
 * in document order, alongside a forward-only cursor into `xml` — the same
 * order the semantic parse functions below (`parseBody`, `parseParagraph`,
 * `parseRun`, ...) independently traverse this same tree in, so a node's
 * entry lands in `info` before anything in this file ever needs to look it
 * up. A tag that can't be relocated (should not happen for well-formed XML
 * fast-xml-parser itself just parsed, but never trusted to be impossible)
 * simply gets no entry — `parseUnknownNode` degrades to tree-rebuild for
 * that one node rather than the whole parse failing.
 */
function computeSourceRangeContext(
  xml: string,
  raw: ReadonlyArray<OrderedXmlNode>,
): SourceRangeContext {
  const info = new WeakMap<OrderedXmlNode, RawNodeInfo>()
  // The root level's own namespace declarations are intentionally excluded
  // from what gets tracked as "inherited" (see `parseUnknownNode`'s doc
  // comment) — `mergeOwnNamespaces: false` only for this outermost call.
  computeRawNodeInfo(raw, xml, 0, EMPTY_NAMESPACE_SCOPE, info, false)
  return { xml, info }
}

function computeRawNodeInfo(
  nodes: ReadonlyArray<OrderedXmlNode>,
  xml: string,
  cursor: number,
  inheritedScope: ReadonlyMap<string, string>,
  info: WeakMap<OrderedXmlNode, RawNodeInfo>,
  mergeOwnNamespaces: boolean,
): number {
  let position = cursor

  for (const node of nodes) {
    const name = nodeName(node)
    // Text/comment/processing-instruction pseudo-nodes ('#text', '#comment',
    // '?xml', ...) have no tag to relocate and are never individually passed
    // to `parseUnknownNode` — skip without advancing the cursor; the next
    // real element's own forward search naturally skips over them.
    if (name === undefined || name.startsWith('#') || name.startsWith('?')) {
      continue
    }

    const tagStart = findRawTagStart(xml, position, name)
    if (tagStart === -1) {
      continue
    }

    const openTag = findRawTagOpenEnd(xml, tagStart)
    if (openTag === undefined) {
      continue
    }

    if (openTag.selfClosing) {
      info.set(node, { start: tagStart, end: openTag.end, inheritedNamespaces: inheritedScope })
      position = openTag.end
      continue
    }

    const childScope = mergeOwnNamespaces
      ? mergeNamespaceScopes(inheritedScope, collectLocalNamespaces(node))
      : inheritedScope

    const childrenEnd = computeRawNodeInfo(nodeChildren(node), xml, openTag.end, childScope, info, true)
    const closeEnd = findRawTagCloseEnd(xml, childrenEnd, name)
    if (closeEnd === undefined) {
      position = childrenEnd
      continue
    }

    info.set(node, { start: tagStart, end: closeEnd, inheritedNamespaces: inheritedScope })
    position = closeEnd
  }

  return position
}

/** Finds the next `<name` at or after `fromIndex` whose name isn't a longer identifier's prefix (e.g. searching `w:p` must not match `w:pPr`). */
function findRawTagStart(xml: string, fromIndex: number, name: string): number {
  const needle = `<${name}`
  let index = fromIndex

  for (;;) {
    index = xml.indexOf(needle, index)
    if (index === -1) {
      return -1
    }

    const after = xml.charAt(index + needle.length)
    if (after === '' || /[\s/>]/.test(after)) {
      return index
    }

    index += needle.length
  }
}

/** From a tag's `<` at `tagStart`, finds the unquoted `>` that closes the start tag. */
function findRawTagOpenEnd(
  xml: string,
  tagStart: number,
): { readonly end: number; readonly selfClosing: boolean } | undefined {
  let quote: string | undefined
  for (let i = tagStart; i < xml.length; i += 1) {
    const c = xml.charAt(i)
    if (quote !== undefined) {
      if (c === quote) {
        quote = undefined
      }
      continue
    }
    if (c === '"' || c === "'") {
      quote = c
      continue
    }
    if (c === '>') {
      return { end: i + 1, selfClosing: xml.charAt(i - 1) === '/' }
    }
  }
  return undefined
}

function findRawTagCloseEnd(xml: string, fromIndex: number, name: string): number | undefined {
  const needle = `</${name}>`
  const index = xml.indexOf(needle, fromIndex)
  return index === -1 ? undefined : index + needle.length
}

const XMLNS_ATTR_PREFIX = '@_xmlns:'

function collectLocalNamespaces(node: OrderedXmlNode): ReadonlyMap<string, string> {
  const attributes = node[':@']
  if (attributes === undefined) {
    return EMPTY_NAMESPACE_SCOPE
  }

  let namespaces: Map<string, string> | undefined
  for (const [key, value] of Object.entries(attributes)) {
    if (key.startsWith(XMLNS_ATTR_PREFIX) && value !== undefined) {
      namespaces ??= new Map()
      namespaces.set(key.slice(XMLNS_ATTR_PREFIX.length), value)
    }
  }

  return namespaces ?? EMPTY_NAMESPACE_SCOPE
}

function mergeNamespaceScopes(
  base: ReadonlyMap<string, string>,
  overrides: ReadonlyMap<string, string>,
): ReadonlyMap<string, string> {
  if (overrides.size === 0) {
    return base
  }
  return new Map([...base, ...overrides])
}

/**
 * Referenced-prefix detection is a plain substring scan across the whole
 * fragment (tag names, attribute names, AND ordinary text content) rather
 * than a full tokenizer: a false positive — plain text that happens to
 * contain a `word:word` pattern (e.g. "Ratio a:b") — only ever causes a
 * harmless redundant `xmlns:` re-declaration (still valid XML), never a
 * missed one, so erring toward over-matching here is the safe direction.
 */
const PREFIX_REFERENCE_RE = /[<\s/]([A-Za-z_][\w.-]*):[A-Za-z_]/g
const NAMESPACE_DECLARATION_RE = /\bxmlns:([A-Za-z_][\w.-]*)\s*=/g

function collectMatches(xml: string, pattern: RegExp): ReadonlySet<string> {
  const matches = new Set<string>()
  pattern.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = pattern.exec(xml)) !== null) {
    matches.add(match[1])
  }
  return matches
}

/**
 * Re-declares, on `fragment`'s own outermost tag, whichever namespace
 * prefixes it references but doesn't already declare itself, out of the
 * ones `inheritedNamespaces` says a (non-root) ancestor made available —
 * see `parseUnknownNode`'s doc comment for why this is needed at all.
 */
function injectMissingNamespaces(
  fragment: string,
  inheritedNamespaces: ReadonlyMap<string, string>,
): string {
  if (inheritedNamespaces.size === 0) {
    return fragment
  }

  const referenced = collectMatches(fragment, PREFIX_REFERENCE_RE)
  const declared = collectMatches(fragment, NAMESPACE_DECLARATION_RE)

  const missing: Array<readonly [string, string]> = []
  for (const prefix of referenced) {
    if (declared.has(prefix)) {
      continue
    }
    const uri = inheritedNamespaces.get(prefix)
    if (uri !== undefined) {
      missing.push([prefix, uri])
    }
  }

  if (missing.length === 0) {
    return fragment
  }

  const openTag = findRawTagOpenEnd(fragment, 0)
  if (openTag === undefined) {
    return fragment
  }

  const insertPos = openTag.selfClosing ? openTag.end - 2 : openTag.end - 1
  const declarations = missing
    .map(([prefix, uri]) => ` xmlns:${prefix}="${escapeXmlAttributeValue(uri)}"`)
    .join('')

  return fragment.slice(0, insertPos) + declarations + fragment.slice(insertPos)
}

function escapeXmlAttributeValue(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
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

function parsePageBorderDisplay(value: string | undefined): SectionProps['pgBorderDisplay'] | undefined {
  switch (value) {
    case 'allPages':
    case 'firstPage':
    case 'notFirstPage':
      return value
    default:
      return undefined
  }
}

function parsePageBorderOffsetFrom(value: string | undefined): SectionProps['pgBorderOffsetFrom'] | undefined {
  switch (value) {
    case 'page':
    case 'text':
      return value
    default:
      return undefined
  }
}

function parsePageBorderZOrder(value: string | undefined): SectionProps['pgBorderZOrder'] | undefined {
  switch (value) {
    case 'front':
    case 'back':
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
