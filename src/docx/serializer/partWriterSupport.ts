import { XMLBuilder, XMLParser } from 'fast-xml-parser'

import { assertNever, type Block, type Paragraph, type Table } from '../model'
import {
  buildNamespaceDeclarationAttributes,
  buildParagraphWithState,
  buildTableWithState,
  restoreUnknownXml,
  STANDARD_NAMESPACE_URIS,
  type SerializeState,
} from './documentWriter'

export { createSerializeState, type SerializeState } from './documentWriter'

type ObjectXmlPrimitive = string | number | boolean
type ObjectXmlValue = ObjectXmlPrimitive | ObjectXmlNode | ObjectXmlValue[]

interface ObjectXmlNode {
  readonly [key: string]: ObjectXmlValue | undefined
}

interface OrderedXmlAttributes {
  readonly [key: string]: string | undefined
}

export interface OrderedXmlNode {
  readonly [key: string]: OrderedXmlNode[] | OrderedXmlAttributes | string | undefined
}

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
export const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
export const RELATIONSHIP_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const orderedXmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

const fragmentXmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

const orderedXmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
})

/**
 * Serializes the block content of a standalone DOCX part (header, footer,
 * footnote, endnote, or comment body) — paragraphs and tables alike (wave 1
 * follow-up: these parts previously supported only paragraph blocks, so a
 * table inside one — e.g. a letterhead, or a table-formatted footnote —
 * made `saveDocx` throw instead of round-tripping it).
 *
 * `UnknownNode` *blocks* (a top-level item in `blocks` itself being of kind
 * `'unknown'`) are not supported: the schema for these wrapper elements
 * (`w:hdr`/`w:ftr`/`w:footnote`/`w:endnote`/`w:comment`) only ever contains
 * paragraphs and tables directly, so a top-level unknown block would mean
 * something has gone wrong upstream. A node type this serializer doesn't
 * model *nested inside* a paragraph or table (e.g. `w:proofErr`, ubiquitous
 * in real Word documents) is fully supported: `state` must be threaded
 * through to {@link serializeWordPart} (or restored directly via
 * `restoreUnknownXml`, for a caller that builds its own root XML) so the
 * `atlas-raw-unknown` placeholders `buildParagraphWithState`/
 * `buildTableWithState` emit for such content get substituted back to the
 * real raw XML — omitting that step leaves the literal placeholder tag in
 * the saved part, an XML well-formedness violation.
 */
export function buildBlockNodes(blocks: ReadonlyArray<Block>, state: SerializeState): ReadonlyArray<OrderedXmlNode> {
  return blocks.map((block) => normalizeBlockNode(block, state))
}

function normalizeBlockNode(block: Block, state: SerializeState): OrderedXmlNode {
  switch (block.kind) {
    case 'paragraph':
      return normalizeParagraphNode(block, state)
    case 'table':
      return normalizeTableNode(block, state)
    case 'unknown':
      throw new Error(
        'Cannot serialize an unrecognized (UnknownNode) block inside a standalone DOCX part '
          + '(header/footer/footnote/endnote/comment) — only paragraph and table blocks are supported here.',
      )
    default:
      return assertNever(block)
  }
}

/**
 * D19 / DXS-08: standalone parts (header/footer/footnote/endnote) previously
 * declared only `xmlns:w` (+ `xmlns:r` when the caller opted in), so a
 * drawing/revision/shape/etc. inside one of those parts emitted an
 * undeclared namespace prefix — an XML well-formedness violation Word may
 * reject or "repair" on open. Declares the same full standard namespace set
 * `documentWriter.ts` declares on the document root instead, so any content
 * `buildBlockNodes` can produce (which reuses `documentWriter.ts`'s own
 * paragraph/table builders) always has its namespaces in scope.
 *
 * `state`, when given, is the same {@link SerializeState} passed to
 * `buildBlockNodes` for these `children` — required to substitute back any
 * `atlas-raw-unknown` placeholder a nested unrecognized node produced (see
 * `buildBlockNodes`'s doc comment). Omit only for a root with no block
 * content at all (nothing to restore).
 */
export function serializeWordPart(
  rootName: string,
  children: ReadonlyArray<OrderedXmlNode>,
  state?: SerializeState,
): string {
  const root: OrderedXmlNode = {
    [rootName]: [...children],
    ':@': buildNamespaceDeclarationAttributes(STANDARD_NAMESPACE_URIS),
  }

  const xml = `${XML_DECLARATION}${orderedXmlBuilder.build([root])}`
  return state !== undefined ? restoreUnknownXml(xml, state) : xml
}

function normalizeParagraphNode(paragraph: Paragraph, state: SerializeState): OrderedXmlNode {
  return normalizeBuiltNode(buildParagraphWithState(paragraph, state), 'w:p', 'buildParagraph')
}

function normalizeTableNode(table: Table, state: SerializeState): OrderedXmlNode {
  return normalizeBuiltNode(buildTableWithState(table, state), 'w:tbl', 'buildTable')
}

/**
 * Normalizes whatever shape a documentWriter.ts `build*` helper returned
 * into the `preserveOrder`-style `OrderedXmlNode` this part writer's own
 * builder expects. In practice `build*` already returns that shape
 * directly (the fast paths below), but the fallback paths handle a plain
 * object-keyed shape defensively without assuming which one a given
 * builder function produces.
 */
function normalizeBuiltNode(built: unknown, tagName: string, builderName: string): OrderedXmlNode {
  if (isOrderedNodeForTag(built, tagName)) {
    return built
  }

  if (isWrappedObjectNodeForTag(built, tagName)) {
    return parseOrderedNode(fragmentXmlBuilder.build(built), builderName)
  }

  if (isObjectXmlNode(built)) {
    return parseOrderedNode(fragmentXmlBuilder.build({ [tagName]: built }), builderName)
  }

  throw new Error(`${builderName} returned an unsupported XML fragment shape.`)
}

function parseOrderedNode(xml: string, builderName: string): OrderedXmlNode {
  const parsed = orderedXmlParser.parse(xml) as OrderedXmlNode[]

  if (parsed.length !== 1 || !isOrderedXmlNode(parsed[0])) {
    throw new Error(`Failed to normalize ${builderName} output for DOCX part serialization.`)
  }

  return parsed[0]
}

function isOrderedNodeForTag(value: unknown, tagName: string): value is OrderedXmlNode {
  return isOrderedXmlNode(value) && Array.isArray(value[tagName])
}

function isWrappedObjectNodeForTag(value: unknown, tagName: string): value is ObjectXmlNode {
  return isObjectXmlNode(value) && tagName in value
}

function isOrderedXmlNode(value: unknown): value is OrderedXmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isObjectXmlNode(value: unknown): value is ObjectXmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
