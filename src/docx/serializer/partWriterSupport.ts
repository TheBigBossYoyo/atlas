import { XMLBuilder, XMLParser } from 'fast-xml-parser'

import { assertNever, type Block, type Paragraph, type Table } from '../model'
import { buildParagraph, buildTable } from './documentWriter'

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
 * `UnknownNode` blocks are not supported at this level: the raw-XML
 * placeholder/restoration mechanism `documentWriter.ts` uses for unknown
 * *document*-body content requires a final restoration pass over the whole
 * generated XML string, which these standalone parts don't currently have
 * a hook for. This is an existing limitation carried forward, not a new
 * one — every block kind other than paragraph already failed to serialize
 * here before this fix; only `table` support is new.
 */
export function buildBlockNodes(blocks: ReadonlyArray<Block>): ReadonlyArray<OrderedXmlNode> {
  return blocks.map((block) => normalizeBlockNode(block))
}

function normalizeBlockNode(block: Block): OrderedXmlNode {
  switch (block.kind) {
    case 'paragraph':
      return normalizeParagraphNode(block)
    case 'table':
      return normalizeTableNode(block)
    case 'unknown':
      throw new Error(
        'Cannot serialize an unrecognized (UnknownNode) block inside a standalone DOCX part '
          + '(header/footer/footnote/endnote/comment) — only paragraph and table blocks are supported here.',
      )
    default:
      return assertNever(block)
  }
}

export function serializeWordPart(
  rootName: string,
  children: ReadonlyArray<OrderedXmlNode>,
  includeRelationshipsNamespace: boolean = false,
): string {
  const root: OrderedXmlNode = {
    [rootName]: [...children],
    ':@': {
      '@_xmlns:w': WORD_NAMESPACE,
      ...(includeRelationshipsNamespace ? { '@_xmlns:r': RELATIONSHIP_NAMESPACE } : {}),
    },
  }

  return `${XML_DECLARATION}${orderedXmlBuilder.build([root])}`
}

function normalizeParagraphNode(paragraph: Paragraph): OrderedXmlNode {
  return normalizeBuiltNode(buildParagraph(paragraph), 'w:p', 'buildParagraph')
}

function normalizeTableNode(table: Table): OrderedXmlNode {
  return normalizeBuiltNode(buildTable(table), 'w:tbl', 'buildTable')
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
