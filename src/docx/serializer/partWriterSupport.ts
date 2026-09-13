import { XMLBuilder, XMLParser } from 'fast-xml-parser'

import type { Block, Paragraph } from '../model'
import { buildParagraph } from './documentWriter'

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

export function buildParagraphBlockNodes(blocks: ReadonlyArray<Block>): ReadonlyArray<OrderedXmlNode> {
  return blocks.map((block) => {
    if (block.kind !== 'paragraph') {
      throw new Error(
        `Only paragraph blocks can be serialized in standalone DOCX parts; received ${block.kind}.`,
      )
    }

    return normalizeParagraphNode(block)
  })
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
  const paragraphNode = buildParagraph(paragraph) as unknown

  if (isOrderedParagraphNode(paragraphNode)) {
    return paragraphNode
  }

  if (isWrappedObjectParagraphNode(paragraphNode)) {
    return parseOrderedNode(fragmentXmlBuilder.build(paragraphNode))
  }

  if (isObjectXmlNode(paragraphNode)) {
    return parseOrderedNode(fragmentXmlBuilder.build({ 'w:p': paragraphNode }))
  }

  throw new Error('buildParagraph returned an unsupported XML fragment shape.')
}

function parseOrderedNode(xml: string): OrderedXmlNode {
  const parsed = orderedXmlParser.parse(xml) as OrderedXmlNode[]

  if (parsed.length !== 1 || !isOrderedXmlNode(parsed[0])) {
    throw new Error('Failed to normalize paragraph XML for DOCX part serialization.')
  }

  return parsed[0]
}

function isOrderedParagraphNode(value: unknown): value is OrderedXmlNode {
  return isOrderedXmlNode(value) && Array.isArray(value['w:p'])
}

function isWrappedObjectParagraphNode(value: unknown): value is ObjectXmlNode {
  return isObjectXmlNode(value) && 'w:p' in value
}

function isOrderedXmlNode(value: unknown): value is OrderedXmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isObjectXmlNode(value: unknown): value is ObjectXmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
