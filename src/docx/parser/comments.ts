import { XMLBuilder, XMLParser } from 'fast-xml-parser'

import { parseDocument } from './document'
import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import type { Comment, Paragraph } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WORD_2012_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml'

// ---------------------------------------------------------------------------
// Internal XML shapes
// ---------------------------------------------------------------------------

interface RawComment {
  '@_w:id'?: string | number
  '@_w:author'?: string
  '@_w:date'?: string
  '@_w:initials'?: string
  '@_w:parentId'?: string | number
  '@_w15:parentId'?: string | number
  'w:p'?: unknown | unknown[]
  [key: string]: unknown
}

interface RawCommentsDoc {
  'w:comments'?: {
    'w:comment'?: RawComment | RawComment[]
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse the XML content of `word/comments.xml`.
 *
 * @param xml - UTF-8 text of the comments part.
 * @returns   Immutable map: comment id → Comment node.
 * @throws    DocxParseError on malformed XML.
 */
export function parseComments(xml: string): ReadonlyMap<string, Comment> {
  assertXmlPartSizeWithinLimit(xml, 'word/comments.xml')
  let parsed: RawCommentsDoc
  try {
    parsed = xmlParser.parse(xml) as RawCommentsDoc
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse comments XML: ${msg}`)
  }

  const rawComments = parsed?.['w:comments']?.['w:comment']
  if (rawComments === undefined || rawComments === null) {
    return new Map()
  }

  const items: RawComment[] = Array.isArray(rawComments) ? rawComments : [rawComments]
  const result = new Map<string, Comment>()

  for (const item of items) {
    const rawId = item['@_w:id']
    if (rawId === undefined || rawId === null) {
      throw new DocxParseError('Comment element is missing required w:id attribute')
    }
    const id = String(rawId)

    // Extract optional metadata; omit keys whose value is absent so the model
    // fields remain `undefined` rather than the empty string.
    const author = item['@_w:author'] !== undefined ? String(item['@_w:author']) : undefined
    const date = item['@_w:date'] !== undefined ? String(item['@_w:date']) : undefined
    const initials = item['@_w:initials'] !== undefined ? String(item['@_w:initials']) : undefined
    const parentIdValue = item['@_w15:parentId'] ?? item['@_w:parentId']
    const parentId = parentIdValue !== undefined ? String(parentIdValue) : undefined
    const body = parseCommentBody(item)

    const comment: Comment = {
      kind: 'comment',
      id,
      ...(author !== undefined ? { author } : {}),
      ...(date !== undefined ? { date } : {}),
      ...(initials !== undefined ? { initials } : {}),
      ...(parentId !== undefined ? { parentId } : {}),
      body,
    }
    result.set(id, comment)
  }

  return result
}

function parseCommentBody(comment: RawComment): ReadonlyArray<Paragraph> {
  const rawParagraphs = comment['w:p']
  if (rawParagraphs === undefined) {
    return []
  }

  const paragraphNodes = Array.isArray(rawParagraphs) ? rawParagraphs : [rawParagraphs]
  if (paragraphNodes.length === 0) {
    return []
  }

  const syntheticXml = xmlBuilder.build({
    'w:document': {
      '@_xmlns:w': WORD_NAMESPACE,
      '@_xmlns:w15': WORD_2012_NAMESPACE,
      'w:body': {
        'w:p': paragraphNodes,
      },
    },
  })

  const document = parseDocument(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${syntheticXml}`)
  const paragraphs: Paragraph[] = []

  for (const section of document.sections) {
    for (const block of section.blocks) {
      if (block.kind === 'paragraph') {
        paragraphs.push(block)
      }
    }
  }

  return paragraphs
}
