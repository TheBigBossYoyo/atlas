import { XMLParser } from 'fast-xml-parser'

import { parseDocument } from './document'
import { DocxParseError } from './unzip'
import type { Comment, Paragraph } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
//
// NOTE: this parser is deliberately NOT `preserveOrder`/`trimValues: false`
// like the main document parser (see parser/document.ts). It is only used
// here to read scalar `w:comment` attributes (id/author/date/…), which are
// unaffected by node ordering or whitespace trimming. Comment BODIES are
// re-parsed from the original XML text (see `extractCommentBodyXmlById`
// below) specifically to avoid round-tripping paragraph content through
// this lossy, non-order-preserving representation — doing so previously
// silently reordered interleaved run/hyperlink siblings and trimmed
// significant leading/trailing whitespace out of `w:t` runs.
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WORD_2012_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml'
const COMMENT_OPEN_TAG_RE = /<w:comment\b([^>]*?)(\/?)>/g
const COMMENT_CLOSE_TAG = '</w:comment>'
const COMMENT_ID_ATTR_RE = /w:id\s*=\s*"([^"]*)"|w:id\s*=\s*'([^']*)'/

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
  const bodyXmlById = extractCommentBodyXmlById(xml)
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
    const body = parseCommentBody(bodyXmlById.get(id))

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

/**
 * Scan the ORIGINAL comments.xml text for each `<w:comment>` element and
 * return a map of comment id → the raw inner XML between its opening and
 * closing tags (verbatim, not re-serialized).
 *
 * This runs against the source text rather than the `xmlParser` output
 * because that parser is not `preserveOrder`/`trimValues: false`: reading
 * paragraph content back out of its parsed object form and rebuilding XML
 * from it (the previous approach) silently reordered interleaved
 * same-tag/different-tag siblings (e.g. run, hyperlink, run → run, run,
 * hyperlink) and trimmed significant leading/trailing whitespace out of
 * `w:t` runs adjacent to another element. Slicing the untouched source text
 * avoids both problems.
 */
function extractCommentBodyXmlById(xml: string): ReadonlyMap<string, string> {
  const bodies = new Map<string, string>()
  COMMENT_OPEN_TAG_RE.lastIndex = 0
  let openTag: RegExpExecArray | null

  while ((openTag = COMMENT_OPEN_TAG_RE.exec(xml)) !== null) {
    const attrs = openTag[1]
    const isSelfClosing = openTag[2] === '/'
    const idMatch = COMMENT_ID_ATTR_RE.exec(attrs)
    const id = idMatch !== null ? (idMatch[1] ?? idMatch[2]) : undefined

    if (isSelfClosing) {
      if (id !== undefined) {
        bodies.set(id, '')
      }
      continue
    }

    const searchStart = COMMENT_OPEN_TAG_RE.lastIndex
    const closeIndex = xml.indexOf(COMMENT_CLOSE_TAG, searchStart)
    if (closeIndex === -1) {
      // Malformed/truncated XML — the outer `xmlParser.parse` above will
      // already have thrown for this before we get here in practice.
      break
    }

    if (id !== undefined) {
      bodies.set(id, xml.slice(searchStart, closeIndex))
    }

    COMMENT_OPEN_TAG_RE.lastIndex = closeIndex + COMMENT_CLOSE_TAG.length
  }

  return bodies
}

function parseCommentBody(innerXml: string | undefined): ReadonlyArray<Paragraph> {
  if (innerXml === undefined || innerXml.trim().length === 0) {
    return []
  }

  const syntheticXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w15="${WORD_2012_NAMESPACE}">` +
    `<w:body>${innerXml}</w:body></w:document>`

  const document = parseDocument(syntheticXml)
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
