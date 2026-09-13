import { XMLParser } from 'fast-xml-parser'

import { extractElementInnerXmlsById, parseBlocksFromXmlFragment } from './partBody'
import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import type { Block, Comment } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
//
// NOTE: this parser is deliberately NOT `preserveOrder`/`trimValues: false`
// like the main document parser (see parser/document.ts). It is only used
// here to read scalar `w:comment` attributes (id/author/date/…), which are
// unaffected by node ordering or whitespace trimming. Comment BODIES are
// re-parsed from the original XML text (via `partBody.ts`'s
// `extractElementInnerXmlsById`, this module's own proven technique lifted
// into a shared helper) specifically to avoid round-tripping paragraph
// content through this lossy, non-order-preserving representation — doing
// so previously silently reordered interleaved run/hyperlink siblings and
// trimmed significant leading/trailing whitespace out of `w:t` runs.
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

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
  const bodyXmlById = extractElementInnerXmlsById(xml, 'w:comment')
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
 * Parse a comment's raw inner XML (extracted verbatim from the original
 * source text by `extractElementInnerXmlsById` above — never round-tripped
 * through the non-order-preserving `xmlParser`) into `Block[]`.
 *
 * Previously restricted its result to `block.kind === 'paragraph'` (a wave 1
 * follow-up fix): a comment containing a table — legitimate OOXML content —
 * had that table silently discarded even though `parseBlocksFromXmlFragment`
 * parsed it correctly, because this function threw it away afterwards.
 */
function parseCommentBody(innerXml: string | undefined): ReadonlyArray<Block> {
  return parseBlocksFromXmlFragment(innerXml)
}
