/**
 * Atlas — DOCX commentsExtended.xml parser (D16 / DXS-11)
 *
 * `word/commentsExtended.xml` is where Word records comment-thread metadata
 * that doesn't live on `<w:comment>` itself — most importantly whether a
 * thread has been marked resolved (`w15:done`). Each `<w15:commentEx>` is
 * keyed by `w15:paraId`, which matches the `w14:paraId` attribute Word
 * stamps onto the comment's own first body paragraph in `comments.xml`
 * (captured by `parser/document.ts` as `Paragraph.paraId`, see D19/DXS-10).
 *
 * This module only reads the done/resolved flag; `w15:paraIdParent` (used
 * for reply-thread nesting) is intentionally not modeled — Atlas already
 * tracks parent/child comment relationships via `w15:parentId` on
 * `<w:comment>` itself (see `comments.ts`).
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

interface RawCommentEx {
  readonly '@_w15:paraId'?: string | number
  readonly '@_w15:done'?: string | number
  readonly [key: string]: unknown
}

interface RawCommentsExtendedDoc {
  readonly 'w15:commentsEx'?: {
    readonly 'w15:commentEx'?: RawCommentEx | ReadonlyArray<RawCommentEx>
  }
}

function isDone(value: string | number | undefined): boolean {
  return value === '1' || value === 1 || value === 'true'
}

/**
 * Parse the XML content of `word/commentsExtended.xml`.
 *
 * @param xml - UTF-8 text of the commentsExtended part.
 * @returns   Immutable map: `w15:paraId` -> whether that comment thread is
 *            marked resolved/done. Entries with no `w15:paraId` are skipped
 *            (there is nothing to key them by).
 * @throws    DocxParseError on malformed XML.
 */
export function parseCommentsExtended(xml: string): ReadonlyMap<string, boolean> {
  assertXmlPartSizeWithinLimit(xml, 'word/commentsExtended.xml')
  let parsed: RawCommentsExtendedDoc
  try {
    parsed = xmlParser.parse(xml) as RawCommentsExtendedDoc
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse commentsExtended XML: ${msg}`)
  }

  const raw = parsed?.['w15:commentsEx']?.['w15:commentEx']
  if (raw === undefined || raw === null) {
    return new Map()
  }

  const items: ReadonlyArray<RawCommentEx> = Array.isArray(raw) ? raw : [raw]
  const result = new Map<string, boolean>()

  for (const item of items) {
    const paraId = item['@_w15:paraId']
    if (paraId === undefined || paraId === null || paraId === '') {
      continue
    }
    result.set(String(paraId), isDone(item['@_w15:done']))
  }

  return result
}
