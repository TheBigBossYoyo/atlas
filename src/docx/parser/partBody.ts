/**
 * Atlas — shared "part body" parsing helpers
 *
 * Comments, headers, footers, footnotes, and endnotes all hold ordinary
 * block content (paragraphs, and — the point of this module's rewrite,
 * wave 1 follow-up — tables too) nested inside their own part-specific
 * wrapper element (`w:comment`, `w:hdr`, `w:ftr`, `w:footnote`,
 * `w:endnote`). None of those wrappers has a `w:body` child, so none of
 * them can be handed directly to `parseDocument`, which only knows how to
 * find `w:document`/`w:body`.
 *
 * `comments.ts` solved this for comment bodies by re-rooting the extracted
 * inner XML under a synthetic `<w:document><w:body>` element and running it
 * back through `parseDocument`. This module lifts that same technique into
 * shared helpers so headers/footers/footnotes/endnotes get it too, and
 * generalizes it to keep `w:tbl` blocks — a header/footer built around a
 * letterhead table, or a footnote/endnote containing a table, previously
 * had that table silently discarded because the caller only ever asked for
 * `w:p` children (Wave A.5's "Option B" stub — DXP-03/DXP-04 — narrowed the
 * gap to "parses but drops tables" rather than "doesn't parse at all", but
 * never closed it).
 *
 * Extracting the exact source substring for each wrapper (rather than
 * re-serializing an already-parsed representation) also preserves the
 * original document order of interleaved paragraphs/tables, which an
 * object-keyed (non-`preserveOrder`) intermediate parse cannot.
 */

import { parseDocument } from './document'
import type { Block } from '../model/document'

const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WORD_2010_NAMESPACE = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
const WORD_2012_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml'
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const ID_ATTR_RE = /w:id\s*=\s*"([^"]*)"|w:id\s*=\s*'([^']*)'/

/**
 * Parse a raw XML fragment — the inner content of a wrapper element, i.e.
 * everything between its opening and closing tag — into `Block[]`
 * (paragraphs and tables alike) by re-rooting it under a synthetic
 * `<w:document><w:body>` and running it through the real `parseDocument()`.
 *
 * @param innerXml - Inner XML of a `w:hdr`/`w:ftr`/`w:footnote`/`w:endnote`/
 *   `w:comment` element (or `undefined`/empty for one with no content).
 * @returns Blocks in source order; `[]` when there are none.
 */
export function parseBlocksFromXmlFragment(innerXml: string | undefined): ReadonlyArray<Block> {
  if (innerXml === undefined || innerXml.trim().length === 0) {
    return []
  }

  const syntheticXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w10="${WORD_2010_NAMESPACE}" xmlns:w15="${WORD_2012_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}">`
    + `<w:body>${innerXml}</w:body></w:document>`

  const document = parseDocument(syntheticXml)
  const blocks: Block[] = []
  for (const section of document.sections) {
    blocks.push(...section.blocks)
  }
  return blocks
}

/**
 * Extract the exact inner XML of the single top-level `<tagName>` wrapper
 * in `xml` — for parts that contain exactly one such wrapper (header/footer
 * files: the whole part IS one `w:hdr`/`w:ftr`).
 *
 * @returns The inner XML substring, or `''` when the element is missing,
 *   self-closing, or its closing tag can't be found.
 */
export function extractSingleElementInnerXml(xml: string, tagName: string): string {
  const openTagRe = new RegExp(`<${tagName}\\b[^>]*?(/?)>`)
  const match = openTagRe.exec(xml)
  if (match === null) {
    return ''
  }
  if (match[1] === '/') {
    return ''
  }

  const closeTag = `</${tagName}>`
  const startIndex = match.index + match[0].length
  const closeIndex = xml.lastIndexOf(closeTag)
  if (closeIndex === -1 || closeIndex < startIndex) {
    return ''
  }

  return xml.slice(startIndex, closeIndex)
}

/**
 * Extract the exact inner XML of every `<tagName w:id="...">...</tagName>`
 * sibling in `xml`, keyed by its `w:id` attribute — for parts that hold
 * multiple such wrappers (footnotes.xml/endnotes.xml: several
 * `<w:footnote>`/`<w:endnote>` elements). Mirrors `comments.ts`'s
 * proven per-element regex-scanning technique for `<w:comment>`,
 * generalized here by tag name.
 */
export function extractElementInnerXmlsById(
  xml: string,
  tagName: string,
): ReadonlyMap<string, string> {
  const openTagRe = new RegExp(`<${tagName}\\b([^>]*?)(/?)>`, 'g')
  const closeTag = `</${tagName}>`
  const bodies = new Map<string, string>()

  let openTag: RegExpExecArray | null
  while ((openTag = openTagRe.exec(xml)) !== null) {
    const attrsText = openTag[1]
    const isSelfClosing = openTag[2] === '/'
    const idMatch = ID_ATTR_RE.exec(attrsText)
    const id = idMatch !== null ? (idMatch[1] ?? idMatch[2]) : undefined

    if (isSelfClosing) {
      if (id !== undefined) {
        bodies.set(id, '')
      }
      continue
    }

    const searchStart = openTagRe.lastIndex
    const closeIndex = xml.indexOf(closeTag, searchStart)
    if (closeIndex === -1) {
      // Malformed/truncated XML — the caller's own outer parse (used to
      // enumerate ids/types) will already have thrown for this before we
      // get here in practice.
      break
    }

    if (id !== undefined) {
      bodies.set(id, xml.slice(searchStart, closeIndex))
    }

    openTagRe.lastIndex = closeIndex + closeTag.length
  }

  return bodies
}
