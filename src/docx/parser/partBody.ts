/**
 * Atlas — shared "paragraph-array body" parser
 *
 * Comments, headers, footers, footnotes, and endnotes all hold ordinary
 * paragraph content nested inside their own part-specific wrapper element
 * (`w:comment`, `w:hdr`, `w:ftr`, `w:footnote`, `w:endnote`). None of those
 * wrappers has a `w:body` child, so none of them can be handed directly to
 * `parseDocument`, which only knows how to find `w:document`/`w:body`.
 *
 * `comments.ts` solved this for comment bodies by re-rooting the extracted
 * `w:p` nodes under a synthetic `<w:document><w:body>` element and running
 * them back through `parseDocument`. This module lifts that same technique
 * into a shared helper so headers/footers/footnotes/endnotes (Wave A.5's
 * "Option B" stub — DXP-03/DXP-04/DXS-01) can produce real `Paragraph[]`
 * blocks the same way, instead of one opaque `UnknownNode` that made
 * `buildParagraphBlockNodes` throw on every save.
 */

import { XMLBuilder } from 'fast-xml-parser'

import { parseDocument } from './document'
import type { Paragraph } from '../model/document'

const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WORD_2010_NAMESPACE = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
const WORD_2012_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml'
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

/**
 * Parse the raw `w:p` node(s) extracted from a part's wrapper element into
 * real `Paragraph[]` blocks.
 *
 * @param paragraphNodes - The value of a `w:p` key as produced by an
 *   object-keyed (non-`preserveOrder`) `fast-xml-parser` parse: a single
 *   node object, an array of node objects, or `undefined`/`null` when the
 *   wrapper has no paragraph children at all.
 * @returns Paragraphs in source order; `[]` when there are none.
 */
export function parseParagraphsFromRawNodes(paragraphNodes: unknown): ReadonlyArray<Paragraph> {
  if (paragraphNodes === undefined || paragraphNodes === null) {
    return []
  }

  const nodes = Array.isArray(paragraphNodes) ? paragraphNodes : [paragraphNodes]
  if (nodes.length === 0) {
    return []
  }

  const syntheticXml = xmlBuilder.build({
    'w:document': {
      '@_xmlns:w': WORD_NAMESPACE,
      '@_xmlns:w10': WORD_2010_NAMESPACE,
      '@_xmlns:w15': WORD_2012_NAMESPACE,
      '@_xmlns:r': RELATIONSHIPS_NAMESPACE,
      'w:body': {
        'w:p': nodes,
      },
    },
  })

  const document = parseDocument(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${syntheticXml}`,
  )
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
