import type { Endnote, Footnote } from '../model'
import type { OrderedXmlNode } from './partWriterSupport'
import { buildParagraphBlockNodes, serializeWordPart } from './partWriterSupport'

export function writeFootnotesXml(footnotes: ReadonlyArray<Footnote>): string {
  return serializeWordPart('w:footnotes', footnotes.map((footnote) => buildNoteNode('w:footnote', footnote)), true)
}

export function writeEndnotesXml(endnotes: ReadonlyArray<Endnote>): string {
  return serializeWordPart('w:endnotes', endnotes.map((endnote) => buildNoteNode('w:endnote', endnote)), true)
}

function buildNoteNode(
  name: 'w:footnote' | 'w:endnote',
  note: Footnote | Endnote,
): OrderedXmlNode {
  return {
    [name]: [...buildParagraphBlockNodes(note.blocks)],
    ':@': {
      '@_w:id': note.id,
      ...(note.noteType !== undefined && note.noteType !== 'normal' ? { '@_w:type': note.noteType } : {}),
    },
  }
}
