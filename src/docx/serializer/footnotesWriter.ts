import type { Endnote, Footnote } from '../model'
import type { OrderedXmlNode, SerializeState } from './partWriterSupport'
import { buildBlockNodes, createSerializeState, serializeWordPart } from './partWriterSupport'

export function writeFootnotesXml(footnotes: ReadonlyArray<Footnote>): string {
  const state = createSerializeState()
  return serializeWordPart(
    'w:footnotes',
    footnotes.map((footnote) => buildNoteNode('w:footnote', footnote, state)),
    state,
  )
}

export function writeEndnotesXml(endnotes: ReadonlyArray<Endnote>): string {
  const state = createSerializeState()
  return serializeWordPart(
    'w:endnotes',
    endnotes.map((endnote) => buildNoteNode('w:endnote', endnote, state)),
    state,
  )
}

function buildNoteNode(
  name: 'w:footnote' | 'w:endnote',
  note: Footnote | Endnote,
  state: SerializeState,
): OrderedXmlNode {
  return {
    [name]: [...buildBlockNodes(note.blocks, state)],
    ':@': {
      '@_w:id': note.id,
      ...(note.noteType !== undefined && note.noteType !== 'normal' ? { '@_w:type': note.noteType } : {}),
    },
  }
}
