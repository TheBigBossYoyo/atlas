import type { Endnote, Footnote, WrapperPassthrough } from '../model'
import { getWrapperRegionsForFragmentBlocks } from '../parser/partBody'
import type { OrderedXmlNode, SerializeState } from './partWriterSupport'
import { buildBlockNodes, createSerializeState, serializeWordPart } from './partWriterSupport'

/**
 * DOCX-2 (round-trip fidelity audit follow-up): unions every footnote/
 * endnote's own recorded wrapper regions (see `getWrapperRegionsForFragmentBlocks`'s
 * doc comment in `../parser/partBody.ts`) into one flat list for a shared
 * `SerializeState` — mirroring `documentWriter.ts`'s single `doc.wrappers`
 * list covering every section. Safe to share across notes: a region only
 * ever matches a note it did not come from if that note's blocks happen to
 * contain the exact same object instances at the exact same position,
 * which never happens (every note's blocks come from an independent
 * `parseDocument` call over its own source text).
 */
function collectWrapperRegions(notes: ReadonlyArray<Footnote | Endnote>): ReadonlyArray<WrapperPassthrough> {
  return notes.flatMap((note) => getWrapperRegionsForFragmentBlocks(note.blocks))
}

export function writeFootnotesXml(footnotes: ReadonlyArray<Footnote>): string {
  const state = createSerializeState(collectWrapperRegions(footnotes))
  return serializeWordPart(
    'w:footnotes',
    footnotes.map((footnote) => buildNoteNode('w:footnote', footnote, state)),
    state,
  )
}

export function writeEndnotesXml(endnotes: ReadonlyArray<Endnote>): string {
  const state = createSerializeState(collectWrapperRegions(endnotes))
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
