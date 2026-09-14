import type { Header } from '../model'
import { buildBlockNodes, createSerializeState, serializeWordPart } from './partWriterSupport'

export function writeHeaderXml(header: Header): string {
  const state = createSerializeState()
  return serializeWordPart('w:hdr', buildBlockNodes(header.blocks, state), state)
}
