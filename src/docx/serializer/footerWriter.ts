import type { Footer } from '../model'
import { buildBlockNodes, createSerializeState, serializeWordPart } from './partWriterSupport'

export function writeFooterXml(footer: Footer): string {
  const state = createSerializeState()
  return serializeWordPart('w:ftr', buildBlockNodes(footer.blocks, state), state)
}
