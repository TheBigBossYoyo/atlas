import type { Footer } from '../model'
import { buildBlockNodes, serializeWordPart } from './partWriterSupport'

export function writeFooterXml(footer: Footer): string {
  return serializeWordPart('w:ftr', buildBlockNodes(footer.blocks), true)
}
