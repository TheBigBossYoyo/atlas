import type { Footer } from '../model'
import { buildParagraphBlockNodes, serializeWordPart } from './partWriterSupport'

export function writeFooterXml(footer: Footer): string {
  return serializeWordPart('w:ftr', buildParagraphBlockNodes(footer.blocks), true)
}
