import type { Header } from '../model'
import { buildParagraphBlockNodes, serializeWordPart } from './partWriterSupport'

export function writeHeaderXml(header: Header): string {
  return serializeWordPart('w:hdr', buildParagraphBlockNodes(header.blocks), true)
}
