import type { Header } from '../model'
import { buildBlockNodes, serializeWordPart } from './partWriterSupport'

export function writeHeaderXml(header: Header): string {
  return serializeWordPart('w:hdr', buildBlockNodes(header.blocks))
}
