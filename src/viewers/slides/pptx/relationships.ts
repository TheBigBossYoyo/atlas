/** OPC relationship (`.rels`) parsing shared by every PPTX resolution step. */

import type { CancelSignal, ZipArchive } from '../shared/xmlUtils'
import { getAttributeValue, getElementsByLocalName, parseXml, readZipText, resolvePartPath } from '../shared/xmlUtils'

export type Relationship = {
  readonly id: string
  readonly type: string
  readonly target: string
}

export function relsPathFor(partPath: string): string {
  const lastSlash = partPath.lastIndexOf('/')
  const directory = lastSlash === -1 ? '' : partPath.slice(0, lastSlash)
  const fileName = lastSlash === -1 ? partPath : partPath.slice(lastSlash + 1)
  return directory ? `${directory}/_rels/${fileName}.rels` : `_rels/${fileName}.rels`
}

export async function parseRelationships(
  zip: ZipArchive,
  sourcePartPath: string,
  signal: CancelSignal,
): Promise<ReadonlyArray<Relationship>> {
  const xml = await readZipText(zip, relsPathFor(sourcePartPath), signal)
  if (!xml) {
    return []
  }

  const document = parseXml(xml)
  const relationships: Relationship[] = []

  for (const relationship of getElementsByLocalName(document, 'Relationship')) {
    const id = relationship.getAttribute('Id')
    const type = relationship.getAttribute('Type')
    const target = relationship.getAttribute('Target')
    const targetMode = relationship.getAttribute('TargetMode')

    if (!id || !type || !target || targetMode === 'External') {
      continue
    }

    relationships.push({ id, type, target: resolvePartPath(sourcePartPath, target) })
  }

  return relationships
}

export function relationshipsById(relationships: ReadonlyArray<Relationship>): Map<string, string> {
  return new Map(relationships.map(relationship => [relationship.id, relationship.target]))
}

export function findRelationshipTarget(
  relationships: ReadonlyArray<Relationship>,
  typeSuffix: string,
): string | undefined {
  return relationships.find(relationship => relationship.type.endsWith(typeSuffix))?.target
}

/** `r:id` reference (slide->layout, layout->master, master->theme, notes, ...). */
export function resolveRefId(element: Element): string | null {
  return getAttributeValue(element, ['r:id', 'id'])
}

/** `r:embed` reference (a picture's `<a:blip>`). */
export function resolveEmbedId(element: Element): string | null {
  return getAttributeValue(element, ['r:embed', 'embed'])
}
