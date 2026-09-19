/**
 * USR-16 — small DOM helpers for rewriting OPC parts (slide XML,
 * relationships, content types) during PPTX editing.
 */
import { EMU_PER_PIXEL } from '../../shared/units'
import {
  childElements,
  descendantElements,
  firstChildElement,
  parseXmlPart,
  serializeXmlPart,
  xmlSafeText,
} from '../../../../office/ooxmlDom'

export const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  types: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const

export const REL_TYPE = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
} as const

export const CONTENT_TYPE = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
} as const

/** Re-exported under this module's shorter names (see src/office/ooxmlDom.ts). */
export const parse = parseXmlPart
export const serialize = serializeXmlPart
export const children = childElements
export const firstChild = firstChildElement
export const descendants = descendantElements

/** Parses an XML fragment (with `p:`/`a:`/`r:` prefixes available) and imports it into `doc`. */
export function importFragment(doc: XMLDocument, xml: string): Element {
  const wrapper = parse(`<w xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${xml}</w>`)
  return doc.importNode(wrapper.documentElement.firstElementChild!, true) as Element
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PIXEL)
}

export function escapeXml(text: string): string {
  return xmlSafeText(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ---------------------------------------------------------------------------
// Relationships
// ---------------------------------------------------------------------------

export type Rel = { readonly id: string; readonly type: string; readonly target: string }

export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/')
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`
}

/** `ppt/slides/slide1.xml` + `../slideLayouts/slideLayout2.xml` -> `ppt/slideLayouts/slideLayout2.xml`. */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const segments = sourcePart.split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part === '..') segments.pop()
    else if (part !== '.' && part !== '') segments.push(part)
  }
  return segments.join('/')
}

/** Inverse of `resolveTarget`, for parts under the same root. */
export function relativeTarget(sourcePart: string, targetPart: string): string {
  const from = sourcePart.split('/').slice(0, -1)
  const to = targetPart.split('/')
  let common = 0
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++
  return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/')
}

export function readRels(relsXml: string | null, sourcePart: string): Rel[] {
  if (!relsXml) return []
  return descendants(parse(relsXml), 'Relationship').flatMap((rel) => {
    const id = rel.getAttribute('Id')
    const type = rel.getAttribute('Type')
    const target = rel.getAttribute('Target')
    if (!id || !type || !target || rel.getAttribute('TargetMode') === 'External') return []
    return [{ id, type, target: resolveTarget(sourcePart, target) }]
  })
}

function emptyRels(): XMLDocument {
  return parse(`<Relationships xmlns="${NS.rels}"></Relationships>`)
}

/** Adds a relationship and returns the new rels XML with its id. */
export function addRel(
  relsXml: string | null,
  sourcePart: string,
  type: string,
  targetPart: string,
): { readonly xml: string; readonly id: string } {
  const doc = relsXml ? parse(relsXml) : emptyRels()
  const used = new Set(descendants(doc, 'Relationship').map((rel) => rel.getAttribute('Id')))
  let n = used.size + 1
  while (used.has(`rId${n}`)) n++
  const rel = doc.createElementNS(NS.rels, 'Relationship')
  rel.setAttribute('Id', `rId${n}`)
  rel.setAttribute('Type', type)
  rel.setAttribute('Target', relativeTarget(sourcePart, targetPart))
  doc.documentElement.appendChild(rel)
  return { xml: serialize(doc), id: `rId${n}` }
}

export function removeRels(relsXml: string, predicate: (rel: Element) => boolean): string {
  const doc = parse(relsXml)
  for (const rel of descendants(doc, 'Relationship')) {
    if (predicate(rel)) rel.parentNode?.removeChild(rel)
  }
  return serialize(doc)
}

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

export function addOverride(contentTypesXml: string, partPath: string, contentType: string): string {
  const doc = parse(contentTypesXml)
  const override = doc.createElementNS(NS.types, 'Override')
  override.setAttribute('PartName', `/${partPath}`)
  override.setAttribute('ContentType', contentType)
  doc.documentElement.appendChild(override)
  return serialize(doc)
}

export function removeOverrides(contentTypesXml: string, partPaths: ReadonlyArray<string>): string {
  const names = new Set(partPaths.map((path) => `/${path}`))
  const doc = parse(contentTypesXml)
  for (const override of descendants(doc, 'Override')) {
    if (names.has(override.getAttribute('PartName') ?? '')) override.parentNode?.removeChild(override)
  }
  return serialize(doc)
}

/** Next free `ppt/<dir>/<stem>N.xml` part name. */
export function nextPartName(existing: Iterable<string>, dir: string, stem: string): string {
  const pattern = new RegExp(`^${dir}/${stem}(\\d+)\\.xml$`)
  let max = 0
  for (const path of existing) {
    const match = pattern.exec(path)
    if (match) max = Math.max(max, Number(match[1]))
  }
  return `${dir}/${stem}${max + 1}.xml`
}
