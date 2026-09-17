/**
 * Minimal DOM helpers shared by the OOXML writers (PPTX slide editing,
 * XLSX save-through-original). Every part is parsed with `DOMParser` and
 * walked by local name, so a differently-prefixed producer still works.
 */
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

export function parseXmlPart(xml: string): XMLDocument {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) throw new Error('Invalid XML part.')
  return doc
}

export function serializeXmlPart(doc: XMLDocument): string {
  return XML_DECLARATION + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '')
}

export function childElements(element: Element, localName: string): Element[] {
  return Array.from(element.children).filter((child) => child.localName === localName)
}

export function firstChildElement(element: Element, localName: string): Element | null {
  return childElements(element, localName)[0] ?? null
}

export function descendantElements(root: Document | Element, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS('*', localName))
}
