/**
 * Minimal DOM helpers shared by the OOXML writers (PPTX slide editing,
 * XLSX save-through-original). Every part is parsed with `DOMParser` and
 * walked by local name, so a differently-prefixed producer still works.
 */
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

/**
 * Text as XML 1.0 allows it. User text can carry C0 control characters
 * (pasted from another program); `XMLSerializer` would write them raw, and
 * Office refuses (or "repairs") a part that contains them.
 */
export function xmlSafeText(text: string): string {
  // eslint-disable-next-line no-control-regex
  return text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
}

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
