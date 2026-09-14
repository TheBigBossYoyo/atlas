/**
 * Shared DOM/zip helpers for the PPTX and ODP parsers (Wave 3-S).
 *
 * Every XML document is parsed with `DOMParser` and walked with
 * `textContent`/`getElementsByTagNameNS('*', localName)` — never `innerHTML`
 * of untrusted document strings — per the plan's parsing constraint.
 */

export type ZipEntry = {
  async(type: 'string'): Promise<string>
  async(type: 'base64'): Promise<string>
}

export type ZipArchive = {
  file(path: string): ZipEntry | null
}

export function parseXml(xml: string): XMLDocument {
  const document = new DOMParser().parseFromString(xml, 'application/xml')
  const parserError = document.getElementsByTagName('parsererror')[0]

  if (parserError) {
    throw new Error(parserError.textContent?.trim() || 'Invalid XML document.')
  }

  return document
}

/** All descendants matching `localName`, regardless of namespace prefix. */
export function getElementsByLocalName(root: XMLDocument | Element, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS('*', localName))
}

export function getFirstByLocalName(root: XMLDocument | Element, localName: string): Element | null {
  return getElementsByLocalName(root, localName)[0] ?? null
}

/** Direct children only (order-preserving), optionally filtered by local name. */
export function getDirectChildren(element: Element, localName?: string): Element[] {
  const children = Array.from(element.children)
  return localName ? children.filter(child => child.localName === localName) : children
}

export function getFirstDirectChild(element: Element, localName: string): Element | null {
  return getDirectChildren(element, localName)[0] ?? null
}

/** Reads the first present attribute among several accepted spellings/prefixes. */
export function getAttributeValue(element: Element, names: ReadonlyArray<string>): string | null {
  for (const name of names) {
    const value = element.getAttribute(name)
    if (value !== null) {
      return value
    }
  }

  return null
}

export function parseNumber(value: string | null): number | undefined {
  if (value === null) {
    return undefined
  }

  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

export function normalizeZipPath(path: string): string {
  const parts: string[] = []

  for (const segment of path.replace(/\\/g, '/').split('/')) {
    if (!segment || segment === '.') {
      continue
    }

    if (segment === '..') {
      parts.pop()
      continue
    }

    parts.push(segment)
  }

  return parts.join('/')
}

export function resolvePartPath(sourcePartPath: string, targetPath: string): string {
  if (targetPath.startsWith('/')) {
    return normalizeZipPath(targetPath.slice(1))
  }

  const sourceParts = sourcePartPath.split('/')
  sourceParts.pop()

  return normalizeZipPath([...sourceParts, targetPath].join('/'))
}

export function getRequiredZipEntry(zip: ZipArchive, path: string): ZipEntry {
  const entry = zip.file(path)

  if (!entry) {
    throw new Error(`Missing archive entry: ${path}`)
  }

  return entry
}

export function getMimeTypeFromPath(path: string): string {
  const extension = path.split('.').pop()?.toLowerCase()

  switch (extension) {
    case 'png':
      return 'image/png'
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg'
    case 'gif':
      return 'image/gif'
    case 'bmp':
      return 'image/bmp'
    case 'svg':
      return 'image/svg+xml'
    case 'webp':
      return 'image/webp'
    case 'tif':
    case 'tiff':
      return 'image/tiff'
    case 'wmf':
      return 'image/wmf'
    case 'emf':
      return 'image/emf'
    default:
      return 'application/octet-stream'
  }
}

export type CancelSignal = {
  cancelled: boolean
}

export async function readZipText(zip: ZipArchive, path: string, signal: CancelSignal): Promise<string | null> {
  if (signal.cancelled) {
    return null
  }

  const entry = zip.file(path)
  if (!entry) {
    return null
  }

  return entry.async('string')
}

export async function toDataUrl(zip: ZipArchive, path: string, signal: CancelSignal): Promise<string | null> {
  if (signal.cancelled) {
    return null
  }

  const entry = zip.file(path)
  if (!entry) {
    return null
  }

  const base64 = await entry.async('base64')

  if (signal.cancelled) {
    return null
  }

  return `data:${getMimeTypeFromPath(path)};base64,${base64}`
}
