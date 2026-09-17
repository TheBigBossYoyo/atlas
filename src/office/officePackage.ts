/**
 * USR-16 — an editable Office package (PPTX, ODP): every part of the original
 * zip held in memory, immutably. Edits replace only the parts they touch, so
 * masters, layouts, themes, media and anything Atlas does not understand pass
 * through a save byte-for-byte — the same passthrough approach as the DOCX
 * serializer. Undo keeps whole package snapshots; they share every untouched
 * part.
 *
 * ODF note: `mimetype` must be the first entry in an ODF zip and must be
 * stored uncompressed, so the writer puts it back exactly that way.
 */
import JSZip from 'jszip'

import type { ZipArchive, ZipEntry } from '../viewers/slides/shared/xmlUtils'

export type OfficePackage = {
  /** Part path -> content: XML/rels as text, everything else as bytes. Insertion order is the zip order on save. */
  readonly parts: ReadonlyMap<string, string | Uint8Array>
}

const TEXT_PART = /\.(xml|rels)$/i

export async function loadOfficePackage(buffer: ArrayBuffer): Promise<OfficePackage> {
  const zip = await JSZip.loadAsync(buffer)
  const parts = new Map<string, string | Uint8Array>()
  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue
    parts.set(entry.name, TEXT_PART.test(entry.name) ? await entry.async('string') : await entry.async('uint8array'))
  }
  return { parts }
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  }
  return btoa(binary)
}

/** Read-only `ZipArchive` view for the slide parser. */
export function packageArchive(pkg: OfficePackage): ZipArchive {
  return {
    file(path: string): ZipEntry | null {
      const content = pkg.parts.get(path)
      if (content === undefined) return null
      const read = async (type: 'string' | 'base64'): Promise<string> => {
        if (type === 'string') return typeof content === 'string' ? content : new TextDecoder().decode(content)
        return bytesToBase64(typeof content === 'string' ? new TextEncoder().encode(content) : content)
      }
      return { async: read } as ZipEntry
    },
  }
}

export function readPart(pkg: OfficePackage, path: string): string | null {
  const content = pkg.parts.get(path)
  return typeof content === 'string' ? content : null
}

/** Returns a package with the given parts replaced/added (`null` removes a part). */
export function withParts(pkg: OfficePackage, changes: Readonly<Record<string, string | null>>): OfficePackage {
  const parts = new Map(pkg.parts)
  for (const [path, content] of Object.entries(changes)) {
    if (content === null) parts.delete(path)
    else parts.set(path, content)
  }
  return { parts }
}

export async function writeOfficePackage(pkg: OfficePackage): Promise<Uint8Array> {
  const zip = new JSZip()
  // ODF: `mimetype` first and uncompressed, or the file is not recognized.
  const mimetype = pkg.parts.get('mimetype')
  if (mimetype !== undefined) zip.file('mimetype', mimetype, { compression: 'STORE' })
  for (const [path, content] of pkg.parts) {
    if (path === 'mimetype') continue
    zip.file(path, content)
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
