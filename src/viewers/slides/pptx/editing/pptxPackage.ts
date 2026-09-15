/**
 * USR-16 — the editable PPTX package: every part of the original zip held in
 * memory, immutably. Edits replace only the parts they touch (see
 * `pptxEdits.ts`/`pptxSlideOps.ts`), so masters, layouts, themes, media and
 * anything Atlas does not understand pass through a save byte-for-byte —
 * the same passthrough approach as the DOCX serializer. Undo keeps whole
 * package snapshots; they share every untouched part.
 */
import JSZip from 'jszip'

import type { ZipArchive, ZipEntry } from '../../shared/xmlUtils'

export type PptxPackage = {
  /** Part path -> content: XML/rels as text, everything else as bytes. Insertion order is the zip order on save. */
  readonly parts: ReadonlyMap<string, string | Uint8Array>
}

const TEXT_PART = /\.(xml|rels)$/i

export async function loadPptxPackage(buffer: ArrayBuffer): Promise<PptxPackage> {
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
export function packageArchive(pkg: PptxPackage): ZipArchive {
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

export function readPart(pkg: PptxPackage, path: string): string | null {
  const content = pkg.parts.get(path)
  return typeof content === 'string' ? content : null
}

/** Returns a package with the given parts replaced/added (`null` removes a part). */
export function withParts(pkg: PptxPackage, changes: Readonly<Record<string, string | null>>): PptxPackage {
  const parts = new Map(pkg.parts)
  for (const [path, content] of Object.entries(changes)) {
    if (content === null) parts.delete(path)
    else parts.set(path, content)
  }
  return { parts }
}

export async function writePptxPackage(pkg: PptxPackage): Promise<Uint8Array> {
  const zip = new JSZip()
  for (const [path, content] of pkg.parts) zip.file(path, content)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
