/**
 * Atlas — DOCX zip packager (Wave D.4)
 *
 * Assembles a map of part-name → content into a .docx Uint8Array using JSZip.
 * Parts are emitted in Microsoft Word's canonical order to minimise diffs.
 */

import JSZip from 'jszip'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export interface DocxPart {
  readonly path: string
  readonly content: string | Uint8Array
}

// ---------------------------------------------------------------------------
// Canonical Word part order
// ---------------------------------------------------------------------------

/**
 * Canonical path patterns in the order Microsoft Word emits them.
 * Exact matches are preferred; prefix matches are used for wildcard-like
 * entries (e.g. `word/header` covers `word/header1.xml`, `word/header2.xml`).
 * Any part not matching a pattern is appended alphabetically at the end.
 */
export const WORD_PART_ORDER: ReadonlyArray<string> = [
  '[Content_Types].xml',
  '_rels/.rels',
  'word/document.xml',
  'word/_rels/document.xml.rels',
  'word/styles.xml',
  'word/numbering.xml',
  'word/settings.xml',
  'word/webSettings.xml',
  'word/fontTable.xml',
  'word/theme/theme1.xml',
  'word/header',
  'word/footer',
  'word/footnotes.xml',
  'word/endnotes.xml',
  'word/comments.xml',
  'word/media/',
  'docProps/core.xml',
  'docProps/app.xml',
]

// ---------------------------------------------------------------------------
// Sorting
// ---------------------------------------------------------------------------

/**
 * Returns the canonical priority index for a given path.
 * Lower index = earlier in the archive.
 * Paths not matched by any pattern get index `WORD_PART_ORDER.length`.
 */
function priorityOf(path: string): number {
  for (let i = 0; i < WORD_PART_ORDER.length; i++) {
    const pattern = WORD_PART_ORDER[i]
    // Exact match wins first.
    if (path === pattern) return i
    // Prefix match for wildcard-like patterns (header*, footer*, media/*).
    if (path.startsWith(pattern)) return i
  }
  return WORD_PART_ORDER.length
}

/**
 * Pure sort — returns a **new** array of parts in canonical Word order.
 * Parts with the same canonical priority are sorted alphabetically by path.
 * The input array is never mutated.
 */
export function sortPartsForWord(
  parts: ReadonlyArray<DocxPart>,
): ReadonlyArray<DocxPart> {
  return [...parts].sort((a, b) => {
    const pa = priorityOf(a.path)
    const pb = priorityOf(b.path)
    if (pa !== pb) return pa - pb
    return a.path < b.path ? -1 : a.path > b.path ? 1 : 0
  })
}

// ---------------------------------------------------------------------------
// Packager
// ---------------------------------------------------------------------------

/**
 * Assembles a DOCX archive from the supplied parts and returns the raw bytes.
 *
 * Parts are sorted into canonical Word order before insertion.
 * String content is stored as UTF-8; `Uint8Array` content is stored as-is.
 * All files are compressed with DEFLATE at level 6 (Word's default).
 *
 * @param parts - Array of `{ path, content }` entries to include.
 * @returns       The assembled `.docx` bytes as a `Uint8Array`.
 */
export async function packDocx(
  parts: ReadonlyArray<DocxPart>,
): Promise<Uint8Array> {
  const zip = new JSZip()

  const sorted = sortPartsForWord(parts)

  for (const part of sorted) {
    zip.file(part.path, part.content, {
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    })
  }

  const buffer = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })

  return buffer
}
