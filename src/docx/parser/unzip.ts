/**
 * Atlas — DOCX unzip (Wave A.1)
 *
 * Wraps JSZip to extract a DOCX (ZIP) archive into an immutable map of
 * path → raw bytes.  All errors are surfaced as `DocxParseError`.
 */

import JSZip from 'jszip'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** Immutable representation of a DOCX archive's files. */
export interface DocxArchive {
  readonly files: ReadonlyMap<string, Uint8Array>
}

/** Thrown when the input cannot be parsed as a valid DOCX / ZIP archive. */
export class DocxParseError extends Error {
  /** The archive-path that triggered the error (empty string for top-level). */
  readonly path: string

  constructor(message: string, path: string = '') {
    super(message)
    this.name = 'DocxParseError'
    this.path = path
    // Restore prototype chain in transpiled environments.
    Object.setPrototypeOf(this, new.target.prototype)
  }
}

// ---------------------------------------------------------------------------
// Implementation
// ---------------------------------------------------------------------------

/**
 * Unzips a DOCX buffer and returns every entry as a `Uint8Array`.
 *
 * @param buffer - Raw DOCX bytes (ArrayBuffer or Uint8Array accepted).
 * @returns       A `DocxArchive` whose `files` map is keyed by ZIP entry path.
 * @throws        `DocxParseError` when the input is not a valid ZIP/DOCX.
 */
export async function unzipDocx(
  buffer: ArrayBuffer | Uint8Array,
): Promise<DocxArchive> {
  let zip: JSZip

  try {
    zip = await JSZip.loadAsync(buffer)
  } catch (cause) {
    const msg =
      cause instanceof Error ? cause.message : 'Unknown decompression error'
    throw new DocxParseError(`Failed to unzip DOCX: ${msg}`)
  }

  const files = new Map<string, Uint8Array>()

  const entries = Object.entries(zip.files).filter(([, entry]) => !entry.dir)

  await Promise.all(
    entries.map(async ([path, entry]) => {
      try {
        const bytes = await entry.async('uint8array')
        files.set(path, bytes)
      } catch (cause) {
        const msg =
          cause instanceof Error ? cause.message : 'Unknown extraction error'
        throw new DocxParseError(
          `Failed to extract entry "${path}": ${msg}`,
          path,
        )
      }
    }),
  )

  return { files }
}
