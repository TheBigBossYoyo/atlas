/**
 * Atlas — DOCX unzip (Wave A.1, hardened in D21)
 *
 * Wraps JSZip to extract a DOCX (ZIP) archive into an immutable map of
 * path → raw bytes.  All errors are surfaced as `DocxParseError`.
 *
 * A DOCX is untrusted input by definition (the app's entire job is opening
 * files a user picked, which may come from anywhere). Zip is trivially
 * "bombable": a small compressed archive can declare an enormous
 * uncompressed size and hang or crash the renderer the moment something
 * tries to materialize it. This module guards against that in two layers
 * (DXP-15):
 *   1. Before decompressing anything, reject any entry whose *declared*
 *      uncompressed size (from the ZIP central directory, which JSZip
 *      exposes on each entry) exceeds a per-entry budget, or whose running
 *      total across all entries exceeds an overall budget.
 *   2. While actually decompressing (in small batches rather than one
 *      unconditional `Promise.all`, to bound peak concurrent memory),
 *      re-check the *actual* decompressed byte length against the same
 *      budgets — a backstop for the rare case a declared size is absent or
 *      does not match reality.
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

/** Options controlling `unzipDocx`'s zip-bomb guards (DXP-15). */
export interface UnzipDocxOptions {
  /**
   * Per-entry uncompressed-size budget, in bytes. An entry declaring (or
   * actually decompressing to) more than this is rejected outright.
   * Defaults to {@link DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES}.
   */
  readonly maxEntryUncompressedBytes?: number
  /**
   * Running-total uncompressed-size budget across every entry combined, in
   * bytes. Defaults to {@link DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES}.
   */
  readonly maxTotalUncompressedBytes?: number
  /**
   * Maximum number of entries decompressed concurrently. Bounds peak memory
   * use during extraction instead of inflating every entry at once.
   * Defaults to {@link DEFAULT_EXTRACTION_CONCURRENCY}.
   */
  readonly concurrency?: number
}

/** Default per-entry uncompressed-size budget: 200 MiB. */
export const DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES = 200 * 1024 * 1024

/** Default running-total uncompressed-size budget: 500 MiB. */
export const DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES = 500 * 1024 * 1024

/** Default number of entries decompressed concurrently. */
export const DEFAULT_EXTRACTION_CONCURRENCY = 8

// ---------------------------------------------------------------------------
// JSZip internals this module reads defensively
// ---------------------------------------------------------------------------

/**
 * JSZip records each entry's compressed/uncompressed size (read straight off
 * the ZIP central directory) on a private `_data` field that isn't part of
 * its public `.d.ts`. It is nonetheless a stable, widely-relied-on shape for
 * exactly this "check the size before inflating" use case. Accessed
 * defensively — every field is optional and a missing/malformed shape simply
 * means pass 1 treats the size as unknown rather than throwing.
 */
interface JSZipEntryWithDeclaredSize extends JSZip.JSZipObject {
  readonly _data?: {
    readonly uncompressedSize?: number
  }
}

function getDeclaredUncompressedSize(entry: JSZip.JSZipObject): number | undefined {
  const size = (entry as JSZipEntryWithDeclaredSize)._data?.uncompressedSize
  return typeof size === 'number' && Number.isFinite(size) ? size : undefined
}

// ---------------------------------------------------------------------------
// Implementation
// ---------------------------------------------------------------------------

/**
 * Unzips a DOCX buffer and returns every entry as a `Uint8Array`.
 *
 * @param buffer  - Raw DOCX bytes (ArrayBuffer or Uint8Array accepted).
 * @param options - Zip-bomb guard budgets; see {@link UnzipDocxOptions}.
 * @returns         A `DocxArchive` whose `files` map is keyed by ZIP entry path.
 * @throws          `DocxParseError` when the input is not a valid ZIP/DOCX,
 *                   or when a size budget is exceeded (a likely zip bomb).
 */
export async function unzipDocx(
  buffer: ArrayBuffer | Uint8Array,
  options: UnzipDocxOptions = {},
): Promise<DocxArchive> {
  const maxEntryBytes = options.maxEntryUncompressedBytes ?? DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES
  const maxTotalBytes = options.maxTotalUncompressedBytes ?? DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES
  const concurrency = options.concurrency ?? DEFAULT_EXTRACTION_CONCURRENCY

  let zip: JSZip

  try {
    zip = await JSZip.loadAsync(buffer)
  } catch (cause) {
    const msg =
      cause instanceof Error ? cause.message : 'Unknown decompression error'
    throw new DocxParseError(`Failed to unzip DOCX: ${msg}`)
  }

  const entries = Object.entries(zip.files).filter(([, entry]) => !entry.dir)

  // Pass 1: reject on DECLARED size, before decompressing a single byte —
  // this is what makes a KB-compressed/GB-uncompressed zip bomb fail
  // instantly instead of hanging while JSZip inflates it.
  let declaredTotal = 0
  for (const [path, entry] of entries) {
    const declaredSize = getDeclaredUncompressedSize(entry)
    if (declaredSize === undefined) {
      continue
    }

    assertWithinEntryBudget(declaredSize, path, maxEntryBytes)
    declaredTotal += declaredSize
    assertWithinTotalBudget(declaredTotal, maxTotalBytes)
  }

  // Pass 2: extract in small batches (bounding peak concurrent memory
  // rather than inflating every entry at once) and re-check the ACTUAL
  // decompressed size against the same budgets as a backstop for the rare
  // case a declared size was absent or understated.
  const files = new Map<string, Uint8Array>()
  let extractedTotal = 0

  for (let start = 0; start < entries.length; start += concurrency) {
    const batch = entries.slice(start, start + concurrency)

    const extracted = await Promise.all(
      batch.map(async ([path, entry]): Promise<readonly [string, Uint8Array]> => {
        try {
          const bytes = await entry.async('uint8array')
          assertWithinEntryBudget(bytes.length, path, maxEntryBytes)
          return [path, bytes]
        } catch (cause) {
          if (cause instanceof DocxParseError) {
            throw cause
          }
          const msg =
            cause instanceof Error ? cause.message : 'Unknown extraction error'
          throw new DocxParseError(`Failed to extract entry "${path}": ${msg}`, path)
        }
      }),
    )

    for (const [path, bytes] of extracted) {
      extractedTotal += bytes.length
      assertWithinTotalBudget(extractedTotal, maxTotalBytes)
      files.set(path, bytes)
    }
  }

  return { files }
}

function assertWithinEntryBudget(sizeBytes: number, path: string, maxEntryBytes: number): void {
  if (sizeBytes > maxEntryBytes) {
    throw new DocxParseError(
      `Entry "${path}" is ${sizeBytes.toLocaleString()} bytes uncompressed, over the ` +
        `${maxEntryBytes.toLocaleString()}-byte per-entry limit; refusing to extract ` +
        '(possible zip bomb).',
      path,
    )
  }
}

function assertWithinTotalBudget(totalBytes: number, maxTotalBytes: number): void {
  if (totalBytes > maxTotalBytes) {
    throw new DocxParseError(
      `Archive's combined uncompressed size is ${totalBytes.toLocaleString()} bytes, over the ` +
        `${maxTotalBytes.toLocaleString()}-byte total limit; refusing to extract ` +
        '(possible zip bomb).',
    )
  }
}
