/**
 * Zip-bomb guard for the spreadsheet passthrough readers (USR-17 security
 * review finding): `xlsxPassthrough.ts`, `spreadsheetTables.ts` and
 * `spreadsheetPanes.ts` each call `JSZip.loadAsync` directly on a workbook
 * buffer that is untrusted input by definition (the app's entire job is
 * opening a file the user picked), with no check on how much memory
 * decompressing it could actually take — the same "small compressed archive,
 * huge declared uncompressed size" zip-bomb risk `docx/parser/unzip.ts`'s
 * `unzipDocx` already guards against for DOCX/PPTX (DXP-15).
 *
 * This is deliberately a lighter guard than `unzipDocx`, not a reuse of it:
 * `unzipDocx` extracts every entry up front into a plain `files` map, but
 * `xlsxPassthrough.ts` needs to keep mutating and re-serializing the SAME
 * live `JSZip` instance so every part it never touches stays byte-identical
 * on save — extracting to a map and rebuilding a fresh zip would defeat the
 * whole point of a passthrough save. So this only runs `unzipDocx`'s pass 1
 * (reject on each entry's DECLARED uncompressed size, read off the zip
 * central directory, before a single byte is inflated) and hands back the
 * `JSZip` itself; there is no pass-2 actual-size backstop here, since each
 * part is still inflated lazily by whichever caller needs it. The declared
 * size is exactly what a crafted small-file/huge-declared zip bomb lies
 * about, so this still refuses it before any inflation starts — the same
 * numbers a normal user's workbook is nowhere near just pass straight
 * through.
 */
import JSZip from 'jszip'

import {
  DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES,
  DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES,
} from '../../docx/parser/unzip'

/** Thrown instead of loading a zip whose declared uncompressed size is over budget. */
export class SpreadsheetZipBombError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'SpreadsheetZipBombError'
    Object.setPrototypeOf(this, new.target.prototype)
  }
}

/** Same private-field shape `docx/parser/unzip.ts` reads defensively — see that module's identical comment. */
interface JSZipEntryWithDeclaredSize extends JSZip.JSZipObject {
  readonly _data?: { readonly uncompressedSize?: number }
}

function declaredUncompressedSize(entry: JSZip.JSZipObject): number | undefined {
  const size = (entry as JSZipEntryWithDeclaredSize)._data?.uncompressedSize
  return typeof size === 'number' && Number.isFinite(size) ? size : undefined
}

/**
 * Loads `buffer` as a zip, first rejecting it (before decompressing
 * anything) if any entry's declared uncompressed size is implausibly large,
 * individually or combined. Budgets default to the same numbers `unzipDocx`
 * uses for DOCX/PPTX, for one consistent "this is too big to be a real
 * document" line across every Office format Atlas opens.
 */
export async function loadWorkbookZip(
  buffer: ArrayBuffer | Uint8Array,
  maxEntryUncompressedBytes: number = DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES,
  maxTotalUncompressedBytes: number = DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES,
): Promise<JSZip> {
  const zip = await JSZip.loadAsync(buffer)

  let total = 0
  for (const [path, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue
    const size = declaredUncompressedSize(entry)
    if (size === undefined) continue // no declared size on record — nothing to check against for this entry

    if (size > maxEntryUncompressedBytes) {
      throw new SpreadsheetZipBombError(
        `Entry "${path}" declares ${size.toLocaleString()} bytes uncompressed, over the ` +
          `${maxEntryUncompressedBytes.toLocaleString()}-byte per-entry limit; refusing to open ` +
          '(possible zip bomb).',
      )
    }
    total += size
    if (total > maxTotalUncompressedBytes) {
      throw new SpreadsheetZipBombError(
        `Archive's combined declared uncompressed size is over the ` +
          `${maxTotalUncompressedBytes.toLocaleString()}-byte total limit; refusing to open ` +
          '(possible zip bomb).',
      )
    }
  }

  return zip
}
