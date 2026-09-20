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
 *
 * `checkWorkbookZipBudgetSync`, below `loadWorkbookZip`, is a second, purely
 * synchronous guard for `spreadsheetGrid.ts`'s `parseWorkbookBuffer` (SHEET-1)
 * — the primary xlsx/ods read path, which has too many pre-existing
 * synchronous callers to route through the `async`/`JSZip`-based check
 * above. See that function's own header for why it exists separately rather
 * than just awaiting `loadWorkbookZip`.
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

// ---------------------------------------------------------------------------
// Synchronous variant (SHEET-1) — for `spreadsheetGrid.ts`'s
// `parseWorkbookBuffer`, the primary xlsx/ods read path (reached from the
// viewer, the parsing Worker, and PDF/CSV export).
// ---------------------------------------------------------------------------

const EOCD_SIGNATURE = 0x06054b50 // 'PK\x05\x06', little-endian as a uint32
const EOCD_FIXED_SIZE = 22 // signature through the comment-length field, before the variable-length comment
const EOCD_MAX_COMMENT_BYTES = 0xffff // a zip comment's length field is 16-bit
const CENTRAL_DIR_SIGNATURE = 0x02014b50 // 'PK\x01\x02'
const CENTRAL_DIR_FIXED_SIZE = 46 // signature through the relative-offset field, before name/extra/comment
const ZIP64_SENTINEL = 0xffffffff // the 32-bit "see the Zip64 extra field instead" marker

/** Searches backward for the End Of Central Directory record's signature, the standard way every zip reader locates it (it's always the LAST such record, but a trailing archive comment of up to 64KiB can follow it). `undefined` when none is found in that trailing window — not a zip this reader can make sense of. */
function findEndOfCentralDirectory(view: DataView): number | undefined {
  const searchWindow = Math.min(view.byteLength, EOCD_FIXED_SIZE + EOCD_MAX_COMMENT_BYTES)
  const highest = view.byteLength - EOCD_FIXED_SIZE
  const lowest = view.byteLength - searchWindow
  for (let offset = highest; offset >= lowest && offset >= 0; offset--) {
    if (view.getUint32(offset, true) === EOCD_SIGNATURE) return offset
  }
  return undefined
}

/**
 * Synchronous sibling of {@link loadWorkbookZip}, reading the SAME
 * information — each entry's declared uncompressed size, off the zip
 * central directory, before any inflation — directly off the raw bytes, no
 * JSZip instance and no `Promise` involved.
 *
 * Why a second implementation instead of just awaiting `loadWorkbookZip`:
 * `spreadsheetGrid.ts`'s `parseWorkbookBuffer` (the primary xlsx/ods read
 * path — the viewer, the parsing Worker, PDF/CSV export ALL go through it)
 * is called synchronously from many places across this codebase that
 * predate this guard; threading a `Promise` through every one of them is a
 * far bigger, separate change than this fix. Reading the central directory
 * by hand instead is no less "before a single byte is inflated" than
 * `loadWorkbookZip` — if anything more so, since it never even constructs a
 * JSZip entry object, just reads the fixed-offset size field straight off
 * the buffer.
 *
 * Deliberately narrower than `loadWorkbookZip`, on top of everything that
 * module's own header already documents as out of scope: NO Zip64 support.
 * An End Of Central Directory record that can't be found, or one whose
 * entry count or central-directory offset is pinned at its 32-bit sentinel
 * value (`0xFFFFFFFF`/`0xFFFF`, meaning "see the Zip64 extra field
 * instead"), means "can't determine this from the plain 32-bit fields" —
 * that archive (or, per-entry, that one entry's declared size) is skipped
 * rather than misread as literally 4GB. `XLSX.read` still runs afterwards
 * and reports its own error for a genuinely malformed file, exactly as it
 * did before this guard existed. A real Zip64 workbook is vanishingly rare
 * — that variant only exists for a >4GB archive or >65,535 entries, far
 * outside anything a spreadsheet produces — and skipping the check for one
 * leaves it exactly as safe as every xlsx/ods file was before this guard:
 * failing open on the format's edge case, not on the small-file/huge-
 * declared-size shape this guard exists to catch.
 */
export function checkWorkbookZipBudgetSync(
  buffer: ArrayBuffer | Uint8Array,
  maxEntryUncompressedBytes: number = DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES,
  maxTotalUncompressedBytes: number = DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES,
): void {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)

  const eocdOffset = findEndOfCentralDirectory(view)
  if (eocdOffset === undefined) return // no plain (non-Zip64) EOCD found — let XLSX.read report whatever is wrong with it

  const totalEntries = view.getUint16(eocdOffset + 10, true)
  const centralDirSize = view.getUint32(eocdOffset + 12, true)
  const centralDirOffset = view.getUint32(eocdOffset + 16, true)
  if (totalEntries === 0xffff || centralDirOffset === ZIP64_SENTINEL) return // Zip64 — see header

  const centralDirEnd = Math.min(bytes.byteLength, centralDirOffset + centralDirSize)
  let total = 0
  let pos = centralDirOffset

  for (let i = 0; i < totalEntries && pos + CENTRAL_DIR_FIXED_SIZE <= centralDirEnd; i++) {
    if (view.getUint32(pos, true) !== CENTRAL_DIR_SIGNATURE) break // truncated/malformed — stop, let XLSX.read report it

    const uncompressedSize = view.getUint32(pos + 24, true)
    const nameLength = view.getUint16(pos + 28, true)
    const extraLength = view.getUint16(pos + 30, true)
    const commentLength = view.getUint16(pos + 32, true)
    const nameStart = pos + CENTRAL_DIR_FIXED_SIZE
    // A directory entry's own name always ends in '/' (the same signal JSZip
    // uses); it declares no meaningful uncompressed size and is never the
    // zip-bomb attack shape, so it's excluded exactly like `loadWorkbookZip`
    // excludes `entry.dir`.
    const isDirectory = nameLength > 0 && bytes[nameStart + nameLength - 1] === 0x2f

    if (uncompressedSize !== ZIP64_SENTINEL && !isDirectory) {
      if (uncompressedSize > maxEntryUncompressedBytes) {
        throw new SpreadsheetZipBombError(
          `A zip entry declares ${uncompressedSize.toLocaleString()} bytes uncompressed, over the ` +
            `${maxEntryUncompressedBytes.toLocaleString()}-byte per-entry limit; refusing to open ` +
            '(possible zip bomb).',
        )
      }
      total += uncompressedSize
      if (total > maxTotalUncompressedBytes) {
        throw new SpreadsheetZipBombError(
          `Archive's combined declared uncompressed size is over the ` +
            `${maxTotalUncompressedBytes.toLocaleString()}-byte total limit; refusing to open ` +
            '(possible zip bomb).',
        )
      }
    }

    pos = nameStart + nameLength + extraLength + commentLength
  }
}
