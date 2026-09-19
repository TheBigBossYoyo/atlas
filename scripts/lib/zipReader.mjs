// Minimal, dependency-free ZIP reader used by the office-file validation
// harness (scripts/validate-office-file.mjs). Deliberately independent of
// JSZip (which Atlas's own writers use) and of any Atlas parsing code: it
// reads the raw PKZIP structures (End Of Central Directory, Central
// Directory records, Local File Headers) directly off the bytes, the same
// way a strict consumer would, so it can catch things a forgiving zip
// library would silently paper over — a non-stored `mimetype`, a backslash
// or leading-slash entry name, entries out of physical order, etc.
//
// Supports exactly what Office/ODF packages need: DEFLATE (method 8) and
// STORE (method 0), no Zip64, no encryption, no spanning. Throws a plain
// Error with a descriptive message for anything else — callers surface that
// as a validation issue rather than crashing.
import zlib from 'node:zlib'

const EOCD_SIGNATURE = 0x06054b50
const CENTRAL_DIR_SIGNATURE = 0x02014b50
const LOCAL_HEADER_SIGNATURE = 0x04034b50
const EOCD_MIN_SIZE = 22
const MAX_COMMENT_LENGTH = 0xffff

/**
 * @typedef {Object} ZipEntry
 * @property {string} name - Raw entry name exactly as stored (not normalized).
 * @property {number} method - 0 = stored, 8 = deflate.
 * @property {number} compressedSize
 * @property {number} uncompressedSize
 * @property {number} localHeaderOffset - Byte offset of the entry's local file header, i.e. its physical position in the archive.
 * @property {number} crc32
 * @property {boolean} isDirectory
 */

/**
 * Best-effort directory detection from a central-directory entry's external
 * file attributes, for archives that mark directories without a trailing
 * slash on the name. `externalAttrs` packs two different attribute schemes
 * depending on the "version made by" host — a Unix-mode `st_mode` in the
 * UPPER 16 bits (directory: `(mode & 0xF000) === 0x4000`, i.e. S_IFDIR) or
 * an MS-DOS attribute byte in the LOWER 8 bits (directory: bit 0x10, NOT
 * shifted). Checking the wrong bit range/shift silently misclassifies an
 * ordinary Unix-built regular file as a directory whenever its mode
 * happens to have that unrelated bit set (e.g. group-write, `0o664` has bit
 * 0x10 of the low byte set) — exactly what an earlier version of this
 * function did, which made this reader falsely mark every part of a
 * Unix-zipped `.docx` (mode `100644` = `0x81a4`) as a directory and treat
 * the archive as effectively empty.
 */
function isDirectoryFromExternalAttrs(externalAttrs) {
  const unixMode = (externalAttrs >>> 16) & 0xffff
  if (unixMode !== 0) return (unixMode & 0xf000) === 0x4000
  return (externalAttrs & 0x10) !== 0
}

function findEndOfCentralDirectory(buffer) {
  const maxScan = Math.min(buffer.length, EOCD_MIN_SIZE + MAX_COMMENT_LENGTH)
  for (let offset = buffer.length - EOCD_MIN_SIZE; offset >= buffer.length - maxScan; offset--) {
    if (offset < 0) break
    if (buffer.readUInt32LE(offset) === EOCD_SIGNATURE) return offset
  }
  throw new Error('Not a valid zip archive: no End Of Central Directory record found.')
}

/**
 * Reads every central-directory entry, then returns them sorted by their
 * local-header byte offset — i.e. the ARCHIVE'S PHYSICAL ORDER, which is
 * what matters for ODF's "mimetype must be the first file" rule and is a
 * more honest signal than central-directory order (a conforming writer's
 * central directory always matches physical order, but nothing enforces
 * that a reader may rely on it).
 *
 * @param {Buffer} buffer
 * @returns {ZipEntry[]}
 */
export function readZipEntries(buffer) {
  const eocdOffset = findEndOfCentralDirectory(buffer)
  const totalEntries = buffer.readUInt16LE(eocdOffset + 10)
  const centralDirOffset = buffer.readUInt32LE(eocdOffset + 16)

  if (centralDirOffset === 0xffffffff || totalEntries === 0xffff) {
    throw new Error('Zip64 archives are not supported by this validation harness.')
  }

  const entries = []
  let pointer = centralDirOffset
  for (let i = 0; i < totalEntries; i++) {
    if (buffer.readUInt32LE(pointer) !== CENTRAL_DIR_SIGNATURE) {
      throw new Error(`Malformed central directory: expected entry signature at byte ${pointer}.`)
    }
    const method = buffer.readUInt16LE(pointer + 10)
    const crc32 = buffer.readUInt32LE(pointer + 16)
    const compressedSize = buffer.readUInt32LE(pointer + 20)
    const uncompressedSize = buffer.readUInt32LE(pointer + 24)
    const nameLen = buffer.readUInt16LE(pointer + 28)
    const extraLen = buffer.readUInt16LE(pointer + 30)
    const commentLen = buffer.readUInt16LE(pointer + 32)
    const externalAttrs = buffer.readUInt32LE(pointer + 38)
    const localHeaderOffset = buffer.readUInt32LE(pointer + 42)
    const name = buffer.toString('utf8', pointer + 46, pointer + 46 + nameLen)

    entries.push({
      name,
      method,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
      crc32,
      isDirectory: name.endsWith('/') || isDirectoryFromExternalAttrs(externalAttrs),
    })

    pointer += 46 + nameLen + extraLen + commentLen
  }

  return entries.sort((a, b) => a.localHeaderOffset - b.localHeaderOffset)
}

/**
 * Decompresses one entry's data, reading the actual byte span from its
 * LOCAL file header (name/extra lengths there can legitimately differ from
 * the central directory's copy) rather than trusting the central
 * directory's size fields alone for the offset math.
 *
 * @param {Buffer} buffer
 * @param {ZipEntry} entry
 * @returns {Buffer}
 */
export function readZipEntryData(buffer, entry) {
  const offset = entry.localHeaderOffset
  if (buffer.readUInt32LE(offset) !== LOCAL_HEADER_SIGNATURE) {
    throw new Error(`Malformed local file header for "${entry.name}" at byte ${offset}.`)
  }
  const nameLen = buffer.readUInt16LE(offset + 26)
  const extraLen = buffer.readUInt16LE(offset + 28)
  const dataStart = offset + 30 + nameLen + extraLen
  const dataEnd = dataStart + entry.compressedSize
  const raw = buffer.subarray(dataStart, dataEnd)

  if (entry.method === 0) return Buffer.from(raw)
  if (entry.method === 8) return zlib.inflateRawSync(raw)
  throw new Error(`"${entry.name}" uses unsupported zip compression method ${entry.method}.`)
}

/**
 * Convenience: reads the whole archive into a `Map<name, Buffer>` plus the
 * ordered entry list (physical order), skipping directory entries.
 *
 * @param {Buffer} buffer
 * @returns {{ entries: ZipEntry[], files: Map<string, Buffer> }}
 */
export function readZipArchive(buffer) {
  const entries = readZipEntries(buffer)
  const files = new Map()
  for (const entry of entries) {
    if (entry.isDirectory) continue
    files.set(entry.name, readZipEntryData(buffer, entry))
  }
  return { entries, files }
}
