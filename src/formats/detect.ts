import type { FormatId } from './types'
import { EXTENSION_TO_FORMAT } from './extensionManifest'
import { isValidUtf8 } from '../utils/textDecoding'

// P2.2/ELEC-05/ELEC-15/LOAD-03/LOAD-12 — this used to be a hand-maintained
// literal that disagreed with electron/main.cjs's known-extensions set and
// electron-builder.yml's file associations. All three now derive from the
// single canonical table in `extensionManifest.ts`.
const EXTENSION_MAP: Readonly<Record<string, FormatId>> = EXTENSION_TO_FORMAT

function getExtension(path: string): string | null {
  const fileName = path.split(/[\\/]/).pop() ?? path
  const dotIndex = fileName.lastIndexOf('.')

  if (dotIndex <= 0 || dotIndex === fileName.length - 1) {
    return null
  }

  return fileName.slice(dotIndex + 1).toLowerCase()
}

// P4.10/LOAD-15 — a ZIP-format file (docx/xlsx/pptx/odt/ods/odp) can be
// large, and the old implementation decoded the ENTIRE buffer to a JS string
// one character at a time just to substring-search it. The filename markers
// we look for live in early local-file-header entries ("[Content_Types].xml"
// / "_rels/.rels" / the "mimetype" store are conventionally first) or in the
// central directory at the very end of the archive, so scanning a bounded
// head + tail window finds them without ever materializing a multi-hundred-
// -megabyte string.
const ZIP_SCAN_HEAD_BYTES = 64 * 1024
const ZIP_SCAN_TAIL_BYTES = 16 * 1024

function decodeAsciiWindow(buffer: ArrayBuffer, start: number, end: number): string {
  const bytes = new Uint8Array(buffer, start, end - start)
  let decoded = ''

  for (const byte of bytes) {
    decoded += String.fromCharCode(byte)
  }

  return decoded
}

function scanZipMarkers(buffer: ArrayBuffer): string {
  const length = buffer.byteLength
  const headEnd = Math.min(length, ZIP_SCAN_HEAD_BYTES)
  const head = decodeAsciiWindow(buffer, 0, headEnd)

  if (length <= headEnd) {
    return head
  }

  const tailStart = Math.max(headEnd, length - ZIP_SCAN_TAIL_BYTES)
  return head + decodeAsciiWindow(buffer, tailStart, length)
}

function detectByMagicMarker(buffer: ArrayBuffer): FormatId {
  const bytes = new Uint8Array(buffer)

  if (bytes.length >= 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) {
    return 'pdf'
  }

  if (bytes.length >= 5 && bytes[0] === 0x7b && bytes[1] === 0x5c && bytes[2] === 0x72 && bytes[3] === 0x74 && bytes[4] === 0x66) {
    return 'rtf'
  }

  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) {
    const ascii = scanZipMarkers(buffer)

    if (ascii.includes('word/document.xml')) return 'docx'
    if (ascii.includes('xl/workbook.xml')) return 'xlsx'
    if (ascii.includes('ppt/presentation.xml')) return 'pptx'
    if (ascii.includes('mimetypeapplication/vnd.oasis.opendocument.text')) return 'odt'
    if (ascii.includes('mimetypeapplication/vnd.oasis.opendocument.spreadsheet')) return 'ods'
    if (ascii.includes('mimetypeapplication/vnd.oasis.opendocument.presentation')) return 'odp'
  }

  return 'unknown'
}

export function detectByExtension(path: string): FormatId {
  const extension = getExtension(path)

  if (!extension) {
    return 'unknown'
  }

  return EXTENSION_MAP[extension] ?? 'unknown'
}

export function detectByMagic(buffer: ArrayBuffer): FormatId {
  return detectByMagicMarker(buffer)
}

export function detectFormat(path: string, buffer?: ArrayBuffer): FormatId {
  const extensionFormat = detectByExtension(path)

  if (!buffer) {
    return extensionFormat
  }

  const magicFormat = detectByMagic(buffer)

  if (magicFormat === 'unknown') {
    return extensionFormat
  }

  if (extensionFormat === 'text' || extensionFormat === 'code') {
    return extensionFormat
  }

  if (extensionFormat !== 'unknown' && extensionFormat === magicFormat) {
    return extensionFormat
  }

  return magicFormat
}

// P2.11/LOAD-10 — a cheap heuristic for "is this probably plain text" when
// both extension and magic-byte detection came up empty (an extensionless
// README, LICENSE, Dockerfile, or .gitignore-style file). Only the leading
// sample is checked: large enough to reliably distinguish text from binary,
// small enough to never decode a huge file just to answer a yes/no question.
const TEXT_SNIFF_SAMPLE_BYTES = 8192

/**
 * When the sample is a truncated prefix of a larger buffer, its very last
 * byte(s) may be the lead byte of a multi-byte UTF-8 sequence whose
 * continuation bytes fall just past the cutoff — which would otherwise read
 * as invalid UTF-8 and misclassify an otherwise-plain-text file as binary.
 * Trims that dangling partial sequence (at most 3 bytes) before validation.
 */
function trimIncompleteUtf8Tail(bytes: Uint8Array): Uint8Array {
  const len = bytes.length
  const maxLookback = Math.min(4, len)

  for (let back = 1; back <= maxLookback; back += 1) {
    const byte = bytes[len - back]
    if ((byte & 0xc0) === 0x80) {
      continue // continuation byte — keep walking backwards
    }

    const sequenceLength =
      byte <= 0x7f ? 1
      : (byte & 0xe0) === 0xc0 ? 2
      : (byte & 0xf0) === 0xe0 ? 3
      : (byte & 0xf8) === 0xf0 ? 4
      : -1 // not a valid UTF-8 lead byte — leave as-is, isValidUtf8 rejects it anyway

    if (sequenceLength !== -1 && back < sequenceLength) {
      // A multi-byte sequence starts here but the sample was cut off before
      // its continuation bytes arrived — drop the dangling partial sequence.
      return bytes.subarray(0, len - back)
    }
    return bytes
  }

  return bytes
}

export function looksLikeText(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength === 0) {
    return false
  }

  const sampleLength = Math.min(buffer.byteLength, TEXT_SNIFF_SAMPLE_BYTES)
  const rawSample = new Uint8Array(buffer, 0, sampleLength)

  for (const byte of rawSample) {
    if (byte === 0x00) {
      return false
    }
  }

  const isTruncatedSample = sampleLength < buffer.byteLength
  const bytes = isTruncatedSample ? trimIncompleteUtf8Tail(rawSample) : rawSample

  return isValidUtf8(bytes)
}
