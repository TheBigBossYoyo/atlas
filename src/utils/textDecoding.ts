// Renderer-side counterpart to `electron/lib/textDecoding.cjs` (P1.11 /
// LOAD-05, DAT-03, RUN-05). Used by `useFileHandler`'s unknown-extension
// fallback, where the raw bytes arrive via IPC as an ArrayBuffer rather than
// a Node Buffer, so the decoding logic — while conceptually identical — is
// re-implemented here against `Uint8Array`/`TextDecoder` instead of Node's
// `Buffer` API.
//
// Order of operations:
//   1. UTF-8 BOM (EF BB BF)    -> decode the rest as UTF-8, BOM stripped.
//   2. UTF-16 LE BOM (FF FE)   -> decode the rest as UTF-16 LE.
//   3. UTF-16 BE BOM (FE FF)   -> decode the rest as UTF-16 BE.
//   4. No BOM, valid UTF-8     -> decode as UTF-8 (byte-for-byte identical to
//      today's `new TextDecoder('utf-8').decode(...)` for the overwhelmingly
//      common no-BOM-valid-UTF-8 case).
//   5. No BOM, invalid UTF-8   -> fall back to Windows-1252 instead of
//      emitting U+FFFD replacement-character mojibake.

// Windows-1252 code points for bytes 0x80-0x9F (0xA0-0xFF and 0x00-0x7F are
// identical to their Unicode code points). The five bytes with no assigned
// Windows-1252 character (0x81, 0x8D, 0x8F, 0x90, 0x9D) fall back to their
// raw byte value (Latin-1 identity) rather than throwing, since this is
// already a best-effort fallback path.
const WINDOWS_1252_HIGH_RANGE: ReadonlyArray<number> = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039,
  0x0152, 0x008d, 0x017d, 0x008f, 0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
]

export function isValidUtf8(bytes: Uint8Array): boolean {
  const len = bytes.length
  let i = 0

  while (i < len) {
    const byte1 = bytes[i]

    if (byte1 <= 0x7f) {
      i += 1
      continue
    }

    let extraBytes: number
    let codePoint: number
    let minCodePoint: number

    if ((byte1 & 0xe0) === 0xc0) {
      extraBytes = 1
      codePoint = byte1 & 0x1f
      minCodePoint = 0x80
    } else if ((byte1 & 0xf0) === 0xe0) {
      extraBytes = 2
      codePoint = byte1 & 0x0f
      minCodePoint = 0x800
    } else if ((byte1 & 0xf8) === 0xf0) {
      extraBytes = 3
      codePoint = byte1 & 0x07
      minCodePoint = 0x10000
    } else {
      return false
    }

    if (i + extraBytes >= len) {
      return false
    }

    for (let j = 1; j <= extraBytes; j += 1) {
      const continuationByte = bytes[i + j]
      if ((continuationByte & 0xc0) !== 0x80) {
        return false
      }
      codePoint = (codePoint << 6) | (continuationByte & 0x3f)
    }

    if (codePoint < minCodePoint || codePoint > 0x10ffff) {
      return false
    }
    if (codePoint >= 0xd800 && codePoint <= 0xdfff) {
      return false
    }

    i += extraBytes + 1
  }

  return true
}

export function decodeWindows1252(bytes: Uint8Array): string {
  let result = ''
  for (let i = 0; i < bytes.length; i += 1) {
    const byte = bytes[i]
    if (byte >= 0x80 && byte <= 0x9f) {
      result += String.fromCharCode(WINDOWS_1252_HIGH_RANGE[byte - 0x80])
    } else {
      result += String.fromCharCode(byte)
    }
  }
  return result
}

/**
 * Decodes a text-class file's raw bytes, sniffing a BOM and falling back to
 * Windows-1252 for non-UTF-8 content with no BOM.
 */
export function decodeTextBuffer(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer)

  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3))
  }

  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2))
  }

  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(bytes.subarray(2))
  }

  if (isValidUtf8(bytes)) {
    return new TextDecoder('utf-8').decode(bytes)
  }

  return decodeWindows1252(bytes)
}

// ---------------------------------------------------------------------------
// NIGHT/text-roundtrip — round-trip fidelity metadata (renderer-side mirror
// of `electron/lib/textDecoding.cjs`'s additions). See that file's header
// for the full rationale (SHELL-1/SHELL-2): every save used to silently
// re-encode as BOM-less UTF-8 with whatever newlines the DOM editor
// normalized to, discarding the source file's actual encoding/BOM/newline
// convention. `decodeTextBufferWithMeta` below is used by the
// `prefetchedBuffer` decode path in `useFileHandler.ts` (the Open dialog
// already has the bytes in the renderer, so decoding happens here rather
// than round-tripping through main a second time); the registry lets a save
// (in App.tsx/CodeViewer.tsx, run long after the original decode) look the
// file's convention back up by path. `LoadedFile` itself
// (`src/formats/types.ts`) is not owned by this fix and carries no such
// field, so the registry is the fallback: metadata keyed by path instead of
// carried on the loaded-file object.
// ---------------------------------------------------------------------------

export type TextEncodingName = 'utf-8' | 'utf-16le' | 'utf-16be' | 'windows-1252'
export type NewlineStyle = 'crlf' | 'lf'

export interface TextFileMeta {
  readonly encoding: TextEncodingName
  readonly bom: boolean
  readonly newline: NewlineStyle
}

/** A brand-new/untracked document keeps today's defaults: plain UTF-8, no BOM, LF. */
export const DEFAULT_TEXT_FILE_META: TextFileMeta = { encoding: 'utf-8', bom: false, newline: 'lf' }

const textFileMetaRegistry = new Map<string, TextFileMeta>()

/** Records `path`'s source encoding/BOM/newline so a later save can reapply it. */
export function setTextFileMeta(path: string, meta: TextFileMeta): void {
  textFileMetaRegistry.set(path, meta)
}

/** Looks up `path`'s recorded convention, or today's-defaults for an untracked/new path. */
export function getTextFileMeta(path: string): TextFileMeta {
  return textFileMetaRegistry.get(path) ?? DEFAULT_TEXT_FILE_META
}

/** Save As / rename: `newPath` should keep reapplying `oldPath`'s convention on every later save. A no-op if `oldPath` was never tracked (new document, saved for the first time). */
export function carryTextFileMeta(oldPath: string, newPath: string): void {
  const meta = textFileMetaRegistry.get(oldPath)
  if (meta !== undefined && oldPath !== newPath) {
    textFileMetaRegistry.set(newPath, meta)
  }
}

/** Normalizes every line ending to `\n` (CRLF and lone CR both collapse to LF) — matches what a `<textarea>`/CodeMirror already does to anything typed or pasted into it. */
export function normalizeNewlines(content: string): string {
  return content.replace(/\r\n|\r/g, '\n')
}

/** Detects the dominant line-ending convention of already-decoded text. See the `.cjs` twin's doc comment for the mixed-file tie-break rule. */
export function detectNewline(content: string): NewlineStyle {
  let crlf = 0
  let lf = 0
  for (let i = 0; i < content.length; i += 1) {
    if (content[i] === '\n') {
      if (i > 0 && content[i - 1] === '\r') {
        crlf += 1
      } else {
        lf += 1
      }
    }
  }
  if (crlf === 0) return 'lf'
  if (lf === 0) return 'crlf'
  return crlf > lf ? 'crlf' : 'lf'
}

/**
 * Decodes a text-class file's raw bytes AND detects its round-trip metadata.
 * `content` is always `\n`-normalized; `meta.newline` is detected from the
 * ORIGINAL bytes, before normalization. Mirrors `electron/lib/
 * textDecoding.cjs`'s `decodeTextBufferWithMeta` for the renderer's own
 * `prefetchedBuffer` decode path (`useFileHandler.ts`).
 */
export function decodeTextBufferWithMeta(buffer: ArrayBuffer): { content: string; meta: TextFileMeta } {
  const bytes = new Uint8Array(buffer)
  let encoding: TextEncodingName
  let bom: boolean
  let decoded: string

  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    encoding = 'utf-8'
    bom = true
    decoded = new TextDecoder('utf-8').decode(bytes.subarray(3))
  } else if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = 'utf-16le'
    bom = true
    decoded = new TextDecoder('utf-16le').decode(bytes.subarray(2))
  } else if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    encoding = 'utf-16be'
    bom = true
    decoded = new TextDecoder('utf-16be').decode(bytes.subarray(2))
  } else if (isValidUtf8(bytes)) {
    encoding = 'utf-8'
    bom = false
    decoded = new TextDecoder('utf-8').decode(bytes)
  } else {
    encoding = 'windows-1252'
    bom = false
    decoded = decodeWindows1252(bytes)
  }

  return { content: normalizeNewlines(decoded), meta: { encoding, bom, newline: detectNewline(decoded) } }
}
