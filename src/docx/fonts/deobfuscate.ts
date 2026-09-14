/**
 * Atlas — DOCX embedded-font de-obfuscation (DEFER-4 / DXP-13)
 *
 * Word's "Embed fonts in the file" option obfuscates each embedded font
 * part's first 32 bytes to discourage casual font-file extraction, per
 * ECMA-376 Part 1 §17.8.1 ("Font Obfuscation"). `word/fontTable.xml`'s
 * `w:embedRegular`/`w:embedBold`/`w:embedItalic`/`w:embedBoldItalic`
 * elements each carry a `w:fontKey` GUID; that GUID, reinterpreted as a
 * 16-byte key using the standard Windows/.NET `GUID` struct layout, is
 * XORed against the font part's first 32 bytes (the 16-byte key applied
 * twice, once per 16-byte half) to recover the original TrueType/OpenType
 * font data.
 *
 * The GUID layout matters because a GUID's *string* form
 * (`{D1D1D1D1-D2D2-D3D3-D4D4-D4D4D4D4D4D4}`) prints `Data1`/`Data2`/`Data3`
 * big-endian, but the struct stores them little-endian in memory — the
 * bytes Word actually XORs against the font are the in-memory (little-
 * endian) ones, not a naive left-to-right reading of the hex string. `Data4`
 * (the last 8 bytes) has no such reversal in either form.
 *
 * XOR is its own inverse, so `xorObfuscatedFontHeader` both de-obfuscates a
 * loaded embedded font part and (used the same way) obfuscates a plain font
 * file — `__tests__/embedded.test.ts` uses that symmetry to build a
 * synthetic obfuscated fixture from a bundled OFL TTF without needing a
 * real Word-authored embedded-font document.
 */

export class FontKeyError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'FontKeyError'
  }
}

const GUID_PATTERN =
  /^\{?([0-9a-fA-F]{8})-([0-9a-fA-F]{4})-([0-9a-fA-F]{4})-([0-9a-fA-F]{4})-([0-9a-fA-F]{12})\}?$/

/** Only the first 32 bytes of an embedded font part are obfuscated. */
const OBFUSCATED_HEADER_LENGTH = 32
/** The GUID-derived key is 16 bytes, applied twice across the 32-byte header. */
const FONT_KEY_LENGTH = 16

function hexByteAt(hex: string, charOffset: number): number {
  return Number.parseInt(hex.substring(charOffset, charOffset + 2), 16)
}

/** Reverses byte order within a hex group (little-endian struct storage). */
function hexGroupToBytesLE(hex: string): number[] {
  const bytes: number[] = []
  for (let charOffset = hex.length - 2; charOffset >= 0; charOffset -= 2) {
    bytes.push(hexByteAt(hex, charOffset))
  }
  return bytes
}

/** Keeps byte order within a hex group as written (big-endian / as-is). */
function hexGroupToBytesAsIs(hex: string): number[] {
  const bytes: number[] = []
  for (let charOffset = 0; charOffset < hex.length; charOffset += 2) {
    bytes.push(hexByteAt(hex, charOffset))
  }
  return bytes
}

/**
 * Converts a `w:fontKey` GUID string (braces and hyphens both optional in
 * input, though Word always writes the full braced form) into the 16-byte
 * array actually XORed against the obfuscated font header.
 *
 * @throws `FontKeyError` when `guid` isn't a well-formed GUID string.
 */
export function guidToFontKeyBytes(guid: string): Uint8Array {
  const match = GUID_PATTERN.exec(guid.trim())
  if (match === null) {
    throw new FontKeyError(`Invalid font obfuscation key GUID: "${guid}"`)
  }

  const [, data1, data2, data3, data4High, data4Low] = match
  const bytes = [
    ...hexGroupToBytesLE(data1),
    ...hexGroupToBytesLE(data2),
    ...hexGroupToBytesLE(data3),
    ...hexGroupToBytesAsIs(data4High),
    ...hexGroupToBytesAsIs(data4Low),
  ]

  return Uint8Array.from(bytes)
}

/**
 * XORs the first 32 bytes of `data` against `fontKeyGuid`'s derived 16-byte
 * key (repeated twice), leaving the rest of the buffer untouched. Returns a
 * new `Uint8Array` — `data` is never mutated. Self-inverse: applying this
 * twice with the same key is the identity transform.
 *
 * @throws `FontKeyError` when `fontKeyGuid` isn't a well-formed GUID string.
 */
export function xorObfuscatedFontHeader(data: Uint8Array, fontKeyGuid: string): Uint8Array {
  const key = guidToFontKeyBytes(fontKeyGuid)
  const result = Uint8Array.from(data)
  const headerLength = Math.min(OBFUSCATED_HEADER_LENGTH, result.length)

  for (let index = 0; index < headerLength; index += 1) {
    result[index] ^= key[index % FONT_KEY_LENGTH]
  }

  return result
}
