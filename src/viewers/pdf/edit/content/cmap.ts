/**
 * CMap parsing (PDF-EDIT-2): `/ToUnicode` streams (code -> text) and the
 * embedded `/Encoding` CMaps of composite fonts (code bytes -> CID).
 */

import { tokenize, type ContentToken } from './lexer'

export type CodeSpaceRange = { readonly length: number; readonly low: number; readonly high: number }

export type ParsedCMap = {
  readonly codeSpace: ReadonlyArray<CodeSpaceRange>
  /** code -> Unicode text (bfchar/bfrange). */
  readonly toUnicode: ReadonlyMap<number, string>
  /** code -> CID (cidchar/cidrange). */
  readonly toCid: ReadonlyMap<number, number>
  /** Codes covered by cidrange entries too large to expand (rare). */
  readonly cidRanges: ReadonlyArray<{ readonly low: number; readonly high: number; readonly cid: number }>
  /** A `usecmap` of a predefined CMap this module cannot resolve. */
  readonly usesExternal: boolean
}

function bytesToNumber(bytes: string): number {
  let value = 0
  for (let i = 0; i < bytes.length; i += 1) value = value * 256 + bytes.charCodeAt(i)
  return value
}

/** A bfchar/bfrange destination (UTF-16BE bytes) as a JS string. */
function utf16be(bytes: string): string {
  let out = ''
  for (let i = 0; i + 1 < bytes.length; i += 2) out += String.fromCharCode((bytes.charCodeAt(i) << 8) | bytes.charCodeAt(i + 1))
  if (bytes.length % 2 === 1) out += bytes[bytes.length - 1]
  return out
}

function incrementLastChar(text: string, by: number): string {
  if (text.length === 0) return text
  return text.slice(0, -1) + String.fromCharCode(text.charCodeAt(text.length - 1) + by)
}

const MAX_EXPANDED_RANGE = 0x10000

export function parseCMap(src: string): ParsedCMap {
  const tokens = tokenize(src)
  const codeSpace: CodeSpaceRange[] = []
  const toUnicode = new Map<number, string>()
  const toCid = new Map<number, number>()
  const cidRanges: Array<{ low: number; high: number; cid: number }> = []
  let usesExternal = false

  const str = (token: ContentToken | undefined) => (token?.type === 'string' ? token.bytes : null)
  const num = (token: ContentToken | undefined) => (token?.type === 'number' ? token.value : null)

  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i]
    if (token.type !== 'keyword') continue
    switch (token.value) {
      case 'usecmap':
        usesExternal = true
        break
      case 'begincodespacerange':
        for (i += 1; i + 1 < tokens.length && !(tokens[i].type === 'keyword'); i += 2) {
          const low = str(tokens[i])
          const high = str(tokens[i + 1])
          if (low !== null && high !== null) codeSpace.push({ length: low.length, low: bytesToNumber(low), high: bytesToNumber(high) })
        }
        break
      case 'beginbfchar':
        for (i += 1; i + 1 < tokens.length && !(tokens[i].type === 'keyword'); i += 2) {
          const code = str(tokens[i])
          const dst = tokens[i + 1]
          if (code === null) continue
          if (dst.type === 'string') toUnicode.set(bytesToNumber(code), utf16be(dst.bytes))
          else if (dst.type === 'name') toUnicode.set(bytesToNumber(code), dst.value)
        }
        break
      case 'beginbfrange':
        for (i += 1; i + 2 < tokens.length && !(tokens[i].type === 'keyword'); i += 3) {
          const low = str(tokens[i])
          const high = str(tokens[i + 1])
          const dst = tokens[i + 2]
          if (low === null || high === null) continue
          const lo = bytesToNumber(low)
          const hi = Math.min(bytesToNumber(high), lo + MAX_EXPANDED_RANGE)
          if (dst.type === 'string') {
            const base = utf16be(dst.bytes)
            for (let code = lo; code <= hi; code += 1) toUnicode.set(code, incrementLastChar(base, code - lo))
          } else if (dst.type === 'array') {
            dst.items.forEach((item, offset) => {
              if (item.type === 'string' && lo + offset <= hi) toUnicode.set(lo + offset, utf16be(item.bytes))
            })
          }
        }
        break
      case 'begincidchar':
        for (i += 1; i + 1 < tokens.length && !(tokens[i].type === 'keyword'); i += 2) {
          const code = str(tokens[i])
          const cid = num(tokens[i + 1])
          if (code !== null && cid !== null) toCid.set(bytesToNumber(code), cid)
        }
        break
      case 'begincidrange':
        for (i += 1; i + 2 < tokens.length && !(tokens[i].type === 'keyword'); i += 3) {
          const low = str(tokens[i])
          const high = str(tokens[i + 1])
          const cid = num(tokens[i + 2])
          if (low === null || high === null || cid === null) continue
          cidRanges.push({ low: bytesToNumber(low), high: bytesToNumber(high), cid })
        }
        break
      default:
        break
    }
    // Each section loop above stops ON its closing keyword; the outer loop steps past it.
  }
  return { codeSpace, toUnicode, toCid, cidRanges, usesExternal }
}

/** CID for a code under an embedded encoding CMap, or null when unmapped. */
export function lookupCid(cmap: ParsedCMap, code: number): number | null {
  const direct = cmap.toCid.get(code)
  if (direct !== undefined) return direct
  for (const range of cmap.cidRanges) {
    if (code >= range.low && code <= range.high) return range.cid + (code - range.low)
  }
  return null
}

/**
 * Splits a shown string into codes under `codeSpace`: at each position, the
 * shortest byte length whose range contains the code wins (the rule in
 * PDF 32000-1 §9.7.6.2). Bytes that match no range are taken one at a time.
 */
export function splitCodes(bytes: string, codeSpace: ReadonlyArray<CodeSpaceRange>): Array<{ code: number; length: number }> {
  const out: Array<{ code: number; length: number }> = []
  let pos = 0
  while (pos < bytes.length) {
    let matched = false
    for (let length = 1; length <= 4 && pos + length <= bytes.length; length += 1) {
      const code = bytesToNumber(bytes.slice(pos, pos + length))
      if (codeSpace.some((range) => range.length === length && code >= range.low && code <= range.high)) {
        out.push({ code, length })
        pos += length
        matched = true
        break
      }
    }
    if (!matched) {
      out.push({ code: bytes.charCodeAt(pos), length: 1 })
      pos += 1
    }
  }
  return out
}
