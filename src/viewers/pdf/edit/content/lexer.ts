/**
 * A PDF content-stream tokenizer that keeps byte offsets (PDF-EDIT-2).
 *
 * Real text editing rewrites the page's drawing instructions, and the rule
 * that keeps that safe is: change only the operators being edited, and copy
 * every other byte through untouched. So this parses the stream into
 * operators that each know the exact byte span they came from, and the
 * rewriter splices replacements into those spans.
 *
 * Also used for CMap streams, which share the same token syntax.
 *
 * Bytes are handled as a "binary string" (one char per byte, codes 0-255), so
 * offsets in the string are offsets in the stream.
 */

export type ContentToken =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'name'; readonly value: string }
  | { readonly type: 'string'; readonly bytes: string; readonly hex: boolean }
  | { readonly type: 'array'; readonly items: ReadonlyArray<ContentToken> }
  | { readonly type: 'dict'; readonly entries: ReadonlyMap<string, ContentToken> }
  | { readonly type: 'keyword'; readonly value: string }

export type ContentOp = {
  readonly operator: string
  readonly operands: ReadonlyArray<ContentToken>
  /** Byte offset of the first operand (or of the operator when it has none). */
  readonly start: number
  /** Byte offset just past the operator. */
  readonly end: number
}

const WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20])
const DELIMITERS = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25])

export function bytesToBinaryString(bytes: Uint8Array): string {
  let out = ''
  const CHUNK = 0x8000
  for (let i = 0; i < bytes.length; i += CHUNK) {
    out += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  }
  return out
}

export function binaryStringToBytes(text: string): Uint8Array {
  const out = new Uint8Array(text.length)
  for (let i = 0; i < text.length; i += 1) out[i] = text.charCodeAt(i) & 0xff
  return out
}

type Positioned = { readonly token: ContentToken; readonly start: number; readonly end: number }

class Tokenizer {
  pos = 0
  readonly src: string

  constructor(src: string) {
    this.src = src
  }

  private code(at: number): number {
    return this.src.charCodeAt(at)
  }

  skipWhitespaceAndComments(): void {
    const { src } = this
    while (this.pos < src.length) {
      const c = this.code(this.pos)
      if (WHITESPACE.has(c)) {
        this.pos += 1
      } else if (c === 0x25) {
        while (this.pos < src.length && this.code(this.pos) !== 0x0a && this.code(this.pos) !== 0x0d) this.pos += 1
      } else {
        return
      }
    }
  }

  /** The next token, or null at the end; `]`, `>>` and `}` come back as keywords for the caller to close on. */
  next(): Positioned | null {
    this.skipWhitespaceAndComments()
    const { src } = this
    if (this.pos >= src.length) return null
    const start = this.pos
    const c = this.code(start)
    const ch = src[start]

    if (ch === '(') return { token: this.literalString(), start, end: this.pos }
    if (ch === '<') {
      if (src[start + 1] === '<') {
        this.pos += 2
        return { token: this.dict(), start, end: this.pos }
      }
      return { token: this.hexString(), start, end: this.pos }
    }
    if (ch === '>' && src[start + 1] === '>') {
      this.pos += 2
      return { token: { type: 'keyword', value: '>>' }, start, end: this.pos }
    }
    if (ch === '[') {
      this.pos += 1
      return { token: this.array(), start, end: this.pos }
    }
    if (ch === ']' || ch === '{' || ch === '}' || ch === ')' || ch === '>') {
      this.pos += 1
      return { token: { type: 'keyword', value: ch }, start, end: this.pos }
    }
    if (ch === '/') {
      this.pos += 1
      return { token: { type: 'name', value: this.name() }, start, end: this.pos }
    }
    // A regular-character run: a number or a keyword.
    let end = start
    while (end < src.length && !WHITESPACE.has(this.code(end)) && !DELIMITERS.has(this.code(end))) end += 1
    if (end === start) {
      // A lone delimiter we don't otherwise handle; step over it.
      this.pos += 1
      return { token: { type: 'keyword', value: String.fromCharCode(c) }, start, end: this.pos }
    }
    this.pos = end
    const word = src.slice(start, end)
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) return { token: { type: 'number', value: Number(word) }, start, end }
    // Malformed numbers such as "--3" or "3-" occur in the wild; readers take the number.
    if (/^[+-]*(\d+\.?\d*|\.\d+)-?$/.test(word)) {
      const digits = word.replace(/^[+-]*/, '').replace(/-$/, '')
      const negative = /^[+-]*-/.test(word)
      return { token: { type: 'number', value: Number(digits) * (negative ? -1 : 1) }, start, end }
    }
    return { token: { type: 'keyword', value: word }, start, end }
  }

  private name(): string {
    const { src } = this
    let out = ''
    while (this.pos < src.length) {
      const c = this.code(this.pos)
      if (WHITESPACE.has(c) || DELIMITERS.has(c)) break
      if (c === 0x23 && /^[0-9a-fA-F]{2}$/.test(src.slice(this.pos + 1, this.pos + 3))) {
        out += String.fromCharCode(parseInt(src.slice(this.pos + 1, this.pos + 3), 16))
        this.pos += 3
      } else {
        out += src[this.pos]
        this.pos += 1
      }
    }
    return out
  }

  private literalString(): ContentToken {
    const { src } = this
    this.pos += 1
    let depth = 1
    let out = ''
    while (this.pos < src.length) {
      const ch = src[this.pos]
      if (ch === '\\') {
        const next = src[this.pos + 1]
        this.pos += 2
        switch (next) {
          case 'n': out += '\n'; break
          case 'r': out += '\r'; break
          case 't': out += '\t'; break
          case 'b': out += '\b'; break
          case 'f': out += '\f'; break
          case '(': case ')': case '\\': out += next; break
          case '\r':
            if (src[this.pos] === '\n') this.pos += 1
            break
          case '\n':
            break
          default:
            if (next !== undefined && next >= '0' && next <= '7') {
              let digits = next
              while (digits.length < 3 && src[this.pos] >= '0' && src[this.pos] <= '7') {
                digits += src[this.pos]
                this.pos += 1
              }
              out += String.fromCharCode(parseInt(digits, 8) & 0xff)
            } else if (next !== undefined) {
              out += next
            }
        }
        continue
      }
      if (ch === '\r') {
        // A raw CR or CRLF inside a literal string reads as one LF (PDF 32000-1 §7.3.4.2).
        out += '\n'
        this.pos += src[this.pos + 1] === '\n' ? 2 : 1
        continue
      }
      if (ch === '(') depth += 1
      if (ch === ')') {
        depth -= 1
        if (depth === 0) {
          this.pos += 1
          break
        }
      }
      out += ch
      this.pos += 1
    }
    return { type: 'string', bytes: out, hex: false }
  }

  private hexString(): ContentToken {
    const { src } = this
    this.pos += 1
    let digits = ''
    while (this.pos < src.length && src[this.pos] !== '>') {
      if (/[0-9a-fA-F]/.test(src[this.pos])) digits += src[this.pos]
      this.pos += 1
    }
    this.pos += 1
    if (digits.length % 2 === 1) digits += '0'
    let out = ''
    for (let i = 0; i < digits.length; i += 2) out += String.fromCharCode(parseInt(digits.slice(i, i + 2), 16))
    return { type: 'string', bytes: out, hex: true }
  }

  private array(): ContentToken {
    const items: ContentToken[] = []
    for (;;) {
      const next = this.next()
      if (!next) break
      if (next.token.type === 'keyword' && next.token.value === ']') break
      items.push(next.token)
    }
    return { type: 'array', items }
  }

  private dict(): ContentToken {
    const entries = new Map<string, ContentToken>()
    let key: string | null = null
    for (;;) {
      const next = this.next()
      if (!next) break
      if (next.token.type === 'keyword' && next.token.value === '>>') break
      if (key === null) {
        if (next.token.type === 'name') key = next.token.value
      } else {
        entries.set(key, next.token)
        key = null
      }
    }
    return { type: 'dict', entries }
  }

  /**
   * Skips an inline image's data after its `ID` keyword: binary bytes that
   * may contain anything, ended by `EI` between whitespace. `/L` (`/Length`)
   * is used when present; otherwise the first plausible `EI` ends it.
   */
  skipInlineImageData(length: number | null): void {
    const { src } = this
    // Exactly one whitespace byte follows ID.
    this.pos += 1
    if (length !== null && length >= 0) {
      const after = this.pos + length
      const tail = /^\s*EI(?=[\s]|$)/.exec(src.slice(after, after + 32))
      if (tail) {
        this.pos = after + tail[0].length
        return
      }
    }
    const pattern = /[\s]EI(?=[\s]|$)/g
    pattern.lastIndex = this.pos
    const match = pattern.exec(src)
    this.pos = match ? match.index + match[0].length : src.length
  }
}

/** Every token of a stream, flat (keywords included) — what the CMap parser walks. */
export function tokenize(src: string): ContentToken[] {
  const tokenizer = new Tokenizer(src)
  const tokens: ContentToken[] = []
  for (;;) {
    const next = tokenizer.next()
    if (!next) break
    tokens.push(next.token)
  }
  return tokens
}

/** Parses a content stream into operators with their byte spans. */
export function parseContentStream(src: string): ContentOp[] {
  const tokenizer = new Tokenizer(src)
  const ops: ContentOp[] = []
  let operands: ContentToken[] = []
  let operandStart = -1
  for (;;) {
    const next = tokenizer.next()
    if (!next) break
    const { token, start, end } = next
    if (token.type !== 'keyword') {
      if (operands.length === 0) operandStart = start
      operands.push(token)
      continue
    }
    if (token.value === 'BI') {
      // Inline image: key/value pairs up to ID, then raw data up to EI — one op.
      const entries = new Map<string, ContentToken>()
      let key: string | null = null
      let lengthValue: number | null = null
      for (;;) {
        const item = tokenizer.next()
        if (!item) break
        if (item.token.type === 'keyword' && item.token.value === 'ID') break
        if (key === null && item.token.type === 'name') key = item.token.value
        else if (key !== null) {
          entries.set(key, item.token)
          if ((key === 'L' || key === 'Length') && item.token.type === 'number') lengthValue = item.token.value
          key = null
        }
      }
      tokenizer.skipInlineImageData(lengthValue)
      ops.push({ operator: 'BI', operands: [{ type: 'dict', entries }], start, end: tokenizer.pos })
      operands = []
      continue
    }
    ops.push({ operator: token.value, operands, start: operands.length > 0 ? operandStart : start, end })
    operands = []
  }
  return ops
}

// ---- Writing --------------------------------------------------------------------

/** A number as content-stream syntax: no exponent, at most 4 decimals, no trailing zeros. */
export function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return '0'
  const rounded = Math.round(value * 10000) / 10000
  if (Object.is(rounded, -0) || rounded === 0) return '0'
  return rounded.toFixed(4).replace(/\.?0+$/, '')
}

export function formatHexString(bytes: ReadonlyArray<number>): string {
  return `<${bytes.map((byte) => byte.toString(16).padStart(2, '0')).join('')}>`
}

export function formatName(name: string): string {
  let out = '/'
  for (const ch of name) {
    const code = ch.charCodeAt(0)
    if (code < 0x21 || code > 0x7e || ch === '#' || DELIMITERS.has(code)) out += `#${code.toString(16).padStart(2, '0')}`
    else out += ch
  }
  return out
}

/** Re-serialises a token (used to copy operands such as `"`'s spacing values verbatim). */
export function formatToken(token: ContentToken): string {
  switch (token.type) {
    case 'number':
      return formatNumber(token.value)
    case 'name':
      return formatName(token.value)
    case 'string':
      return formatHexString(Array.from(token.bytes, (ch) => ch.charCodeAt(0)))
    case 'array':
      return `[${token.items.map(formatToken).join(' ')}]`
    case 'dict':
      return `<<${Array.from(token.entries, ([key, value]) => `${formatName(key)} ${formatToken(value)}`).join(' ')}>>`
    case 'keyword':
      return token.value
  }
}
