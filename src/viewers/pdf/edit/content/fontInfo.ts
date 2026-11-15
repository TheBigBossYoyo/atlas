/**
 * What text editing needs to know about a PDF font (PDF-EDIT-2): how shown
 * bytes split into character codes, each code's advance width and Unicode
 * text, and — the other way — whether some new text can be written in this
 * same font and with which codes.
 *
 * That last part is what lets an edit keep the document's own typeface. A
 * subset font only carries the glyphs the original text used, so "can this
 * font draw it?" is answered conservatively: a code counts as available only
 * when the font maps it to text (its ToUnicode CMap, or its encoding for an
 * unsubsetted simple font) AND gives it a real width.
 */

import { Font as StandardFont, type IFontNames } from '@pdf-lib/standard-fonts'
import type { PDFContext, PDFDict } from 'pdf-lib'

import type { PdfLib } from '../writePdf'
import { guessFontStyle } from '../textLayout'
import type { FontStyle } from '../types'
import { baseEncoding, glyphNameToUnicode } from './encodings'
import { lookupCid, parseCMap, splitCodes, type CodeSpaceRange, type ParsedCMap } from './cmap'
import { asDict, dictGet, nameOf, numberOf, numbersOf, resolve, streamText } from './pdfObjects'
import { binaryStringToBytes } from './lexer'
import { readTrueTypeProgram, type TrueTypeProgram } from './trueTypeProgram'

export type CharCode = { readonly code: number; readonly length: number }

export type FontInfo = {
  /** Stable identity (the font dictionary's object reference, when it has one). */
  readonly key: string
  readonly baseFont: string
  readonly composite: boolean
  /** False for fonts whose metrics or codes cannot be worked out (Type 3, vertical, unknown CMaps). */
  readonly editable: boolean
  readonly style: FontStyle
  split(bytes: string): CharCode[]
  /** Advance width in glyph units (1/1000 of the font size). */
  width(code: number): number
  /** The code's text, or U+FFFD when the font gives no way to know it. */
  unicode(code: number): string
  /** Codes (as bytes) for `text` in this font, or null if any character is unavailable. */
  encode(text: string): CharCode[] | null
  /** Byte values for a code, for writing it back into a string. */
  bytesOf(code: CharCode): number[]
}

const STANDARD_14: Record<string, IFontNames> = {
  Courier: 'Courier',
  'Courier-Bold': 'Courier-Bold',
  'Courier-Oblique': 'Courier-Oblique',
  'Courier-BoldOblique': 'Courier-BoldOblique',
  Helvetica: 'Helvetica',
  'Helvetica-Bold': 'Helvetica-Bold',
  'Helvetica-Oblique': 'Helvetica-Oblique',
  'Helvetica-BoldOblique': 'Helvetica-BoldOblique',
  'Times-Roman': 'Times-Roman',
  'Times-Bold': 'Times-Bold',
  'Times-Italic': 'Times-Italic',
  'Times-BoldItalic': 'Times-BoldItalic',
  Symbol: 'Symbol',
  ZapfDingbats: 'ZapfDingbats',
  // Common aliases readers map onto the standard fonts.
  Arial: 'Helvetica',
  'Arial,Bold': 'Helvetica-Bold',
  'Arial,Italic': 'Helvetica-Oblique',
  'Arial,BoldItalic': 'Helvetica-BoldOblique',
  ArialMT: 'Helvetica',
  'Arial-BoldMT': 'Helvetica-Bold',
  TimesNewRoman: 'Times-Roman',
  TimesNewRomanPSMT: 'Times-Roman',
  'TimesNewRoman,Bold': 'Times-Bold',
  CourierNew: 'Courier',
  CourierNewPSMT: 'Courier',
}

const standardFontCache = new Map<IFontNames, StandardFont>()
function standardFont(name: string): StandardFont | null {
  const mapped = STANDARD_14[name]
  if (!mapped) return null
  let font = standardFontCache.get(mapped)
  if (!font) {
    font = StandardFont.load(mapped)
    standardFontCache.set(mapped, font)
  }
  return font
}

const FLAG_FIXED_PITCH = 1
const FLAG_SERIF = 2
const FLAG_ITALIC = 64
const FLAG_FORCE_BOLD = 1 << 18

function styleOf(lib: PdfLib, context: PDFContext, baseFont: string, descriptor: PDFDict | undefined): FontStyle {
  const guessed = guessFontStyle(undefined, baseFont.replace(/^[A-Z]{6}\+/, ''))
  const flags = numberOf(lib, dictGet(lib, context, descriptor, 'Flags')) ?? 0
  const weight = numberOf(lib, dictGet(lib, context, descriptor, 'FontWeight'))
  const italicAngle = numberOf(lib, dictGet(lib, context, descriptor, 'ItalicAngle')) ?? 0
  let family = guessed.family
  if (flags & FLAG_FIXED_PITCH) family = 'courier'
  else if (family === 'helvetica' && flags & FLAG_SERIF && !/sans/i.test(baseFont)) family = 'times'
  return {
    family,
    bold: guessed.bold || (flags & FLAG_FORCE_BOLD) !== 0 || (weight !== undefined && weight >= 600),
    italic: guessed.italic || (flags & FLAG_ITALIC) !== 0 || italicAngle !== 0,
  }
}

/** The embedded font program's kind, and the TrueType program itself when that's what it is. */
function embeddedProgram(
  lib: PdfLib,
  context: PDFContext,
  descriptor: PDFDict | undefined,
): { readonly embedded: boolean; readonly trueType: TrueTypeProgram | null } {
  if (!descriptor) return { embedded: false, trueType: null }
  const fontFile2 = resolve(lib, context, descriptor.get(lib.PDFName.of('FontFile2')))
  if (fontFile2 instanceof lib.PDFStream) {
    const text = streamText(lib, fontFile2)
    return { embedded: true, trueType: text === null ? null : readTrueTypeProgram(binaryStringToBytes(text)) }
  }
  const embedded = ['FontFile', 'FontFile3'].some((key) => descriptor.get(lib.PDFName.of(key)) !== undefined)
  return { embedded, trueType: null }
}

function toUnicodeOf(lib: PdfLib, context: PDFContext, fontDict: PDFDict): ParsedCMap | null {
  const stream = resolve(lib, context, fontDict.get(lib.PDFName.of('ToUnicode')))
  if (!(stream instanceof lib.PDFStream)) return null
  const text = streamText(lib, stream)
  return text === null ? null : parseCMap(text)
}

/** Greedy longest-match encoder over a text -> code table (ligature entries such as "ffi" win). */
function makeEncoder(table: ReadonlyMap<string, CharCode>): (text: string) => CharCode[] | null {
  const longest = Math.max(1, ...Array.from(table.keys(), (key) => key.length))
  return (text) => {
    const out: CharCode[] = []
    let pos = 0
    while (pos < text.length) {
      let found: CharCode | undefined
      let used = 0
      for (let length = Math.min(longest, text.length - pos); length >= 1 && !found; length -= 1) {
        found = table.get(text.slice(pos, pos + length))
        used = length
      }
      if (!found) return null
      out.push(found)
      pos += used
    }
    return out
  }
}

function codeBytes(code: CharCode): number[] {
  const bytes: number[] = []
  for (let shift = (code.length - 1) * 8; shift >= 0; shift -= 8) bytes.push((code.code >> shift) & 0xff)
  return bytes
}

function simpleFont(lib: PdfLib, context: PDFContext, fontDict: PDFDict, key: string, subtype: string): FontInfo {
  const baseFont = nameOf(lib, dictGet(lib, context, fontDict, 'BaseFont')) ?? ''
  const plainName = baseFont.replace(/^[A-Z]{6}\+/, '')
  const isSubset = plainName !== baseFont
  const descriptor = asDict(lib, context, fontDict.get(lib.PDFName.of('FontDescriptor')))
  const firstChar = numberOf(lib, dictGet(lib, context, fontDict, 'FirstChar')) ?? 0
  const widths = numbersOf(lib, context, fontDict.get(lib.PDFName.of('Widths')))
  const missingWidth = numberOf(lib, dictGet(lib, context, descriptor, 'MissingWidth')) ?? 0
  const standard = widths ? null : standardFont(plainName)

  // Encoding: a base table plus /Differences.
  const encodingValue = dictGet(lib, context, fontDict, 'Encoding')
  const encodingDict = encodingValue instanceof lib.PDFDict ? encodingValue : undefined
  const explicitBase = nameOf(lib, encodingValue) ?? nameOf(lib, dictGet(lib, context, encodingDict, 'BaseEncoding'))
  // With no encoding named, a standard font uses its own built-in one.
  const builtIn = STANDARD_14[plainName] === 'Symbol' ? 'Symbol' : STANDARD_14[plainName] === 'ZapfDingbats' ? 'ZapfDingbats' : STANDARD_14[plainName] ? 'StandardEncoding' : undefined
  const names = baseEncoding(explicitBase ?? builtIn)
  const differences = dictGet(lib, context, encodingDict, 'Differences')
  if (differences instanceof lib.PDFArray) {
    let code = 0
    for (const item of differences.asArray()) {
      const resolved = resolve(lib, context, item)
      if (resolved instanceof lib.PDFNumber) code = resolved.asNumber()
      else if (resolved instanceof lib.PDFName) {
        names[code] = resolved.decodeText()
        code += 1
      }
    }
  }

  const toUnicode = toUnicodeOf(lib, context, fontDict)
  const width = (code: number): number => {
    if (widths) {
      const w = widths[code - firstChar]
      return w === undefined ? missingWidth : w
    }
    if (standard) {
      const name = names[code]
      return (name ? standard.getWidthOfGlyph(name) : undefined) ?? 0
    }
    return missingWidth || 500
  }
  const unicode = (code: number): string => {
    const mapped = toUnicode?.toUnicode.get(code)
    if (mapped !== undefined) return mapped
    const name = names[code]
    return (name ? glyphNameToUnicode(name) : null) ?? (code >= 0x20 && code < 0x7f && !isSubset ? String.fromCharCode(code) : '�')
  }

  const program = embeddedProgram(lib, context, descriptor)
  const table = new Map<string, CharCode>()
  for (let code = 0; code < 256; code += 1) {
    if (width(code) <= 0 && code !== 32) continue
    let text: string | undefined = toUnicode?.toUnicode.get(code)
    // Without a ToUnicode map, a subset font's encoding names glyphs it may not
    // contain; only an unsubsetted font's encoding can be trusted for new text.
    if (text === undefined && !isSubset && names[code]) text = glyphNameToUnicode(names[code] as string) ?? undefined
    if (!text || table.has(text)) continue
    if (program.trueType && !/^\s+$/.test(text)) {
      const glyph = program.trueType.glyphForSimpleCode(code, text)
      if (glyph === null || !program.trueType.hasOutline(glyph)) continue
    }
    table.set(text, { code, length: 1 })
  }

  return {
    key,
    baseFont,
    composite: false,
    editable: subtype !== 'Type3',
    style: styleOf(lib, context, baseFont, descriptor),
    split: (bytes) => Array.from(bytes, (ch) => ({ code: ch.charCodeAt(0), length: 1 })),
    width,
    unicode,
    encode: makeEncoder(table),
    bytesOf: codeBytes,
  }
}

function compositeFont(lib: PdfLib, context: PDFContext, fontDict: PDFDict, key: string): FontInfo {
  const baseFont = nameOf(lib, dictGet(lib, context, fontDict, 'BaseFont')) ?? ''
  const descendants = dictGet(lib, context, fontDict, 'DescendantFonts')
  const cidFont = descendants instanceof lib.PDFArray ? asDict(lib, context, descendants.get(0)) : undefined
  const descriptor = asDict(lib, context, cidFont?.get(lib.PDFName.of('FontDescriptor')))
  const defaultWidth = numberOf(lib, dictGet(lib, context, cidFont, 'DW')) ?? 1000

  // /W: [c [w1 w2 ...]] or [cFirst cLast w] groups.
  const cidWidths = new Map<number, number>()
  const widthRanges: Array<{ first: number; last: number; width: number }> = []
  const w = dictGet(lib, context, cidFont, 'W')
  if (w instanceof lib.PDFArray) {
    const items = w.asArray().map((item) => resolve(lib, context, item))
    for (let i = 0; i < items.length; ) {
      const first = numberOf(lib, items[i])
      const next = items[i + 1]
      if (first === undefined) break
      if (next instanceof lib.PDFArray) {
        next.asArray().forEach((value, offset) => {
          const width = numberOf(lib, resolve(lib, context, value))
          if (width !== undefined) cidWidths.set(first + offset, width)
        })
        i += 2
      } else {
        const last = numberOf(lib, next)
        const width = numberOf(lib, items[i + 2])
        if (last === undefined || width === undefined) break
        widthRanges.push({ first, last, width })
        i += 3
      }
    }
  }

  // Encoding: Identity-H/V, or an embedded CMap stream.
  const encoding = resolve(lib, context, fontDict.get(lib.PDFName.of('Encoding')))
  const encodingName = nameOf(lib, encoding)
  let codeSpace: CodeSpaceRange[] = [{ length: 2, low: 0, high: 0xffff }]
  let cmap: ParsedCMap | null = null
  let vertical = false
  let knownEncoding = true
  if (encodingName) {
    vertical = encodingName.endsWith('-V')
    knownEncoding = encodingName === 'Identity-H' || encodingName === 'Identity-V'
  } else if (encoding instanceof lib.PDFStream) {
    const text = streamText(lib, encoding)
    cmap = text === null ? null : parseCMap(text)
    if (!cmap || cmap.usesExternal) knownEncoding = false
    else if (cmap.codeSpace.length > 0) codeSpace = [...cmap.codeSpace]
    vertical = numberOf(lib, dictGet(lib, context, encoding.dict, 'WMode')) === 1
  } else {
    knownEncoding = false
  }

  const cidOf = (code: number): number => (cmap ? (lookupCid(cmap, code) ?? 0) : code)
  const width = (code: number): number => {
    const cid = cidOf(code)
    const direct = cidWidths.get(cid)
    if (direct !== undefined) return direct
    const range = widthRanges.find((r) => cid >= r.first && cid <= r.last)
    return range ? range.width : defaultWidth
  }

  // Which CIDs the embedded program can really draw (see the file comment).
  const program = embeddedProgram(lib, context, descriptor)
  const cidToGid = resolve(lib, context, cidFont?.get(lib.PDFName.of('CIDToGIDMap')))
  const cidToGidBytes = cidToGid instanceof lib.PDFStream ? streamText(lib, cidToGid) : null
  const gidOf = (cid: number): number =>
    cidToGidBytes === null ? cid : cid * 2 + 1 < cidToGidBytes.length ? (cidToGidBytes.charCodeAt(cid * 2) << 8) | cidToGidBytes.charCodeAt(cid * 2 + 1) : 0
  const drawable = (code: number, text: string): boolean => {
    if (/^\s+$/.test(text)) return true
    const cid = cidOf(code)
    if (program.trueType) return program.trueType.hasOutline(gidOf(cid))
    // An embedded non-TrueType subset: its /W lists the glyphs it carries.
    if (program.embedded) return cidWidths.has(cid) || widthRanges.some((r) => cid >= r.first && cid <= r.last)
    return true
  }

  const toUnicode = toUnicodeOf(lib, context, fontDict)
  const table = new Map<string, CharCode>()
  toUnicode?.toUnicode.forEach((text, code) => {
    if (text.length === 0 || table.has(text) || !drawable(code, text)) return
    const length = codeSpace.find((range) => code >= range.low && code <= range.high)?.length ?? 2
    table.set(text, { code, length })
  })

  return {
    key,
    baseFont,
    composite: true,
    editable: knownEncoding && !vertical,
    style: styleOf(lib, context, baseFont, descriptor),
    split: (bytes) => splitCodes(bytes, codeSpace),
    width,
    unicode: (code) => toUnicode?.toUnicode.get(code) ?? '�',
    encode: makeEncoder(table),
    bytesOf: codeBytes,
  }
}

const fontInfoCache = new WeakMap<PDFDict, FontInfo>()

/** Font facts for a font dictionary, cached per dictionary object. */
export function readFontInfo(lib: PdfLib, context: PDFContext, fontDict: PDFDict, key: string): FontInfo {
  const cached = fontInfoCache.get(fontDict)
  if (cached) return cached
  const subtype = nameOf(lib, dictGet(lib, context, fontDict, 'Subtype')) ?? ''
  const info = subtype === 'Type0' ? compositeFont(lib, context, fontDict, key) : simpleFont(lib, context, fontDict, key, subtype)
  fontInfoCache.set(fontDict, info)
  return info
}
