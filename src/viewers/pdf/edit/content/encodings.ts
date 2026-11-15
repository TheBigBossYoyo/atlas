/**
 * Simple-font encodings: byte code -> glyph name -> Unicode (PDF-EDIT-2).
 *
 * Only needed for fonts without a `/ToUnicode` CMap, which most modern
 * producers write. WinAnsi comes from pdf-lib's own tables; MacRoman is
 * spelled out here (old Mac producers still use it); StandardEncoding is
 * WinAnsi for printable ASCII apart from its two curly quotes.
 */

import { Encodings } from '@pdf-lib/standard-fonts'

export type SimpleEncoding = {
  /** code -> glyph name */
  readonly names: ReadonlyArray<string | undefined>
}

const winAnsiNames: Array<string | undefined> = new Array(256)
/** Glyph name -> Unicode text, for every name WinAnsi knows plus common extras. */
const nameToUnicode = new Map<string, string>()

for (const codePoint of Encodings.WinAnsi.supportedCodePoints) {
  const { code, name } = Encodings.WinAnsi.encodeUnicodeCodePoint(codePoint)
  if (winAnsiNames[code] === undefined) winAnsiNames[code] = name
  if (!nameToUnicode.has(name)) nameToUnicode.set(name, String.fromCodePoint(codePoint))
}

const EXTRA_NAMES: Record<string, string> = {
  fi: 'ﬁ',
  fl: 'ﬂ',
  ff: 'ﬀ',
  ffi: 'ﬃ',
  ffl: 'ﬄ',
  dotlessi: 'ı',
  nbspace: String.fromCharCode(0xa0),
  nonbreakingspace: String.fromCharCode(0xa0),
  hyphen: '-',
  sfthyphen: String.fromCharCode(0xad),
  minus: '−',
  fraction: '⁄',
  apple: String.fromCharCode(0xf8ff),
  Omega: 'Ω',
  Delta: '∆',
  pi: 'π',
  mu: 'µ',
  lozenge: '◊',
  infinity: '∞',
  notequal: '≠',
  lessequal: '≤',
  greaterequal: '≥',
  partialdiff: '∂',
  summation: '∑',
  product: '∏',
  integral: '∫',
  radical: '√',
  approxequal: '≈',
  ring: '˚',
  breve: '˘',
  dotaccent: '˙',
  hungarumlaut: '˝',
  ogonek: '˛',
  caron: 'ˇ',
  Euro: '€',
  bullet: '•',
}
for (const [name, text] of Object.entries(EXTRA_NAMES)) if (!nameToUnicode.has(name)) nameToUnicode.set(name, text)

/** Unicode text for a glyph name: the tables above, then the `uniXXXX`/`uXXXX[XX]` conventions. */
export function glyphNameToUnicode(name: string): string | null {
  const known = nameToUnicode.get(name)
  if (known) return known
  const uni = /^uni([0-9A-Fa-f]{4})+$/.exec(name)
  if (uni) {
    const hex = name.slice(3)
    let out = ''
    for (let i = 0; i < hex.length; i += 4) out += String.fromCharCode(parseInt(hex.slice(i, i + 4), 16))
    return out
  }
  const u = /^u([0-9A-Fa-f]{4,6})$/.exec(name)
  if (u) return String.fromCodePoint(parseInt(u[1], 16))
  // Suffixed variants ("a.sc", "one.oldstyle") read as their base glyph.
  const dot = name.indexOf('.')
  if (dot > 0) return glyphNameToUnicode(name.slice(0, dot))
  return null
}

/** Mac OS Roman, 0x80-0xFF, as Unicode code points. */
const MAC_ROMAN_HIGH = [
  0xc4, 0xc5, 0xc7, 0xc9, 0xd1, 0xd6, 0xdc, 0xe1, 0xe0, 0xe2, 0xe4, 0xe3, 0xe5, 0xe7, 0xe9, 0xe8,
  0xea, 0xeb, 0xed, 0xec, 0xee, 0xef, 0xf1, 0xf3, 0xf2, 0xf4, 0xf6, 0xf5, 0xfa, 0xf9, 0xfb, 0xfc,
  0x2020, 0xb0, 0xa2, 0xa3, 0xa7, 0x2022, 0xb6, 0xdf, 0xae, 0xa9, 0x2122, 0xb4, 0xa8, 0x2260, 0xc6, 0xd8,
  0x221e, 0xb1, 0x2264, 0x2265, 0xa5, 0xb5, 0x2202, 0x2211, 0x220f, 0x3c0, 0x222b, 0xaa, 0xba, 0x3a9, 0xe6, 0xf8,
  0xbf, 0xa1, 0xac, 0x221a, 0x192, 0x2248, 0x2206, 0xab, 0xbb, 0x2026, 0xa0, 0xc0, 0xc3, 0xd5, 0x152, 0x153,
  0x2013, 0x2014, 0x201c, 0x201d, 0x2018, 0x2019, 0xf7, 0x25ca, 0xff, 0x178, 0x2044, 0x20ac, 0x2039, 0x203a, 0xfb01, 0xfb02,
  0x2021, 0xb7, 0x201a, 0x201e, 0x2030, 0xc2, 0xca, 0xc1, 0xcb, 0xc8, 0xcd, 0xce, 0xcf, 0xcc, 0xd3, 0xd4,
  0xf8ff, 0xd2, 0xda, 0xdb, 0xd9, 0x131, 0x2c6, 0x2dc, 0xaf, 0x2d8, 0x2d9, 0x2da, 0xb8, 0x2dd, 0x2db, 0x2c7,
]

const unicodeToName = new Map<string, string>()
nameToUnicode.forEach((text, name) => {
  if (!unicodeToName.has(text)) unicodeToName.set(text, name)
})

function nameForUnicode(codePoint: number): string | undefined {
  return unicodeToName.get(String.fromCodePoint(codePoint)) ?? `uni${codePoint.toString(16).toUpperCase().padStart(4, '0')}`
}

const macRomanNames: Array<string | undefined> = winAnsiNames.map((name, code) => (code < 0x80 ? name : undefined))
MAC_ROMAN_HIGH.forEach((codePoint, index) => {
  macRomanNames[0x80 + index] = nameForUnicode(codePoint)
})

const standardNames: Array<string | undefined> = winAnsiNames.map((name, code) => (code < 0x80 ? name : undefined))
standardNames[0x27] = 'quoteright'
standardNames[0x60] = 'quoteleft'

function builtInNames(encoding: typeof Encodings.Symbol): Array<string | undefined> {
  const names: Array<string | undefined> = new Array(256)
  for (const codePoint of encoding.supportedCodePoints) {
    const { code, name } = encoding.encodeUnicodeCodePoint(codePoint)
    if (names[code] === undefined) names[code] = name
  }
  return names
}

const symbolNames = builtInNames(Encodings.Symbol)
const zapfNames = builtInNames(Encodings.ZapfDingbats)

export function baseEncoding(name: string | undefined): Array<string | undefined> {
  switch (name) {
    case 'Symbol':
      return [...symbolNames]
    case 'ZapfDingbats':
      return [...zapfNames]
    case 'MacRomanEncoding':
      return [...macRomanNames]
    case 'StandardEncoding':
      return [...standardNames]
    default:
      return [...winAnsiNames]
  }
}
