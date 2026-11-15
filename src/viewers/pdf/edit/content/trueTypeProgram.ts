/**
 * Reads just enough of an embedded TrueType font program to answer "does this
 * font actually contain a drawable glyph for this code?" (PDF-EDIT-2).
 *
 * Writing new text in the document's own font is only safe for glyphs the
 * embedded subset really has. A font's ToUnicode map is not proof — some
 * producers write identity ranges covering every code — so when the program
 * is TrueType, the answer comes from its own `cmap` and `glyf` tables.
 */

export type TrueTypeProgram = {
  readonly numGlyphs: number
  /** Whether glyph `gid` has an outline (an empty glyph, such as a space, has none). */
  hasOutline(gid: number): boolean
  /** The glyph a simple font's code selects, per PDF 32000-1 §9.6.6.4; null when none. */
  glyphForSimpleCode(code: number, unicode: string | null): number | null
}

type Tables = Map<string, { offset: number; length: number }>

function tablesOf(view: DataView): Tables | null {
  if (view.byteLength < 12) return null
  const tables: Tables = new Map()
  const count = view.getUint16(4)
  for (let i = 0; i < count; i += 1) {
    const record = 12 + i * 16
    if (record + 16 > view.byteLength) return null
    const tag = String.fromCharCode(view.getUint8(record), view.getUint8(record + 1), view.getUint8(record + 2), view.getUint8(record + 3))
    tables.set(tag, { offset: view.getUint32(record + 8), length: view.getUint32(record + 12) })
  }
  return tables
}

/** A cmap subtable as code -> glyph lookups (formats 0, 4 and 6, which cover what PDF producers embed). */
function readSubtable(view: DataView, at: number): ((code: number) => number) | null {
  const format = view.getUint16(at)
  if (format === 0) {
    return (code) => (code >= 0 && code < 256 ? view.getUint8(at + 6 + code) : 0)
  }
  if (format === 6) {
    const first = view.getUint16(at + 6)
    const count = view.getUint16(at + 8)
    return (code) => (code >= first && code < first + count ? view.getUint16(at + 10 + (code - first) * 2) : 0)
  }
  if (format === 4) {
    const segCount = view.getUint16(at + 6) / 2
    const ends = at + 14
    const starts = ends + segCount * 2 + 2
    const deltas = starts + segCount * 2
    const rangeOffsets = deltas + segCount * 2
    return (code) => {
      for (let i = 0; i < segCount; i += 1) {
        if (code > view.getUint16(ends + i * 2)) continue
        const start = view.getUint16(starts + i * 2)
        if (code < start) return 0
        const delta = view.getInt16(deltas + i * 2)
        const rangeOffset = view.getUint16(rangeOffsets + i * 2)
        if (rangeOffset === 0) return (code + delta) & 0xffff
        const glyphAt = rangeOffsets + i * 2 + rangeOffset + (code - start) * 2
        if (glyphAt + 2 > view.byteLength) return 0
        const glyph = view.getUint16(glyphAt)
        return glyph === 0 ? 0 : (glyph + delta) & 0xffff
      }
      return 0
    }
  }
  return null
}

export function readTrueTypeProgram(bytes: Uint8Array): TrueTypeProgram | null {
  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    const tables = tablesOf(view)
    const head = tables?.get('head')
    const maxp = tables?.get('maxp')
    const loca = tables?.get('loca')
    const glyf = tables?.get('glyf')
    if (!tables || !head || !maxp || !loca || !glyf) return null
    const numGlyphs = view.getUint16(maxp.offset + 4)
    const longLoca = view.getInt16(head.offset + 50) === 1
    const range = (gid: number): number => {
      if (gid < 0 || gid >= numGlyphs) return 0
      return longLoca
        ? view.getUint32(loca.offset + (gid + 1) * 4) - view.getUint32(loca.offset + gid * 4)
        : (view.getUint16(loca.offset + (gid + 1) * 2) - view.getUint16(loca.offset + gid * 2)) * 2
    }

    const subtables = new Map<string, (code: number) => number>()
    const cmap = tables.get('cmap')
    if (cmap) {
      const count = view.getUint16(cmap.offset + 2)
      for (let i = 0; i < count; i += 1) {
        const record = cmap.offset + 4 + i * 8
        const key = `${view.getUint16(record)},${view.getUint16(record + 2)}`
        const lookup = readSubtable(view, cmap.offset + view.getUint32(record + 4))
        if (lookup && !subtables.has(key)) subtables.set(key, lookup)
      }
    }

    return {
      numGlyphs,
      hasOutline: (gid) => range(gid) > 0,
      glyphForSimpleCode(code, unicode) {
        const symbolic = subtables.get('3,0')
        if (symbolic) {
          const glyph = symbolic(0xf000 + code) || symbolic(code)
          if (glyph) return glyph
        }
        const mac = subtables.get('1,0')
        if (mac) {
          const glyph = mac(code)
          if (glyph) return glyph
        }
        const windows = subtables.get('3,1')
        if (windows && unicode && unicode.length === 1) {
          const glyph = windows(unicode.charCodeAt(0))
          if (glyph) return glyph
        }
        // No cmap at all: subset producers often map codes straight to glyph IDs.
        return subtables.size === 0 && code < numGlyphs ? code : null
      },
    }
  } catch {
    return null
  }
}
