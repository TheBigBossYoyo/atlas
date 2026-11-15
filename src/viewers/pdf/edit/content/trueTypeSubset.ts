/**
 * A minimal TrueType subsetter for embedding the bundled fallback fonts
 * (PDF-EDIT-2).
 *
 * pdf-lib's own subsetting (through fontkit) renumbers glyphs and rebuilds the
 * font, and for two of the four bundled families it silently loses most of
 * them — "Added with Atlas" rendered as "as". This takes the conservative
 * route instead: every glyph keeps its ID, glyphs that aren't needed simply
 * get an empty outline, and only the tables a PDF reader needs to draw
 * outlines are kept. With IDs unchanged, `hmtx`, `loca` indexing and the
 * CIDs written in the content stream all stay trivially consistent.
 *
 * Variation tables are dropped, so a variable font comes out as its default
 * instance — the regular weight for the bundled Arimo.
 */

const KEEP_TABLES = ['OS/2', 'cvt ', 'fpgm', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'post', 'prep']

const ARG_1_AND_2_ARE_WORDS = 0x0001
const WE_HAVE_A_SCALE = 0x0008
const MORE_COMPONENTS = 0x0020
const WE_HAVE_AN_X_AND_Y_SCALE = 0x0040
const WE_HAVE_A_TWO_BY_TWO = 0x0080

type Table = { readonly tag: string; readonly offset: number; readonly length: number }

function readTables(view: DataView): Map<string, Table> {
  const count = view.getUint16(4)
  const tables = new Map<string, Table>()
  for (let i = 0; i < count; i += 1) {
    const record = 12 + i * 16
    const tag = String.fromCharCode(
      view.getUint8(record),
      view.getUint8(record + 1),
      view.getUint8(record + 2),
      view.getUint8(record + 3),
    )
    tables.set(tag, { tag, offset: view.getUint32(record + 8), length: view.getUint32(record + 12) })
  }
  return tables
}

function checksum(bytes: Uint8Array): number {
  const padded = new Uint8Array(Math.ceil(bytes.length / 4) * 4)
  padded.set(bytes)
  const view = new DataView(padded.buffer)
  let sum = 0
  for (let i = 0; i < padded.length; i += 4) sum = (sum + view.getUint32(i)) >>> 0
  return sum
}

/**
 * A copy of `font` that keeps only `glyphIds` (plus `.notdef` and the
 * components of any composite glyph among them). Throws on a font that isn't
 * TrueType-outlined.
 */
export function subsetTrueType(font: Uint8Array, glyphIds: ReadonlySet<number>): Uint8Array {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength)
  const tables = readTables(view)
  const head = tables.get('head')
  const maxp = tables.get('maxp')
  const loca = tables.get('loca')
  const glyf = tables.get('glyf')
  if (!head || !maxp || !loca || !glyf) throw new Error('Not a TrueType-outlined font.')

  const numGlyphs = view.getUint16(maxp.offset + 4)
  const longLoca = view.getInt16(head.offset + 50) === 1
  const glyphRange = (id: number): [number, number] => {
    if (longLoca) return [view.getUint32(loca.offset + id * 4), view.getUint32(loca.offset + (id + 1) * 4)]
    return [view.getUint16(loca.offset + id * 2) * 2, view.getUint16(loca.offset + (id + 1) * 2) * 2]
  }

  // Close over composite glyphs' components.
  const keep = new Set<number>([0])
  const pending = [...glyphIds]
  while (pending.length > 0) {
    const id = pending.pop() as number
    if (id < 0 || id >= numGlyphs || keep.has(id)) continue
    keep.add(id)
    const [start, end] = glyphRange(id)
    if (end - start < 10) continue
    const at = glyf.offset + start
    if (view.getInt16(at) >= 0) continue
    let cursor = at + 10
    for (;;) {
      const flags = view.getUint16(cursor)
      pending.push(view.getUint16(cursor + 2))
      cursor += 4 + (flags & ARG_1_AND_2_ARE_WORDS ? 4 : 2)
      if (flags & WE_HAVE_A_SCALE) cursor += 2
      else if (flags & WE_HAVE_AN_X_AND_Y_SCALE) cursor += 4
      else if (flags & WE_HAVE_A_TWO_BY_TWO) cursor += 8
      if (!(flags & MORE_COMPONENTS)) break
    }
  }

  // New glyf and long-format loca, glyph IDs unchanged.
  const pieces: Uint8Array[] = []
  const newLoca = new DataView(new ArrayBuffer((numGlyphs + 1) * 4))
  let glyfLength = 0
  for (let id = 0; id < numGlyphs; id += 1) {
    newLoca.setUint32(id * 4, glyfLength)
    if (!keep.has(id)) continue
    const [start, end] = glyphRange(id)
    if (end <= start) continue
    const length = end - start
    const padded = new Uint8Array(Math.ceil(length / 4) * 4)
    padded.set(font.subarray(glyf.offset + start, glyf.offset + end))
    pieces.push(padded)
    glyfLength += padded.length
  }
  newLoca.setUint32(numGlyphs * 4, glyfLength)
  const newGlyf = new Uint8Array(glyfLength)
  let offset = 0
  for (const piece of pieces) {
    newGlyf.set(piece, offset)
    offset += piece.length
  }

  const tableData = new Map<string, Uint8Array>()
  for (const tag of KEEP_TABLES) {
    const table = tables.get(tag)
    if (!table) continue
    tableData.set(tag, font.slice(table.offset, table.offset + table.length))
  }
  tableData.set('glyf', newGlyf)
  tableData.set('loca', new Uint8Array(newLoca.buffer))
  const newHead = tableData.get('head') as Uint8Array
  const headView = new DataView(newHead.buffer, newHead.byteOffset, newHead.byteLength)
  headView.setInt16(50, 1) // long loca
  headView.setUint32(8, 0) // checkSumAdjustment, recomputed below
  const post = tableData.get('post')
  if (post && post.length >= 32) {
    // Version 3: no glyph names, so the table no longer lists glyphs that aren't there.
    const trimmed = post.slice(0, 32)
    new DataView(trimmed.buffer).setUint32(0, 0x00030000)
    tableData.set('post', trimmed)
  }

  // Assemble: table directory, then 4-byte-aligned tables in tag order.
  const tags = [...tableData.keys()].sort()
  const headerLength = 12 + tags.length * 16
  let total = headerLength
  for (const tag of tags) total += Math.ceil((tableData.get(tag) as Uint8Array).length / 4) * 4
  const out = new Uint8Array(total)
  const outView = new DataView(out.buffer)
  outView.setUint32(0, 0x00010000)
  outView.setUint16(4, tags.length)
  const entrySelector = Math.floor(Math.log2(tags.length))
  const searchRange = 2 ** entrySelector * 16
  outView.setUint16(6, searchRange)
  outView.setUint16(8, entrySelector)
  outView.setUint16(10, tags.length * 16 - searchRange)
  let cursor = headerLength
  let headOffset = 0
  tags.forEach((tag, index) => {
    const data = tableData.get(tag) as Uint8Array
    const record = 12 + index * 16
    for (let i = 0; i < 4; i += 1) outView.setUint8(record + i, tag.charCodeAt(i))
    outView.setUint32(record + 4, checksum(data))
    outView.setUint32(record + 8, cursor)
    outView.setUint32(record + 12, data.length)
    out.set(data, cursor)
    if (tag === 'head') headOffset = cursor
    cursor += Math.ceil(data.length / 4) * 4
  })
  outView.setUint32(headOffset + 8, (0xb1b0afba - checksum(out)) >>> 0)
  return out
}
