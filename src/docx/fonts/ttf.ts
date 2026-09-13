export class FontParseError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'FontParseError'
  }
}

export type TtfHeadTable = Readonly<{
  unitsPerEm: number
}>

export type TtfHheaTable = Readonly<{
  ascender: number
  descender: number
  lineGap: number
  numberOfHMetrics: number
}>

export type TtfHmtxTable = Readonly<{
  advanceWidths: readonly number[]
  leftSideBearings: readonly number[]
}>

export type TtfOs2Table = Readonly<{
  sTypoAscender: number
  sTypoDescender: number
  sxHeight: number | null
  sCapHeight: number | null
}>

export type TtfNameTable = Readonly<{
  familyName: string
}>

export type TtfTables = Readonly<{
  head: TtfHeadTable
  hhea: TtfHheaTable
  hmtx: TtfHmtxTable
  cmap: ReadonlyMap<number, number>
  os2: TtfOs2Table
  name: TtfNameTable
}>

type TableRecord = Readonly<{
  offset: number
  length: number
}>

type NameCandidate = Readonly<{
  familyName: string
  score: number
}>

const TRUE_TYPE_SIGNATURE = 0x00010000
const OPEN_TYPE_SIGNATURE = 0x4f54544f
const HEAD_UNITS_PER_EM_OFFSET = 18
const HHEA_ASCENDER_OFFSET = 4
const HHEA_DESCENDER_OFFSET = 6
const HHEA_LINE_GAP_OFFSET = 8
const HHEA_NUMBER_OF_HMETRICS_OFFSET = 34
const MAXP_NUM_GLYPHS_OFFSET = 4
const OS2_VERSION_OFFSET = 0
const OS2_TYPO_ASCENDER_OFFSET = 68
const OS2_TYPO_DESCENDER_OFFSET = 70
const OS2_X_HEIGHT_OFFSET = 86
const OS2_CAP_HEIGHT_OFFSET = 88

export function parseTtf(buffer: ArrayBuffer): TtfTables {
  const view = new DataView(buffer)
  const signature = readUint32(view, 0)
  if (signature !== TRUE_TYPE_SIGNATURE && signature !== OPEN_TYPE_SIGNATURE) {
    throw new FontParseError('Unsupported font signature')
  }

  const tableCount = readUint16(view, 4)
  const records = readTableDirectory(view, tableCount)
  const headRecord = requireTable(records, 'head')
  const hheaRecord = requireTable(records, 'hhea')
  const hmtxRecord = requireTable(records, 'hmtx')
  const cmapRecord = requireTable(records, 'cmap')
  const maxpRecord = requireTable(records, 'maxp')
  const os2Record = requireTable(records, 'OS/2')
  const nameRecord = requireTable(records, 'name')

  const head = parseHead(view, headRecord)
  const hhea = parseHhea(view, hheaRecord)
  const numGlyphs = parseNumGlyphs(view, maxpRecord)
  const hmtx = parseHmtx(view, hmtxRecord, hhea.numberOfHMetrics, numGlyphs)
  const cmap = parseCmap(view, cmapRecord)
  const os2 = parseOs2(view, os2Record)
  const name = parseName(view, nameRecord)

  return {
    head,
    hhea,
    hmtx,
    cmap,
    os2,
    name,
  }
}

function readTableDirectory(view: DataView, tableCount: number): Map<string, TableRecord> {
  const records = new Map<string, TableRecord>()
  const directoryOffset = 12

  for (let index = 0; index < tableCount; index += 1) {
    const recordOffset = directoryOffset + index * 16
    assertRange(view, recordOffset, 16, 'table directory record')
    const tag = readTag(view, recordOffset)
    const offset = readUint32(view, recordOffset + 8)
    const length = readUint32(view, recordOffset + 12)
    assertRange(view, offset, length, `${tag} table`)
    records.set(tag, { offset, length })
  }

  return records
}

function requireTable(records: ReadonlyMap<string, TableRecord>, tag: string): TableRecord {
  const record = records.get(tag)
  if (!record) {
    throw new FontParseError(`Missing required ${tag} table`)
  }

  return record
}

function parseHead(view: DataView, record: TableRecord): TtfHeadTable {
  assertRange(view, record.offset, HEAD_UNITS_PER_EM_OFFSET + 2, 'head table')
  return {
    unitsPerEm: readUint16(view, record.offset + HEAD_UNITS_PER_EM_OFFSET),
  }
}

function parseHhea(view: DataView, record: TableRecord): TtfHheaTable {
  assertRange(view, record.offset, HHEA_NUMBER_OF_HMETRICS_OFFSET + 2, 'hhea table')
  return {
    ascender: readInt16(view, record.offset + HHEA_ASCENDER_OFFSET),
    descender: readInt16(view, record.offset + HHEA_DESCENDER_OFFSET),
    lineGap: readInt16(view, record.offset + HHEA_LINE_GAP_OFFSET),
    numberOfHMetrics: readUint16(view, record.offset + HHEA_NUMBER_OF_HMETRICS_OFFSET),
  }
}

function parseNumGlyphs(view: DataView, record: TableRecord): number {
  assertRange(view, record.offset, MAXP_NUM_GLYPHS_OFFSET + 2, 'maxp table')
  return readUint16(view, record.offset + MAXP_NUM_GLYPHS_OFFSET)
}

function parseHmtx(
  view: DataView,
  record: TableRecord,
  numberOfHMetrics: number,
  numGlyphs: number,
): TtfHmtxTable {
  if (numberOfHMetrics < 1 || numberOfHMetrics > numGlyphs) {
    throw new FontParseError('Invalid numberOfHMetrics value')
  }

  const advanceWidths: number[] = []
  const leftSideBearings: number[] = []
  let offset = record.offset

  for (let index = 0; index < numberOfHMetrics; index += 1) {
    assertRange(view, offset, 4, 'hmtx longHorMetric')
    advanceWidths.push(readUint16(view, offset))
    leftSideBearings.push(readInt16(view, offset + 2))
    offset += 4
  }

  const lastAdvanceWidth = advanceWidths[advanceWidths.length - 1]
  for (let index = numberOfHMetrics; index < numGlyphs; index += 1) {
    assertRange(view, offset, 2, 'hmtx leftSideBearing')
    advanceWidths.push(lastAdvanceWidth)
    leftSideBearings.push(readInt16(view, offset))
    offset += 2
  }

  return {
    advanceWidths,
    leftSideBearings,
  }
}

function parseCmap(view: DataView, record: TableRecord): ReadonlyMap<number, number> {
  assertRange(view, record.offset, 4, 'cmap table')
  const subtableCount = readUint16(view, record.offset + 2)
  const combined = new Map<number, number>()
  const subtables = new Array<{ format: number; offset: number }>()

  for (let index = 0; index < subtableCount; index += 1) {
    const encodingRecordOffset = record.offset + 4 + index * 8
    assertRange(view, encodingRecordOffset, 8, 'cmap encoding record')
    const platformId = readUint16(view, encodingRecordOffset)
    const encodingId = readUint16(view, encodingRecordOffset + 2)
    const subtableOffset = record.offset + readUint32(view, encodingRecordOffset + 4)
    assertRange(view, subtableOffset, 2, 'cmap subtable header')
    const format = readUint16(view, subtableOffset)

    if (!isUnicodeEncoding(platformId, encodingId)) {
      continue
    }

    if (format === 4 || format === 12) {
      subtables.push({ format, offset: subtableOffset })
    }
  }

  subtables.sort((left, right) => left.format - right.format)

  for (const subtable of subtables) {
    const mappings = subtable.format === 4
      ? parseCmapFormat4(view, subtable.offset)
      : parseCmapFormat12(view, subtable.offset)

    for (const [codepoint, glyphId] of mappings) {
      combined.set(codepoint, glyphId)
    }
  }

  if (combined.size === 0) {
    throw new FontParseError('No supported Unicode cmap subtable found')
  }

  return combined
}

function parseCmapFormat4(view: DataView, offset: number): Map<number, number> {
  assertRange(view, offset, 14, 'cmap format 4 header')
  const length = readUint16(view, offset + 2)
  const segCount = readUint16(view, offset + 6) / 2
  if (segCount < 1) {
    throw new FontParseError('Invalid cmap format 4 segment count')
  }

  assertRange(view, offset, length, 'cmap format 4 table')
  const endCodeOffset = offset + 14
  const startCodeOffset = endCodeOffset + segCount * 2 + 2
  const idDeltaOffset = startCodeOffset + segCount * 2
  const idRangeOffsetOffset = idDeltaOffset + segCount * 2
  const mappings = new Map<number, number>()

  for (let segmentIndex = 0; segmentIndex < segCount; segmentIndex += 1) {
    const endCode = readUint16(view, endCodeOffset + segmentIndex * 2)
    const startCode = readUint16(view, startCodeOffset + segmentIndex * 2)
    const idDelta = readInt16(view, idDeltaOffset + segmentIndex * 2)
    const idRangeOffsetAddress = idRangeOffsetOffset + segmentIndex * 2
    const idRangeOffset = readUint16(view, idRangeOffsetAddress)

    if (startCode === 0xffff && endCode === 0xffff) {
      continue
    }

    if (startCode > endCode) {
      throw new FontParseError('Invalid cmap format 4 segment range')
    }

    for (let codepoint = startCode; codepoint <= endCode; codepoint += 1) {
      let glyphId = 0
      if (idRangeOffset === 0) {
        glyphId = (codepoint + idDelta) & 0xffff
      } else {
        const glyphIndexOffset = idRangeOffsetAddress + idRangeOffset + (codepoint - startCode) * 2
        assertRange(view, glyphIndexOffset, 2, 'cmap format 4 glyph index')
        glyphId = readUint16(view, glyphIndexOffset)
        if (glyphId !== 0) {
          glyphId = (glyphId + idDelta) & 0xffff
        }
      }

      if (glyphId !== 0) {
        mappings.set(codepoint, glyphId)
      }
    }
  }

  return mappings
}

function parseCmapFormat12(view: DataView, offset: number): Map<number, number> {
  assertRange(view, offset, 16, 'cmap format 12 header')
  const length = readUint32(view, offset + 4)
  const groupCount = readUint32(view, offset + 12)
  assertRange(view, offset, length, 'cmap format 12 table')
  const mappings = new Map<number, number>()
  let groupOffset = offset + 16

  for (let index = 0; index < groupCount; index += 1) {
    assertRange(view, groupOffset, 12, 'cmap format 12 group')
    const startCharCode = readUint32(view, groupOffset)
    const endCharCode = readUint32(view, groupOffset + 4)
    const startGlyphId = readUint32(view, groupOffset + 8)

    if (startCharCode > endCharCode) {
      throw new FontParseError('Invalid cmap format 12 group range')
    }

    for (let codepoint = startCharCode; codepoint <= endCharCode; codepoint += 1) {
      mappings.set(codepoint, startGlyphId + (codepoint - startCharCode))
    }

    groupOffset += 12
  }

  return mappings
}

function parseOs2(view: DataView, record: TableRecord): TtfOs2Table {
  assertRange(view, record.offset, OS2_TYPO_DESCENDER_OFFSET + 2, 'OS/2 table')
  const version = readUint16(view, record.offset + OS2_VERSION_OFFSET)
  const sxHeight = version >= 2 && record.length >= OS2_X_HEIGHT_OFFSET + 2
    ? readInt16(view, record.offset + OS2_X_HEIGHT_OFFSET)
    : null
  const sCapHeight = version >= 2 && record.length >= OS2_CAP_HEIGHT_OFFSET + 2
    ? readInt16(view, record.offset + OS2_CAP_HEIGHT_OFFSET)
    : null

  return {
    sTypoAscender: readInt16(view, record.offset + OS2_TYPO_ASCENDER_OFFSET),
    sTypoDescender: readInt16(view, record.offset + OS2_TYPO_DESCENDER_OFFSET),
    sxHeight,
    sCapHeight,
  }
}

function parseName(view: DataView, record: TableRecord): TtfNameTable {
  assertRange(view, record.offset, 6, 'name table header')
  const count = readUint16(view, record.offset + 2)
  const stringOffset = readUint16(view, record.offset + 4)
  const stringsBaseOffset = record.offset + stringOffset
  const candidates: NameCandidate[] = []

  for (let index = 0; index < count; index += 1) {
    const entryOffset = record.offset + 6 + index * 12
    assertRange(view, entryOffset, 12, 'name record')
    const platformId = readUint16(view, entryOffset)
    const encodingId = readUint16(view, entryOffset + 2)
    const languageId = readUint16(view, entryOffset + 4)
    const nameId = readUint16(view, entryOffset + 6)
    const length = readUint16(view, entryOffset + 8)
    const nameOffset = readUint16(view, entryOffset + 10)

    if (nameId !== 1) {
      continue
    }

    const start = stringsBaseOffset + nameOffset
    assertRange(view, start, length, 'name string')
    const bytes = new Uint8Array(view.buffer, start, length)
    const familyName = decodeNameString(bytes, platformId, encodingId)
    if (!familyName) {
      continue
    }

    candidates.push({
      familyName,
      score: nameRecordScore(platformId, encodingId, languageId),
    })
  }

  const bestCandidate = candidates.sort((left, right) => right.score - left.score)[0]
  if (!bestCandidate) {
    throw new FontParseError('Missing font family name')
  }

  return { familyName: bestCandidate.familyName }
}

function nameRecordScore(platformId: number, encodingId: number, languageId: number): number {
  const englishScore = languageId === 0x0409 || languageId === 0 ? 10 : 0
  const windowsUnicodeScore = platformId === 3 && (encodingId === 1 || encodingId === 10) ? 100 : 0
  const unicodeScore = platformId === 0 ? 80 : 0
  const macScore = platformId === 1 ? 40 : 0
  return windowsUnicodeScore + unicodeScore + macScore + englishScore
}

function decodeNameString(bytes: Uint8Array, platformId: number, encodingId: number): string {
  if (platformId === 0 || (platformId === 3 && (encodingId === 1 || encodingId === 10))) {
    return decodeUtf16Be(bytes)
  }

  if (platformId === 1) {
    return decodeAscii(bytes)
  }

  return ''
}

function decodeUtf16Be(bytes: Uint8Array): string {
  if (bytes.length % 2 !== 0) {
    throw new FontParseError('Malformed UTF-16BE string in name table')
  }

  let result = ''
  for (let index = 0; index < bytes.length; index += 2) {
    result += String.fromCharCode((bytes[index] << 8) | bytes[index + 1])
  }
  return result.replace(/\0/g, '').trim()
}

function decodeAscii(bytes: Uint8Array): string {
  return Array.from(bytes, (value) => String.fromCharCode(value)).join('').trim()
}

function isUnicodeEncoding(platformId: number, encodingId: number): boolean {
  return platformId === 0 || (platformId === 3 && (encodingId === 1 || encodingId === 10))
}

function readTag(view: DataView, offset: number): string {
  assertRange(view, offset, 4, 'tag')
  return String.fromCharCode(
    readUint8(view, offset),
    readUint8(view, offset + 1),
    readUint8(view, offset + 2),
    readUint8(view, offset + 3),
  )
}

function assertRange(view: DataView, offset: number, length: number, label: string): void {
  if (offset < 0 || length < 0 || offset + length > view.byteLength) {
    throw new FontParseError(`Malformed ${label}`)
  }
}

function readUint8(view: DataView, offset: number): number {
  assertRange(view, offset, 1, 'uint8 read')
  return view.getUint8(offset)
}

function readUint16(view: DataView, offset: number): number {
  assertRange(view, offset, 2, 'uint16 read')
  return view.getUint16(offset, false)
}

function readInt16(view: DataView, offset: number): number {
  assertRange(view, offset, 2, 'int16 read')
  return view.getInt16(offset, false)
}

function readUint32(view: DataView, offset: number): number {
  assertRange(view, offset, 4, 'uint32 read')
  return view.getUint32(offset, false)
}
