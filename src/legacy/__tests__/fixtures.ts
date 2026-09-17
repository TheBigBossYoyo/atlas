/**
 * Shared in-memory fixture builders for legacy .doc/.ppt tests.
 *
 * Every fixture here is a REAL OLE2/CFB container, built via `XLSX.CFB.utils`
 * (the exact codec `src/legacy/cfb.ts` wraps) and round-tripped through
 * `CFB.write`/`CFB.read` the same way an actual file on disk would be —
 * never a hand-rolled byte layout pretending to be one. Kept well under the
 * ~50KB fixture-size guidance (the largest of these is a few hundred bytes).
 *
 * Not a `*.test.ts` file itself, so Vitest's `include` glob never runs it as
 * a suite — just a helper module imported by the real test files:
 * `cfb.test.ts`, `doc/__tests__/index.test.ts`, `ppt/__tests__/slides.test.ts`,
 * and the `LegacyDocViewer`/`LegacyPptViewer` render tests.
 */
import * as XLSX from 'xlsx'

export function buildCfbBytes(streams: ReadonlyArray<readonly [string, Uint8Array]>): Uint8Array {
  const cfb = XLSX.CFB.utils.cfb_new()
  for (const [name, bytes] of streams) {
    XLSX.CFB.utils.cfb_add(cfb, name, bytes)
  }
  return new Uint8Array(XLSX.CFB.write(cfb, { type: 'array' }) as ArrayLike<number>)
}

// ---------------------------------------------------------------------------
// .doc fixtures
// ---------------------------------------------------------------------------

const FC_COMPRESSED_FLAG = 0x40000000
// Body text starts safely after every fixed FIB field `parseFib` reads (the
// furthest is `lcbClx` at byte offset 422).
const FIB_STREAM_HEADROOM = 512

function buildWordDocumentStream(bodyText: string): Uint8Array {
  const buffer = new ArrayBuffer(FIB_STREAM_HEADROOM + bodyText.length)
  const view = new DataView(buffer)

  view.setUint16(0, 0xa5ec, true) // wIdent
  view.setUint16(0x0a, 0, true) // flags1 -> fWhichTblStm clear -> "0Table"
  view.setUint32(76, bodyText.length, true) // ccpText
  view.setUint32(418, 0, true) // fcClx (into the Table stream)
  // lcbClx (offset 422) is patched in by buildMinimalDocBytes once the Table
  // stream's size is known.

  const bytes = new Uint8Array(buffer)
  bytes.set(Uint8Array.from(bodyText, (c) => c.charCodeAt(0)), FIB_STREAM_HEADROOM)
  return bytes
}

/** A Table stream holding a CLX with exactly one compressed (cp1252) piece covering `numChars` characters starting at `byteOffset`. */
function buildSinglePieceTableStream(numChars: number, byteOffset: number): Uint8Array {
  const plcPcdSize = 2 * 4 + 8 // 2 CPs + 1 PCD
  const buffer = new ArrayBuffer(1 + 4 + plcPcdSize)
  const view = new DataView(buffer)

  view.setUint8(0, 0x02) // Pcdt block
  view.setUint32(1, plcPcdSize, true)
  view.setUint32(5, 0, true) // cp[0]
  view.setUint32(9, numChars, true) // cp[1]
  view.setUint16(13, 0, true) // pcd.flags (unused by the reader)
  view.setUint32(15, (byteOffset * 2) | FC_COMPRESSED_FLAG, true) // pcd.fc (compressed)
  view.setUint16(19, 0, true) // pcd.prm (unused by the reader)

  return new Uint8Array(buffer)
}

/** A minimal but complete, well-formed .doc (as CFB bytes) whose main story is exactly `bodyText`, cp1252-encoded. */
export function buildMinimalDocBytes(bodyText: string): Uint8Array {
  const wordDocument = buildWordDocumentStream(bodyText)
  const tableStream = buildSinglePieceTableStream(bodyText.length, FIB_STREAM_HEADROOM)

  new DataView(wordDocument.buffer).setUint32(422, tableStream.length, true) // lcbClx

  return buildCfbBytes([
    ['WordDocument', wordDocument],
    ['0Table', tableStream],
  ])
}

// ---------------------------------------------------------------------------
// .ppt fixtures
// ---------------------------------------------------------------------------

const RT_TEXT_HEADER_ATOM = 3999
const RT_TEXT_CHARS_ATOM = 4000
const RT_DOCUMENT = 1000
const RT_SLIDE = 1006

function atomRecord(recVer: number, type: number, payload: Uint8Array): Uint8Array {
  const record = new Uint8Array(8 + payload.length)
  const view = new DataView(record.buffer)
  view.setUint16(0, recVer & 0x000f, true)
  view.setUint16(2, type, true)
  view.setUint32(4, payload.length, true)
  record.set(payload, 8)
  return record
}

function concatBytes(arrays: ReadonlyArray<Uint8Array>): Uint8Array {
  const total = arrays.reduce((sum, a) => sum + a.length, 0)
  const out = new Uint8Array(total)
  let offset = 0
  for (const a of arrays) {
    out.set(a, offset)
    offset += a.length
  }
  return out
}

function containerRecord(type: number, children: ReadonlyArray<Uint8Array>): Uint8Array {
  return atomRecord(0x0f, type, concatBytes(children))
}

function textHeaderAtom(textType: number): Uint8Array {
  const payload = new Uint8Array(4)
  new DataView(payload.buffer).setInt32(0, textType, true)
  return atomRecord(0x0, RT_TEXT_HEADER_ATOM, payload)
}

function textCharsAtom(text: string): Uint8Array {
  const bytes = new Uint8Array(new Uint16Array([...text].map((c) => c.charCodeAt(0))).buffer)
  return atomRecord(0x0, RT_TEXT_CHARS_ATOM, bytes)
}

export type PptFixtureSlide = {
  readonly title?: string
  readonly body?: string
}

/** One Slide container per entry in `slides`, each with a Title (textType 0) and/or Body (textType 1) text run. */
function buildPptDocumentStream(slides: ReadonlyArray<PptFixtureSlide>): Uint8Array {
  const slideRecords = slides.map((slide) => {
    const children: Uint8Array[] = []
    if (slide.title !== undefined) {
      children.push(textHeaderAtom(0), textCharsAtom(slide.title))
    }
    if (slide.body !== undefined) {
      children.push(textHeaderAtom(1), textCharsAtom(slide.body))
    }
    return containerRecord(RT_SLIDE, children)
  })

  return containerRecord(RT_DOCUMENT, slideRecords)
}

/** A minimal but complete, well-formed .ppt (as CFB bytes) with one Slide container per entry in `slides`. */
export function buildMinimalPptBytes(slides: ReadonlyArray<PptFixtureSlide>): Uint8Array {
  return buildCfbBytes([['PowerPoint Document', buildPptDocumentStream(slides)]])
}
