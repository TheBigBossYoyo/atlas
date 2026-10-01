/**
 * VERSIONS-2 — `encodeTextBytes` (renderer) against `encodeTextBuffer` (main).
 *
 * The two are deliberate mirrors: main has `Buffer`, the renderer has only
 * `TextEncoder`, and the decode half of the pair is already duplicated the same
 * way. Mirrors drift, and this one drifts expensively — the renderer's copy
 * exists so a version captured from what is on screen holds the same bytes a
 * save would have written, and a version is restored by writing its bytes
 * verbatim. If the two disagree, restoring a periodic capture silently rewrites
 * the file's encoding or line endings, which no test of either one alone would
 * notice.
 *
 * So this compares them byte for byte, over every combination the metadata can
 * actually take, rather than testing the renderer's copy against a hand-written
 * expectation.
 */
import { describe, expect, it } from 'vitest'

import { encodeTextBuffer } from '../../electron/lib/textDecoding.cjs'
import {
  DEFAULT_TEXT_FILE_META,
  encodeTextBytes,
  type NewlineStyle,
  type TextEncodingName,
  type TextFileMeta,
} from '../utils/textDecoding'

const ENCODINGS: ReadonlyArray<TextEncodingName> = ['utf-8', 'utf-16le', 'utf-16be', 'windows-1252']
const NEWLINES: ReadonlyArray<NewlineStyle> = ['lf', 'crlf']

const SAMPLES: ReadonlyArray<readonly [string, string]> = [
  ['empty', ''],
  ['ascii, no line breaks', 'The opening paragraph.'],
  ['multi-line', '# Title\n\nLine one.\nLine two.\n'],
  ['latin-1 accents', 'héllo wörld — ça va\n'],
  // The cp1252 high range (0x80-0x9F): representable there, and a different
  // code point in Unicode, which is exactly where a reverse map goes wrong.
  ['cp1252 high range', '€ “ quoted ” … ™\n'],
  // Not representable in cp1252 at all, which must take the UTF-8 fallback.
  ['beyond cp1252', 'ал­фавит 日本語 \u{1F600}\n'],
  ['lone CR kept as content', 'a\rb\n'],
]

/** Every metadata combination the decoder can actually produce. */
function metas(): ReadonlyArray<TextFileMeta> {
  const out: TextFileMeta[] = []
  for (const encoding of ENCODINGS) {
    for (const newline of NEWLINES) {
      // Windows-1252 has no BOM convention, and `decodeTextBufferWithMeta`
      // never reports one for it — so `{ windows-1252, bom: true }` is not a
      // state that exists, and pinning the two implementations against each
      // other on it would pin undefined behaviour.
      const boms = encoding === 'windows-1252' ? [false] : [false, true]
      for (const bom of boms) out.push({ encoding, bom, newline })
    }
  }
  return out
}

describe('encodeTextBytes mirrors main encodeTextBuffer (VERSIONS-2)', () => {
  for (const [label, content] of SAMPLES) {
    for (const meta of metas()) {
      const name = `${label} as ${meta.encoding}${meta.bom ? ' with BOM' : ''}, ${meta.newline}`
      it(name, () => {
        const { buffer } = encodeTextBuffer(content, meta)
        const mine = encodeTextBytes(content, meta)
        // Compared as plain arrays so a mismatch prints the bytes, not
        // "Uint8Array !== Buffer" — the two are different classes by design.
        expect(Array.from(mine)).toEqual(Array.from(buffer))
      })
    }
  }

  it('encodes the default metadata as plain UTF-8', () => {
    // The fallback every never-saved document uses, spelled out so a change to
    // DEFAULT_TEXT_FILE_META that broke it would not pass silently.
    expect(Array.from(encodeTextBytes('abc\n', DEFAULT_TEXT_FILE_META))).toEqual([97, 98, 99, 10])
  })

  it('would notice a drift between the two', () => {
    // The self-check: this whole file is a comparison, so it is only worth
    // anything if a difference actually fails it. A deliberately wrong newline
    // on one side must produce different bytes.
    const content = 'one\ntwo\n'
    const asCrlf = encodeTextBytes(content, { encoding: 'utf-8', bom: false, newline: 'crlf' })
    const { buffer: asLf } = encodeTextBuffer(content, { encoding: 'utf-8', bom: false, newline: 'lf' })
    expect(Array.from(asCrlf)).not.toEqual(Array.from(asLf))
  })
})
