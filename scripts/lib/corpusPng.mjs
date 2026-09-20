// Minimal, dependency-free PNG encoder used by the DOCX corpus generator
// (scripts/generate-docx-corpus.mjs) to embed tiny, fully deterministic
// raster images without checking binary assets into the repo.
import { deflateSync } from 'node:zlib'

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])
const CRC_TABLE = buildCrcTable()

function buildCrcTable() {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    }
    table[n] = c >>> 0
  }
  return table
}

/**
 * @param {Buffer} buf
 * @returns {number}
 */
function crc32(buf) {
  let crc = 0xffffffff
  for (const byte of buf) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8)
  }
  return (crc ^ 0xffffffff) >>> 0
}

/**
 * @param {string} type
 * @param {Buffer} data
 * @returns {Buffer}
 */
function pngChunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length, 0)
  const typeBuf = Buffer.from(type, 'ascii')
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0)
  return Buffer.concat([length, typeBuf, data, crc])
}

/**
 * Builds a tiny, deterministic 8-bit RGB PNG filled with a solid color.
 * Deterministic across runs given the same Node/zlib version: no timestamps,
 * no randomness, fixed compression level.
 *
 * @param {number} width
 * @param {number} height
 * @param {readonly [number, number, number]} rgb
 * @returns {Buffer}
 */
export function createSolidPng(width, height, [r, g, b]) {
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // color type: truecolor (RGB)
  ihdr[10] = 0 // compression method
  ihdr[11] = 0 // filter method
  ihdr[12] = 0 // interlace method

  const rowSize = 1 + width * 3
  const raw = Buffer.alloc(rowSize * height)
  for (let y = 0; y < height; y++) {
    const rowStart = y * rowSize
    raw[rowStart] = 0 // per-row filter: none
    for (let x = 0; x < width; x++) {
      const pixelStart = rowStart + 1 + x * 3
      raw[pixelStart] = r
      raw[pixelStart + 1] = g
      raw[pixelStart + 2] = b
    }
  }

  const idat = deflateSync(raw, { level: 9 })

  return Buffer.concat([
    PNG_SIGNATURE,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', idat),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}
