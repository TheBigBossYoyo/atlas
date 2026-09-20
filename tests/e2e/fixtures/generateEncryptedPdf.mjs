// Builds a genuinely RC4-40 (Standard Security Handler, V1/R2) encrypted PDF
// from scratch, per PDF 1.7 spec (ISO 32000-1) section 7.6. No third-party
// PDF/crypto-for-PDF library is available in this repo, so this implements
// Algorithms 2/3/4 directly with Node's built-in `crypto` (md5 + rc4).
import crypto from 'node:crypto'
import fs from 'node:fs'

const PADDING = Buffer.from([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
])

/** @param {string} pw @returns {Buffer} */
function padPassword(pw) {
  const buf = Buffer.from(pw, 'latin1').subarray(0, 32)
  return Buffer.concat([buf, PADDING], 32)
}

/** @param {...Buffer} parts @returns {Buffer} */
function md5(...parts) {
  const h = crypto.createHash('md5')
  for (const p of parts) h.update(p)
  return h.digest()
}

// Node's OpenSSL 3 default provider dropped RC4 from `crypto.getCiphers()`,
// so this is a plain from-scratch RC4 (it's a simple stream cipher — no
// external library needed, and pdf.js's own decrypt path implements the
// exact same algorithm to read it back).
/**
 * @param {Buffer | Uint8Array} key
 * @param {Buffer} data
 * @returns {Buffer}
 */
function rc4(key, data) {
  const s = new Uint8Array(256)
  for (let i = 0; i < 256; i++) s[i] = i
  let j = 0
  for (let i = 0; i < 256; i++) {
    j = (j + s[i] + key[i % key.length]) & 0xff
    ;[s[i], s[j]] = [s[j], s[i]]
  }
  const out = Buffer.alloc(data.length)
  let i = 0
  j = 0
  for (let k = 0; k < data.length; k++) {
    i = (i + 1) & 0xff
    j = (j + s[i]) & 0xff
    ;[s[i], s[j]] = [s[j], s[i]]
    out[k] = data[k] ^ s[(s[i] + s[j]) & 0xff]
  }
  return out
}

// Revision 2, 40-bit (5-byte) key, per Algorithm 3.2.
/**
 * @param {string} userPw
 * @param {Buffer} ownerHashO
 * @param {number} permissionsP
 * @param {Buffer} idBytes
 * @returns {Buffer}
 */
function computeEncryptionKey(userPw, ownerHashO, permissionsP, idBytes) {
  const padded = padPassword(userPw)
  const pBuf = Buffer.alloc(4)
  pBuf.writeInt32LE(permissionsP, 0)
  const hash = md5(padded, ownerHashO, pBuf, idBytes)
  return hash.subarray(0, 5) // 40-bit key
}

// Algorithm 3.3 (rev 2): owner password (or user pw if owner blank).
/**
 * @param {string} ownerPw
 * @param {string} userPw
 * @returns {Buffer}
 */
function computeO(ownerPw, userPw) {
  const ownerPadded = padPassword(ownerPw || userPw)
  const key = md5(ownerPadded).subarray(0, 5)
  const userPadded = padPassword(userPw)
  return rc4(key, userPadded)
}

// Algorithm 3.4 (rev 2): U = RC4(fileKey, PADDING)
/** @param {Buffer} fileKey @returns {Buffer} */
function computeU(fileKey) {
  return rc4(fileKey, PADDING)
}

/**
 * @param {Buffer} fileKey
 * @param {number} objNum
 * @param {number} genNum
 * @returns {Buffer}
 */
function objectKey(fileKey, objNum, genNum) {
  const extra = Buffer.alloc(5)
  extra.writeUIntLE(objNum, 0, 3)
  extra.writeUIntLE(genNum, 3, 2)
  const hash = md5(fileKey, extra)
  const n = Math.min(fileKey.length + 5, 16)
  return hash.subarray(0, n)
}

/**
 * @param {Buffer} fileKey
 * @param {number} objNum
 * @param {number} genNum
 * @param {Buffer} data
 * @returns {Buffer}
 */
function encryptForObject(fileKey, objNum, genNum, data) {
  return rc4(objectKey(fileKey, objNum, genNum), data)
}

/** @param {Buffer} buf @returns {string} */
function pdfString(buf) {
  // Literal string with octal-escaped bytes >= 0x80 or special chars — safe or generic.
  let out = '('
  for (const byte of buf) {
    if (byte === 0x28 || byte === 0x29 || byte === 0x5c) {
      out += '\\' + String.fromCharCode(byte)
    } else if (byte < 0x20 || byte > 0x7e) {
      out += '\\' + byte.toString(8).padStart(3, '0')
    } else {
      out += String.fromCharCode(byte)
    }
  }
  return out + ')'
}

/**
 * @param {{ userPassword: string, ownerPassword: string, permissions?: number }} options
 * @returns {Buffer}
 */
function buildEncryptedPdf({ userPassword, ownerPassword, permissions = -4 }) {
  const idBytes = crypto.randomBytes(16)
  const O = computeO(ownerPassword, userPassword)
  const fileKey = computeEncryptionKey(userPassword, O, permissions, idBytes)
  const U = computeU(fileKey)

  const contentText = 'BT /F1 24 Tf 72 700 Td (Secret content: unlock me) Tj ET'
  const contentBuf = Buffer.from(contentText, 'latin1')

  // Object numbers: 1 catalog, 2 pages, 3 page, 4 content stream, 5 font, 6 encrypt dict (not encrypted itself)
  const objects = []

  objects[1] = `<< /Type /Catalog /Pages 2 0 R >>`
  objects[2] = `<< /Type /Pages /Kids [3 0 R] /Count 1 >>`
  objects[3] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>`

  const encryptedContent = encryptForObject(fileKey, 4, 0, contentBuf)
  objects[4] = { stream: encryptedContent, dict: `<< /Length ${encryptedContent.length} >>` }

  objects[5] = `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>`

  const encryptDictBody = `<< /Filter /Standard /V 1 /R 2 /O ${pdfString(O)} /U ${pdfString(U)} /P ${permissions} >>`

  let pdf = '%PDF-1.4\n%\xE2\xE3\xCF\xD3\n'
  const offsets = [0]

  /** @param {number} num @param {string} body */
  function pushObj(num, body) {
    offsets[num] = Buffer.byteLength(pdf, 'latin1')
    pdf += `${num} 0 obj\n${body}\nendobj\n`
  }

  pushObj(1, objects[1])
  pushObj(2, objects[2])
  pushObj(3, objects[3])

  offsets[4] = Buffer.byteLength(pdf, 'latin1')
  pdf += `4 0 obj\n${objects[4].dict}\nstream\n`
  const head = Buffer.from(pdf, 'latin1')
  const full = Buffer.concat([head, objects[4].stream, Buffer.from('\nendstream\nendobj\n', 'latin1')])
  pdf = full.toString('latin1')

  pushObj(5, objects[5])

  offsets[6] = Buffer.byteLength(pdf, 'latin1')
  pdf += `6 0 obj\n${encryptDictBody}\nendobj\n`

  const xrefOffset = Buffer.byteLength(pdf, 'latin1')
  const total = 7
  pdf += `xref\n0 ${total}\n`
  pdf += '0000000000 65535 f \n'
  for (let i = 1; i < total; i++) {
    pdf += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`
  }
  const idHex = idBytes.toString('hex')
  pdf += `trailer\n<< /Size ${total} /Root 1 0 R /Encrypt 6 0 R /ID [<${idHex}> <${idHex}>] >>\nstartxref\n${xrefOffset}\n%%EOF\n`

  return Buffer.from(pdf, 'latin1')
}

const outPath = process.argv[2] || 'encrypted.pdf'
const userPassword = process.argv[3] || 'secret123'
const ownerPassword = process.argv[4] || 'owner456'

const bytes = buildEncryptedPdf({ userPassword, ownerPassword })
fs.writeFileSync(outPath, bytes)
console.log('wrote', outPath, bytes.length, 'bytes; userPassword=', userPassword)
