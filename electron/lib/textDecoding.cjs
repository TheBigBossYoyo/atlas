// Text-class file decoding with BOM sniffing and a Windows-1252 fallback
// (LOAD-05/DAT-03/RUN-05).
//
// Order of operations:
//   1. UTF-8 BOM (EF BB BF)        -> decode the rest as UTF-8, BOM stripped.
//   2. UTF-16 LE BOM (FF FE)       -> decode the rest as UTF-16 LE.
//   3. UTF-16 BE BOM (FE FF)       -> byte-swap then decode as UTF-16 LE.
//   4. No BOM, valid UTF-8         -> decode as UTF-8 (matches today's
//      `fs.readFileSync(path, 'utf-8')` behavior byte-for-byte for the
//      overwhelmingly common no-BOM-valid-UTF-8 case, so markdown's existing
//      decoding is unchanged).
//   5. No BOM, invalid UTF-8       -> fall back to Windows-1252 instead of
//      emitting U+FFFD replacement-character mojibake.

// Windows-1252 code points for bytes 0x80-0x9F (0xA0-0xFF and 0x00-0x7F are
// identical to their Unicode code points). The five bytes with no assigned
// Windows-1252 character (0x81, 0x8D, 0x8F, 0x90, 0x9D) fall back to their
// raw byte value (Latin-1 identity) rather than throwing, since this is
// already a best-effort fallback path.
const WINDOWS_1252_HIGH_RANGE = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039,
  0x0152, 0x008d, 0x017d, 0x008f, 0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
];

/**
 * @param {Buffer} buffer
 * @returns {boolean}
 */
function isValidUtf8(buffer) {
  const len = buffer.length;
  let i = 0;

  while (i < len) {
    const byte1 = buffer[i];

    if (byte1 <= 0x7f) {
      i += 1;
      continue;
    }

    let extraBytes;
    let codePoint;
    let minCodePoint;

    if ((byte1 & 0xe0) === 0xc0) {
      extraBytes = 1;
      codePoint = byte1 & 0x1f;
      minCodePoint = 0x80;
    } else if ((byte1 & 0xf0) === 0xe0) {
      extraBytes = 2;
      codePoint = byte1 & 0x0f;
      minCodePoint = 0x800;
    } else if ((byte1 & 0xf8) === 0xf0) {
      extraBytes = 3;
      codePoint = byte1 & 0x07;
      minCodePoint = 0x10000;
    } else {
      return false;
    }

    if (i + extraBytes >= len) {
      return false;
    }

    for (let j = 1; j <= extraBytes; j += 1) {
      const continuationByte = buffer[i + j];
      if ((continuationByte & 0xc0) !== 0x80) {
        return false;
      }
      codePoint = (codePoint << 6) | (continuationByte & 0x3f);
    }

    if (codePoint < minCodePoint || codePoint > 0x10ffff) {
      return false;
    }
    if (codePoint >= 0xd800 && codePoint <= 0xdfff) {
      return false;
    }

    i += extraBytes + 1;
  }

  return true;
}

/**
 * @param {Buffer} buffer
 * @returns {string}
 */
function decodeWindows1252(buffer) {
  let result = '';
  for (let i = 0; i < buffer.length; i += 1) {
    const byte = buffer[i];
    if (byte >= 0x80 && byte <= 0x9f) {
      result += String.fromCharCode(WINDOWS_1252_HIGH_RANGE[byte - 0x80]);
    } else {
      result += String.fromCharCode(byte);
    }
  }
  return result;
}

/**
 * @param {Buffer} buffer
 * @returns {string}
 */
function decodeUtf16Be(buffer) {
  const swapped = Buffer.from(buffer);
  for (let i = 0; i + 1 < swapped.length; i += 2) {
    const hi = swapped[i];
    swapped[i] = swapped[i + 1];
    swapped[i + 1] = hi;
  }
  return swapped.toString('utf16le');
}

/**
 * Decodes a text-class file's raw bytes, sniffing a BOM and falling back to
 * Windows-1252 for non-UTF-8 content with no BOM.
 * @param {Buffer} buffer
 * @returns {string}
 */
function decodeTextBuffer(buffer) {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    return buffer.toString('utf8', 3);
  }

  if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
    return buffer.toString('utf16le', 2);
  }

  if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) {
    return decodeUtf16Be(buffer.subarray(2));
  }

  if (isValidUtf8(buffer)) {
    return buffer.toString('utf8');
  }

  return decodeWindows1252(buffer);
}

module.exports = { decodeTextBuffer, isValidUtf8, decodeWindows1252 };
