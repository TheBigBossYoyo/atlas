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
//
// NIGHT/text-roundtrip — round-trip fidelity. None of the above was ever
// remembered past the initial decode: every save re-encoded as plain,
// BOM-less UTF-8 regardless of what was actually on disk (SHELL-1/SHELL-2 —
// a CRLF markdown file loses its `\r`s, a UTF-8-BOM file loses its BOM; the
// same shape of bug applies to a genuinely UTF-16/Windows-1252 file, which
// gets silently transcoded to UTF-8 on first save). `decodeTextBufferWithMeta`
// is the fixed entry point: same decoded (newline-normalized) string, plus
// `{ encoding, bom, newline }` so a save can call `encodeTextBuffer(content,
// meta)` to reproduce the original byte shape for everything the user didn't
// touch. Content is always normalized to `\n`-only here (matching what every
// DOM text-input widget the app uses — a `<textarea>`, CodeMirror — already
// reduces `\r\n`/`\r` to on read/edit) so the *editor's* string model is
// deterministic; `newline` records what to convert back to on save.
// `decodeTextBuffer` is kept as a thin wrapper (content only, NOT
// newline-normalized — unchanged from before) for callers with no save path
// to round-trip (e.g. `UnknownViewer`'s best-effort text preview).

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

// ---------------------------------------------------------------------------
// NIGHT/text-roundtrip additions
// ---------------------------------------------------------------------------

/** Reverse of `WINDOWS_1252_HIGH_RANGE`: Unicode code point -> the single cp1252 byte (0x80-0x9F) that maps to it. Built once at module load. */
const WINDOWS_1252_REVERSE = new Map();
WINDOWS_1252_HIGH_RANGE.forEach((codePoint, index) => {
  // Two of the 32 slots (0x81, 0x8D, 0x8F, 0x90, 0x9D are unassigned and use
  // Latin-1 identity already handled by the `byte <= 0xff` fast path below)
  // — only register the ones with a *distinct* high code point so the
  // identity range isn't shadowed.
  if (!WINDOWS_1252_REVERSE.has(codePoint)) {
    WINDOWS_1252_REVERSE.set(codePoint, 0x80 + index);
  }
});

/**
 * @param {string} content
 * @returns {Buffer | null} the cp1252-encoded bytes, or `null` if any
 *   character in `content` has no Windows-1252 representation.
 */
function encodeWindows1252(content) {
  const bytes = Buffer.alloc(content.length);
  for (let i = 0; i < content.length; i += 1) {
    const codePoint = content.charCodeAt(i);
    if (codePoint <= 0xff && !(codePoint >= 0x80 && codePoint <= 0x9f)) {
      // 0x00-0x7F and 0xA0-0xFF are Windows-1252 identity with Unicode.
      bytes[i] = codePoint;
      continue;
    }
    const mapped = WINDOWS_1252_REVERSE.get(codePoint);
    if (mapped === undefined) {
      return null;
    }
    bytes[i] = mapped;
  }
  return bytes;
}

/**
 * Normalizes every line ending in `content` to `\n` (CRLF and lone CR both
 * collapse to LF), matching what a `<textarea>`/CodeMirror already does to
 * anything typed or pasted into it.
 * @param {string} content
 * @returns {string}
 */
function normalizeNewlines(content) {
  return content.replace(/\r\n|\r/g, '\n');
}

/**
 * Detects the dominant line-ending convention of already-decoded text.
 * Mixed files (both `\r\n` and lone `\n` present) pick whichever is more
 * frequent — a deliberate, documented choice (not left as an accident of
 * whatever ran last): on a tie, or a file with no line breaks at all, `lf`
 * wins, since that's what a brand-new Atlas document already uses and it's
 * the narrower transformation of the two (every `\r\n` this session's own
 * `<textarea>`/CodeMirror edits produce is already `\n`, so converting the
 * REST of a tied-mixed file to `\n` too changes fewer bytes than the other
 * way around).
 * @param {string} content
 * @returns {'crlf' | 'lf'}
 */
function detectNewline(content) {
  let crlf = 0;
  let lf = 0;
  for (let i = 0; i < content.length; i += 1) {
    if (content[i] === '\n') {
      if (i > 0 && content[i - 1] === '\r') {
        crlf += 1;
      } else {
        lf += 1;
      }
    }
  }
  if (crlf === 0) return 'lf';
  if (lf === 0) return 'crlf';
  return crlf > lf ? 'crlf' : 'lf';
}

/**
 * @typedef {{
 *   encoding: 'utf-8' | 'utf-16le' | 'utf-16be' | 'windows-1252';
 *   bom: boolean;
 *   newline: 'crlf' | 'lf';
 * }} TextFileMeta
 */

/**
 * Decodes a text-class file's raw bytes AND remembers everything needed to
 * reproduce its original byte shape on save (SHELL-1/SHELL-2 fix). `content`
 * is always `\n`-normalized (see module header); `meta.newline` is detected
 * from the ORIGINAL bytes, before normalization.
 * @param {Buffer} buffer
 * @returns {{ content: string; meta: TextFileMeta }}
 */
function decodeTextBufferWithMeta(buffer) {
  /** @type {TextFileMeta['encoding']} */
  let encoding;
  let bom;
  let decoded;

  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    encoding = 'utf-8';
    bom = true;
    decoded = buffer.toString('utf8', 3);
  } else if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
    encoding = 'utf-16le';
    bom = true;
    decoded = buffer.toString('utf16le', 2);
  } else if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) {
    encoding = 'utf-16be';
    bom = true;
    decoded = decodeUtf16Be(buffer.subarray(2));
  } else if (isValidUtf8(buffer)) {
    encoding = 'utf-8';
    bom = false;
    decoded = buffer.toString('utf8');
  } else {
    encoding = 'windows-1252';
    bom = false;
    decoded = decodeWindows1252(buffer);
  }

  return {
    content: normalizeNewlines(decoded),
    meta: { encoding, bom, newline: detectNewline(decoded) },
  };
}

/**
 * Inverse of `decodeTextBufferWithMeta`: re-encodes `content` (assumed
 * `\n`-normalized, i.e. the editor's own string model) back to `meta`'s
 * encoding/BOM/newline convention.
 *
 * Windows-1252: only used when every character in `content` is actually
 * representable in it (checked here, not assumed from `meta`, since an edit
 * may have typed a character the original file's charset never had — e.g.
 * an emoji into a cp1252 French text file). When that happens this falls
 * back to plain UTF-8 (no BOM) instead of either throwing or silently
 * mangling the unrepresentable character, and reports the fallback via
 * `encodingFallback` so the caller can surface it. (No i18n string exists
 * for that yet — this repo's owner needs a decision on the exact English/
 * French copy; see this fix's report.)
 * @param {string} content
 * @param {TextFileMeta} meta
 * @returns {{ buffer: Buffer; encodingFallback: boolean }}
 */
function encodeTextBuffer(content, meta) {
  const withNewline = meta.newline === 'crlf' ? content.replace(/\n/g, '\r\n') : content;

  if (meta.encoding === 'utf-16le') {
    const body = Buffer.from(withNewline, 'utf16le');
    return { buffer: meta.bom ? Buffer.concat([Buffer.from([0xff, 0xfe]), body]) : body, encodingFallback: false };
  }

  if (meta.encoding === 'utf-16be') {
    const le = Buffer.from(withNewline, 'utf16le');
    const be = Buffer.from(le);
    for (let i = 0; i + 1 < be.length; i += 2) {
      const hi = be[i];
      be[i] = be[i + 1];
      be[i + 1] = hi;
    }
    return { buffer: meta.bom ? Buffer.concat([Buffer.from([0xfe, 0xff]), be]) : be, encodingFallback: false };
  }

  if (meta.encoding === 'windows-1252') {
    const cp1252 = encodeWindows1252(withNewline);
    if (cp1252 !== null) {
      // Windows-1252 has no BOM convention — `meta.bom` is never true for it
      // (decodeTextBufferWithMeta never sets it), nothing to prepend.
      return { buffer: cp1252, encodingFallback: false };
    }
    // Not representable in cp1252 (the user typed something the original
    // charset can't hold) — fall back to honest, lossless UTF-8.
    return { buffer: Buffer.from(withNewline, 'utf-8'), encodingFallback: true };
  }

  const body = Buffer.from(withNewline, 'utf-8');
  return { buffer: meta.bom ? Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), body]) : body, encodingFallback: false };
}

module.exports = {
  decodeTextBuffer,
  decodeTextBufferWithMeta,
  encodeTextBuffer,
  isValidUtf8,
  decodeWindows1252,
  detectNewline,
  normalizeNewlines,
};
