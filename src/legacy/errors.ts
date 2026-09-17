/**
 * Shared error type for the legacy Office (.doc/.ppt) readers (wave-4
 * legacy-office).
 *
 * Every byte this package touches comes from an arbitrary file the user
 * opened — never a format Atlas produced itself — so a malformed or
 * malicious input must fail with a clear, catchable message rather than an
 * uncaught `RangeError`/`TypeError` surfacing from deep inside a parser.
 * Mirrors `docx/parser/unzip.ts`'s `DocxParseError` (same shape, same
 * `Object.setPrototypeOf` restoration for transpiled targets).
 */
export class LegacyFormatError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'LegacyFormatError'
    Object.setPrototypeOf(this, new.target.prototype)
  }
}
