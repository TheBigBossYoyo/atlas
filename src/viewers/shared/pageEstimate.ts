/**
 * Approximate page-count estimation for RTF/ODT (T7/DAT-17).
 *
 * Neither `rtf.js` nor `odf-kit` reports a page count — RTF has no such
 * concept in its parsed form, and ODT's page count is purely a function of
 * how the flowed text reflows at render time (not stored in the file at
 * all). Both viewers previously hardcoded `pages: 1`, which is simply wrong
 * for any multi-page document. The best available proxy without a full
 * pagination engine is: render height ÷ one page's height, rounded to the
 * nearest whole page (a partially-filled last page still counts as one).
 *
 * `pageHeightPx` should come from the document's own page layout when known
 * (`parseLengthToPx(doc.pageLayout?.height)` for ODT); `DEFAULT_PAGE_HEIGHT_PX`
 * (US Letter at 96dpi) is the fallback for RTF, which carries no page-layout
 * metadata at all, and for ODT documents with no page-layout element.
 */

const PX_PER_INCH = 96
const CM_PER_INCH = 2.54
const MM_PER_INCH = 25.4
const PT_PER_INCH = 72

/** US Letter (8.5in × 11in) full page height at 96dpi — the fallback when no page-layout metadata is available. */
export const DEFAULT_PAGE_HEIGHT_PX = 11 * PX_PER_INCH

/**
 * Converts a CSS-ready ODF length string ("29.7cm", "11in", "297mm", "842pt")
 * to a pixel value at 96dpi. Returns `undefined` for an absent or
 * unrecognized value rather than guessing.
 */
export function parseLengthToPx(value: string | undefined): number | undefined {
  if (!value) return undefined

  const match = /^([\d.]+)\s*(cm|mm|in|pt|px)$/i.exec(value.trim())
  if (!match) return undefined

  const amount = Number(match[1])
  if (!Number.isFinite(amount)) return undefined

  switch (match[2].toLowerCase()) {
    case 'cm':
      return (amount / CM_PER_INCH) * PX_PER_INCH
    case 'mm':
      return (amount / MM_PER_INCH) * PX_PER_INCH
    case 'in':
      return amount * PX_PER_INCH
    case 'pt':
      return (amount / PT_PER_INCH) * PX_PER_INCH
    case 'px':
      return amount
    default:
      return undefined
  }
}

/**
 * Estimates a page count from a rendered element's height. Always returns at
 * least 1 (an empty or unmeasurable document is still "one page" for display
 * purposes, matching the previous hardcoded behavior for the common case).
 */
export function estimatePageCount(scrollHeightPx: number, pageHeightPx: number = DEFAULT_PAGE_HEIGHT_PX): number {
  if (!(scrollHeightPx > 0) || !(pageHeightPx > 0)) {
    return 1
  }
  return Math.max(1, Math.round(scrollHeightPx / pageHeightPx))
}
