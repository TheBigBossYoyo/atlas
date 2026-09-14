/** Unit conversions shared by the PPTX (EMU) and ODP (CSS-length) parsers. */

import { parseNumber } from './xmlUtils'

/** 914400 EMU per inch / 96 px per inch. */
export const EMU_PER_PIXEL = 9525

/** 96 px per inch / 2.54 cm per inch. */
export const CENTIMETERS_TO_PIXELS = 37.7952755906

export function emuToPx(value: string | null): number | undefined {
  const parsed = parseNumber(value)
  return parsed === undefined ? undefined : parsed / EMU_PER_PIXEL
}

/** Points -> pixels at 96dpi (1pt = 1/72in). */
export function pointsToPx(points: number): number {
  return points * (96 / 72)
}

/** OOXML angle attributes (`rot`, gradient `ang`) are in 60,000ths of a degree. */
export function ooxmlAngleToDegrees(value: string | null): number | undefined {
  const parsed = parseNumber(value)
  return parsed === undefined ? undefined : parsed / 60000
}

/**
 * OOXML `sz` (font size) attributes are in hundredths of a point.
 * `spcPts` (paragraph spacing) values are also hundredths of a point.
 */
export function hundredthsOfPointToPx(value: string | null): number | undefined {
  const parsed = parseNumber(value)
  return parsed === undefined ? undefined : pointsToPx(parsed / 100)
}

/** ODF (`svg:x`, `fo:font-size`, ...) lengths: `cm`, `mm`, `in`, `pt`, `pc`, or bare `px`. */
export function parseOdfLengthToPixels(value: string | null): number | undefined {
  if (!value) {
    return undefined
  }

  const match = value.trim().match(/^(-?\d*\.?\d+)(cm|mm|in|pt|pc|px)?$/i)
  if (!match) {
    return undefined
  }

  const amount = Number(match[1])
  if (!Number.isFinite(amount)) {
    return undefined
  }

  const unit = (match[2] ?? 'px').toLowerCase()

  switch (unit) {
    case 'cm':
      return amount * CENTIMETERS_TO_PIXELS
    case 'mm':
      return amount * (CENTIMETERS_TO_PIXELS / 10)
    case 'in':
      return amount * 96
    case 'pt':
      return pointsToPx(amount)
    case 'pc':
      return amount * 16
    case 'px':
      return amount
    default:
      return undefined
  }
}
