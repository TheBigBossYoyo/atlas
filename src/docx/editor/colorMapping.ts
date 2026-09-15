/**
 * Atlas — shared CSS <-> OOXML color helpers.
 *
 * Two small, independent utilities shared by the toolbar's color/highlight
 * pickers (`toolbarAdapter.ts`) and rich-paste fidelity (`htmlPasteRich.ts`,
 * DXE-19): `toHighlightColor` maps a hex color onto OOXML's fixed
 * `w:highlight` enum (which — unlike run/shading color — has no arbitrary-hex
 * form), and `parseCssColor` normalizes whatever color syntax a pasted
 * `style="color:...; background-color:..."` attribute uses down to a plain
 * `#rrggbb` hex string.
 */

import type { HighlightColor } from '../model'

/**
 * Maps an arbitrary hex color to the nearest OOXML `w:highlight` value.
 * Unlike run/shading color, `w:highlight` is a closed enum (Word's own
 * "text highlight color" swatch), so this is a lookup against exactly the
 * hex values that swatch itself produces — not a general nearest-color
 * search. Returns `null` for anything outside that fixed palette.
 */
export function toHighlightColor(colorHex: string): HighlightColor | null {
  switch (colorHex.trim().toLowerCase()) {
    case '#000000':
      return 'black'
    case '#0000ff':
      return 'blue'
    case '#00ffff':
      return 'cyan'
    case '#00008b':
      return 'darkBlue'
    case '#008b8b':
      return 'darkCyan'
    case '#a9a9a9':
    case '#666666':
      return 'darkGray'
    case '#006400':
      return 'darkGreen'
    case '#8b008b':
      return 'darkMagenta'
    case '#8b0000':
    case '#980000':
      return 'darkRed'
    case '#b8860b':
    case '#ff9900':
      return 'darkYellow'
    case '#00ff00':
      return 'green'
    case '#d3d3d3':
    case '#cccccc':
    case '#d9d9d9':
    case '#efefef':
      return 'lightGray'
    case '#ff00ff':
      return 'magenta'
    case 'transparent':
    case 'none':
      return 'none'
    case '#ff0000':
      return 'red'
    case '#ffffff':
      return 'white'
    case '#ffff00':
      return 'yellow'
    default:
      return null
  }
}

function componentToHex(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0')
}

/**
 * DXE-19 — normalizes a CSS color value (as it appears in a pasted
 * `style="color:...` / `background-color:...` attribute, or a legacy
 * `<font color>` attribute) to a plain `#rrggbb` hex string. Handles
 * `#rgb`/`#rrggbb` and `rgb()`/`rgba()` (alpha is ignored — OOXML run/
 * highlight color has no alpha channel); returns `null` for anything else
 * (named CSS colors, `hsl()`, `currentColor`, etc.) rather than guessing.
 */
export function parseCssColor(value: string): string | null {
  const trimmed = value.trim()

  const hexMatch = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(trimmed)
  if (hexMatch !== null) {
    const hex = hexMatch[1]
    if (hex.length === 3) {
      const [r, g, b] = hex.split('')
      return `#${r}${r}${g}${g}${b}${b}`.toLowerCase()
    }
    return `#${hex.toLowerCase()}`
  }

  const rgbMatch = /^rgba?\(\s*(-?[\d.]+)\s*,\s*(-?[\d.]+)\s*,\s*(-?[\d.]+)\s*(?:,\s*-?[\d.]+\s*)?\)$/i.exec(
    trimmed,
  )
  if (rgbMatch !== null) {
    const [, r, g, b] = rgbMatch
    return `#${componentToHex(Number(r))}${componentToHex(Number(g))}${componentToHex(Number(b))}`
  }

  return null
}
