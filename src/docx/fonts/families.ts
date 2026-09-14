// Font URLs are imported via Vite's `?url` suffix so the bundler emits each TTF
// to dist/assets/ with a hashed filename and rewrites the URL to be RELATIVE to
// the bundled index.html. This is required for the packaged Electron app where
// CSS-absolute URLs like `/fonts/X.ttf` resolve to `file:///C:/fonts/X.ttf`
// (which does not exist) instead of the asset location inside the asar bundle.

import CarlitoRegular from '/fonts/Carlito-Regular.ttf?url'
import CarlitoBold from '/fonts/Carlito-Bold.ttf?url'
import CarlitoItalic from '/fonts/Carlito-Italic.ttf?url'
import CarlitoBoldItalic from '/fonts/Carlito-BoldItalic.ttf?url'

import CaladeaRegular from '/fonts/Caladea-Regular.ttf?url'
import CaladeaBold from '/fonts/Caladea-Bold.ttf?url'
import CaladeaItalic from '/fonts/Caladea-Italic.ttf?url'
import CaladeaBoldItalic from '/fonts/Caladea-BoldItalic.ttf?url'

import CourierPrimeRegular from '/fonts/CourierPrime-Regular.ttf?url'
import CourierPrimeBold from '/fonts/CourierPrime-Bold.ttf?url'
import CourierPrimeItalic from '/fonts/CourierPrime-Italic.ttf?url'
import CourierPrimeBoldItalic from '/fonts/CourierPrime-BoldItalic.ttf?url'

import TinosRegular from '/fonts/Tinos-Regular.ttf?url'
import TinosBold from '/fonts/Tinos-Bold.ttf?url'
import TinosItalic from '/fonts/Tinos-Italic.ttf?url'
import TinosBoldItalic from '/fonts/Tinos-BoldItalic.ttf?url'

import ArimoVariable from '/fonts/Arimo-wght.ttf?url'
import ArimoItalicVariable from '/fonts/Arimo-Italic-wght.ttf?url'

export type FontVariant = 'regular' | 'bold' | 'italic' | 'boldItalic'

export type FontFiles = Readonly<Record<FontVariant, string>>

export type FontFamily = Readonly<{
  wordName: string
  substituteName: string
  files: FontFiles
}>

export const FONT_FAMILIES: readonly FontFamily[] = [
  {
    wordName: 'Calibri',
    substituteName: 'Carlito',
    files: {
      regular: CarlitoRegular,
      bold: CarlitoBold,
      italic: CarlitoItalic,
      boldItalic: CarlitoBoldItalic,
    },
  },
  {
    wordName: 'Cambria',
    substituteName: 'Caladea',
    files: {
      regular: CaladeaRegular,
      bold: CaladeaBold,
      italic: CaladeaItalic,
      boldItalic: CaladeaBoldItalic,
    },
  },
  {
    wordName: 'Courier New',
    substituteName: 'Courier Prime',
    files: {
      regular: CourierPrimeRegular,
      bold: CourierPrimeBold,
      italic: CourierPrimeItalic,
      boldItalic: CourierPrimeBoldItalic,
    },
  },
  {
    wordName: 'Times New Roman',
    substituteName: 'Tinos',
    files: {
      regular: TinosRegular,
      bold: TinosBold,
      italic: TinosItalic,
      boldItalic: TinosBoldItalic,
    },
  },
  {
    wordName: 'Arial',
    substituteName: 'Arimo',
    files: {
      regular: ArimoVariable,
      bold: ArimoVariable,
      italic: ArimoItalicVariable,
      boldItalic: ArimoItalicVariable,
    },
  },
] as const

function normalizeFamilyName(value: string): string {
  return value.trim().replace(/['"]/g, '').replace(/\s+/g, ' ').toLowerCase()
}

const FAMILY_BY_NAME = new Map<string, FontFamily>()

for (const family of FONT_FAMILIES) {
  FAMILY_BY_NAME.set(normalizeFamilyName(family.wordName), family)
  FAMILY_BY_NAME.set(normalizeFamilyName(family.substituteName), family)
}

const FAMILY_ALIASES = new Map<string, FontFamily>([
  ['calibri light', FONT_FAMILIES[0]],
  ['calibri body', FONT_FAMILIES[0]],
  ['cambria body', FONT_FAMILIES[1]],
  ['cambria heading', FONT_FAMILIES[1]],
  ['couriernew', FONT_FAMILIES[2]],
  ['times', FONT_FAMILIES[3]],
  ['times roman', FONT_FAMILIES[3]],
  ['timesnewroman', FONT_FAMILIES[3]],
  ['times new roman psmt', FONT_FAMILIES[3]],
  ['arialmt', FONT_FAMILIES[4]],
  ['arial narrow', FONT_FAMILIES[4]],
  // DXP-12: "Aptos" has been Word's default body/UI font since late 2023
  // (replacing Calibri), with "Aptos Display" as the matching default
  // heading font. No OFL/Apache-licensed metric-compatible Aptos clone is
  // available to bundle (checked against Google Fonts and the other
  // common open font catalogs Atlas already draws its substitutes from —
  // none exists as of this pass, since Aptos is a proprietary Microsoft/
  // Aptos Font Foundry design with no open clone project yet). Mapping the
  // sans-serif Aptos variants to Carlito (already bundled as Calibri's
  // substitute) is a pragmatic approximation for glyph measurement, not a
  // true metric-compatible substitute — both are similar-weight humanist
  // sans-serifs sized for on-screen reading, so this is far closer than
  // falling through to whatever sans-serif the OS defaults to (which would
  // have arbitrarily different metrics and defeat the point of this
  // substitution architecture entirely). The serif and monospace Aptos
  // variants map to the existing bundled family in the same category
  // instead, for the same reason.
  ['aptos', FONT_FAMILIES[0]],
  ['aptos display', FONT_FAMILIES[0]],
  ['aptos display narrow', FONT_FAMILIES[0]],
  ['aptos narrow', FONT_FAMILIES[0]],
  ['aptos serif', FONT_FAMILIES[3]],
  ['aptos mono', FONT_FAMILIES[2]],
])

/**
 * Word font family names Atlas has seen with no bundled substitute
 * available (DXP-12) — queryable rather than logged to the console, since
 * a packaged Electron app's console isn't visible to users and this
 * project's style forbids console output in production code. UI code can
 * use this to surface an "unsupported font" notice however is appropriate;
 * `resolveFontFamily` itself only records, it never displays anything.
 */
const unmappedFontNames = new Set<string>()

export function getUnmappedFontNames(): ReadonlySet<string> {
  return unmappedFontNames
}

/** Test-only hook: clear the unmapped-font registry between test cases. */
export function __resetUnmappedFontNamesForTests(): void {
  unmappedFontNames.clear()
}

export function resolveFontFamily(wordName: string): FontFamily | null {
  const normalized = normalizeFamilyName(wordName)
  const resolved = FAMILY_BY_NAME.get(normalized) ?? FAMILY_ALIASES.get(normalized) ?? null

  if (resolved === null && normalized !== '') {
    unmappedFontNames.add(wordName)
  }

  return resolved
}
