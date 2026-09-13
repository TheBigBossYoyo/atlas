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
])

export function resolveFontFamily(wordName: string): FontFamily | null {
  const normalized = normalizeFamilyName(wordName)
  return FAMILY_BY_NAME.get(normalized) ?? FAMILY_ALIASES.get(normalized) ?? null
}
