/**
 * The bundled fonts PDF text editing falls back to (PDF-EDIT-2), loaded on
 * demand in the app: the same files the DOCX renderer already ships in
 * `public/fonts`, plus fontkit to embed subsets of them.
 */

import ArimoVariable from '/fonts/Arimo-wght.ttf?url'
import ArimoItalicVariable from '/fonts/Arimo-Italic-wght.ttf?url'
import CarlitoBold from '/fonts/Carlito-Bold.ttf?url'
import CarlitoBoldItalic from '/fonts/Carlito-BoldItalic.ttf?url'
import CourierPrimeBold from '/fonts/CourierPrime-Bold.ttf?url'
import CourierPrimeBoldItalic from '/fonts/CourierPrime-BoldItalic.ttf?url'
import CourierPrimeItalic from '/fonts/CourierPrime-Italic.ttf?url'
import CourierPrimeRegular from '/fonts/CourierPrime-Regular.ttf?url'
import TinosBold from '/fonts/Tinos-Bold.ttf?url'
import TinosBoldItalic from '/fonts/Tinos-BoldItalic.ttf?url'
import TinosItalic from '/fonts/Tinos-Italic.ttf?url'
import TinosRegular from '/fonts/Tinos-Regular.ttf?url'

import type { FallbackFontSource } from './content/rewriteText'
import type { FontStyle } from './types'
import { fallbackFontFile } from './content/textEdits'

const URLS: Record<string, string> = {
  'Arimo-wght.ttf': ArimoVariable,
  'Arimo-Italic-wght.ttf': ArimoItalicVariable,
  'Carlito-Bold.ttf': CarlitoBold,
  'Carlito-BoldItalic.ttf': CarlitoBoldItalic,
  'CourierPrime-Regular.ttf': CourierPrimeRegular,
  'CourierPrime-Bold.ttf': CourierPrimeBold,
  'CourierPrime-Italic.ttf': CourierPrimeItalic,
  'CourierPrime-BoldItalic.ttf': CourierPrimeBoldItalic,
  'Tinos-Regular.ttf': TinosRegular,
  'Tinos-Bold.ttf': TinosBold,
  'Tinos-Italic.ttf': TinosItalic,
  'Tinos-BoldItalic.ttf': TinosBoldItalic,
}

const bytesCache = new Map<string, Promise<Uint8Array>>()

function fetchFont(file: string): Promise<Uint8Array> {
  let pending = bytesCache.get(file)
  if (!pending) {
    pending = fetch(URLS[file])
      .then((response) => {
        if (!response.ok) throw new Error(`Could not load the bundled font ${file}.`)
        return response.arrayBuffer()
      })
      .then((buffer) => new Uint8Array(buffer))
    pending.catch(() => bytesCache.delete(file))
    bytesCache.set(file, pending)
  }
  return pending
}

let sourcePromise: Promise<FallbackFontSource> | null = null

export function loadFallbackFonts(): Promise<FallbackFontSource> {
  sourcePromise ??= import('@pdf-lib/fontkit')
    .then((module) => ({
      fontkit: module.default,
      load: (style: FontStyle) => fetchFont(fallbackFontFile(style)),
    }))
    .catch((error: unknown) => {
      sourcePromise = null
      throw error
    })
  return sourcePromise
}
