/**
 * T6 / DAT-14 — extension -> shiki language map, derived from the single
 * canonical extension manifest (`src/formats/extensionManifest.ts`) instead
 * of a hand-maintained literal that only covered 39 of the 61 extensions
 * `detect.ts` actually routes to `CodeViewer` (`extensionManifest.test.ts`
 * asserts every code-routed extension is covered and resolves to a real
 * shiki bundled language).
 */
import { CODE_EXTENSION_TO_SHIKI_LANG } from '../formats/extensionManifest'

export const extToLang: Readonly<Record<string, string>> = Object.freeze(
  Object.fromEntries(
    Object.entries(CODE_EXTENSION_TO_SHIKI_LANG).map(([ext, lang]) => [`.${ext}`, lang]),
  ),
)

export function getLangForExt(ext: string): string | undefined {
  const normalized = ext.startsWith('.') ? ext.toLowerCase() : `.${ext.toLowerCase()}`
  return extToLang[normalized]
}
