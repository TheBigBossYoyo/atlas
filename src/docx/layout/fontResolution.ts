/**
 * Atlas — Shared font-family resolution (3.0.3 fix)
 *
 * Single source of truth for "what concrete font family does this run want?".
 * Used by itemize.ts (word measurement), breakLines.ts (line metrics) AND
 * render/style.ts (DOM rendering) so measurement and painting never diverge.
 *
 * Resolution order (matches Word):
 *   1. Theme reference on rFonts (asciiTheme/hAnsiTheme/csTheme/eastAsiaTheme)
 *      → look up in document theme.fontScheme via resolveThemeFont().
 *   2. Literal rFonts.ascii / hAnsi / cs / eastAsia.
 *   3. DEFAULT_FONT_FAMILY ('Calibri').
 *
 * Script / direction selects which slot is consulted first:
 *   - EastAsian script  → eastAsia* → ascii* → hAnsi* → cs*
 *   - rtl run           → cs*       → ascii* → hAnsi* → eastAsia*
 *   - default (latin)   → ascii*    → hAnsi* → eastAsia* → cs*
 */

import type { FontSet, FontTheme, RunProps } from '../model/styles'
import { resolveThemeFont, type Theme } from '../parser/theme'

export const DEFAULT_FONT_FAMILY = 'Calibri'

export type FontSlot = 'ascii' | 'hAnsi' | 'cs' | 'eastAsia'

const LATIN_ORDER: ReadonlyArray<FontSlot> = ['ascii', 'hAnsi', 'eastAsia', 'cs']
const EAST_ASIAN_ORDER: ReadonlyArray<FontSlot> = ['eastAsia', 'ascii', 'hAnsi', 'cs']
const RTL_ORDER: ReadonlyArray<FontSlot> = ['cs', 'ascii', 'hAnsi', 'eastAsia']

function themeAttr(fonts: FontSet, slot: FontSlot): FontTheme | undefined {
  switch (slot) {
    case 'ascii':
      return fonts.asciiTheme
    case 'hAnsi':
      return fonts.hAnsiTheme
    case 'cs':
      return fonts.csTheme
    case 'eastAsia':
      return fonts.eastAsiaTheme
  }
}

function literalAttr(fonts: FontSet, slot: FontSlot): string | undefined {
  switch (slot) {
    case 'ascii':
      return fonts.ascii
    case 'hAnsi':
      return fonts.hAnsi
    case 'cs':
      return fonts.cs
    case 'eastAsia':
      return fonts.eastAsia
  }
}

function resolveSlot(
  fonts: FontSet,
  slot: FontSlot,
  theme: Theme | undefined,
): string | undefined {
  const themed = resolveThemeFont(theme, themeAttr(fonts, slot))
  if (themed !== undefined && themed.length > 0) return themed
  const literal = literalAttr(fonts, slot)
  if (literal !== undefined && literal.length > 0) return literal
  return undefined
}

export type ScriptHint = 'latin' | 'eastAsian' | 'rtl'

/**
 * Pick the effective concrete font family for a run.
 *
 * @param runProps  Effective run properties (post-cascade).
 * @param theme     Parsed document theme (or undefined if missing).
 * @param hint      Script hint — determines which rFonts slot is consulted first.
 */
export function effectiveFontFamily(
  runProps: Pick<RunProps, 'rFonts' | 'rtl'> | undefined,
  theme: Theme | undefined,
  hint: ScriptHint,
): string {
  const fonts = runProps?.rFonts
  if (fonts === undefined) return DEFAULT_FONT_FAMILY

  const order =
    hint === 'eastAsian'
      ? EAST_ASIAN_ORDER
      : hint === 'rtl' || runProps?.rtl === true
        ? RTL_ORDER
        : LATIN_ORDER

  for (const slot of order) {
    const resolved = resolveSlot(fonts, slot, theme)
    if (resolved !== undefined) return resolved
  }

  return DEFAULT_FONT_FAMILY
}
