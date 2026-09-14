/**
 * Tests for src/docx/fonts/families.ts (D22 / DXP-12)
 */
import { afterEach, describe, expect, it } from 'vitest'

import {
  FONT_FAMILIES,
  __resetUnmappedFontNamesForTests,
  getUnmappedFontNames,
  resolveFontFamily,
} from '../families'

describe('resolveFontFamily', () => {
  afterEach(() => {
    __resetUnmappedFontNamesForTests()
  })

  describe('Aptos mapping (DXP-12)', () => {
    it('maps the plain "Aptos" body font to the bundled Calibri substitute', () => {
      expect(resolveFontFamily('Aptos')).toBe(FONT_FAMILIES[0])
    })

    it('maps "Aptos Display" (the default heading font) to the same sans-serif substitute', () => {
      expect(resolveFontFamily('Aptos Display')).toBe(FONT_FAMILIES[0])
    })

    it('maps "Aptos Display Narrow" and "Aptos Narrow"', () => {
      expect(resolveFontFamily('Aptos Display Narrow')).toBe(FONT_FAMILIES[0])
      expect(resolveFontFamily('Aptos Narrow')).toBe(FONT_FAMILIES[0])
    })

    it('maps "Aptos Serif" to the bundled serif substitute, not the sans one', () => {
      expect(resolveFontFamily('Aptos Serif')).toBe(FONT_FAMILIES[3])
    })

    it('maps "Aptos Mono" to the bundled monospace substitute', () => {
      expect(resolveFontFamily('Aptos Mono')).toBe(FONT_FAMILIES[2])
    })

    it('is case- and whitespace-insensitive, matching the existing alias behaviour', () => {
      expect(resolveFontFamily('  aptos DISPLAY  ')).toBe(FONT_FAMILIES[0])
    })

    it('does not record Aptos as an unmapped font', () => {
      resolveFontFamily('Aptos')
      expect(getUnmappedFontNames().has('Aptos')).toBe(false)
    })
  })

  describe('unmapped-font tracking (DXP-12)', () => {
    it('records a font name with no bundled substitute', () => {
      resolveFontFamily('Comic Sans MS')
      expect(getUnmappedFontNames().has('Comic Sans MS')).toBe(true)
    })

    it('returns null for an unmapped font, same as before', () => {
      expect(resolveFontFamily('Comic Sans MS')).toBeNull()
    })

    it('does not record a successfully-resolved font', () => {
      resolveFontFamily('Calibri')
      expect(getUnmappedFontNames().has('Calibri')).toBe(false)
    })

    it('does not record an empty font name', () => {
      resolveFontFamily('')
      expect(getUnmappedFontNames().size).toBe(0)
    })

    it('accumulates distinct unmapped font names across calls', () => {
      resolveFontFamily('Comic Sans MS')
      resolveFontFamily('Papyrus')
      expect(getUnmappedFontNames()).toEqual(new Set(['Comic Sans MS', 'Papyrus']))
    })
  })
})
