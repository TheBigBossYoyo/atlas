/**
 * Atlas — DOCX theme parser (Wave A.5)
 *
 * Parses `word/theme/theme1.xml` (or any `a:theme` XML part).
 *
 * NOTE: `Theme` is NOT exported from `src/docx/model` (Wave A.2 did not
 * define it).  We define it here as a named export so the parser barrel
 * (`index.ts`) can re-export it.  If the model ever adds `Theme`, remove the
 * local definition and import from `../model/document` instead.
 *
 * Colour values: each slot may be `<a:srgbClr val="HEX">` (direct hex) or
 * `<a:sysClr lastClr="HEX">` (system colour with last-known fallback).
 * We always store the resolved hex string.
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import { hexColor } from '../model/styles'
import type { HexColor } from '../model/styles'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export interface FontSet {
  readonly latin: string
  readonly ea?: string
  readonly cs?: string
}

export interface FontScheme {
  readonly major: FontSet
  readonly minor: FontSet
}

/**
 * Parsed representation of an OOXML `a:theme` part.
 *
 * Defined locally because Wave A.2 (model) does not export `Theme`.
 * TODO (Wave A.6): move to src/docx/model/document.ts and import from there.
 */
export interface Theme {
  readonly fontScheme: FontScheme
  readonly colorScheme: ReadonlyMap<string, HexColor>
}

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Internal XML shapes
// ---------------------------------------------------------------------------

interface RawFontRef {
  '@_typeface'?: string
  [key: string]: unknown
}

interface RawMajorMinor {
  'a:latin'?: RawFontRef
  'a:ea'?: RawFontRef
  'a:cs'?: RawFontRef
  [key: string]: unknown
}

interface RawFontScheme {
  'a:majorFont'?: RawMajorMinor
  'a:minorFont'?: RawMajorMinor
  [key: string]: unknown
}

interface RawColorSlot {
  'a:srgbClr'?: { '@_val'?: string }
  'a:sysClr'?: { '@_lastClr'?: string }
  [key: string]: unknown
}

type RawColorScheme = Record<string, RawColorSlot | unknown>

interface RawThemeElements {
  'a:fontScheme'?: RawFontScheme
  'a:clrScheme'?: RawColorScheme
  [key: string]: unknown
}

interface RawTheme {
  'a:theme'?: {
    'a:themeElements'?: RawThemeElements
    [key: string]: unknown
  }
  [key: string]: unknown
}

// ---------------------------------------------------------------------------
// Known colour slot names per OOXML spec
// ---------------------------------------------------------------------------

const COLOR_SLOT_NAMES: ReadonlyArray<string> = [
  'a:dk1', 'a:lt1', 'a:dk2', 'a:lt2',
  'a:accent1', 'a:accent2', 'a:accent3', 'a:accent4', 'a:accent5', 'a:accent6',
  'a:hlink', 'a:folHlink',
]

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function extractTypeface(ref: RawFontRef | undefined): string {
  return ref?.['@_typeface'] ?? ''
}

function parseFontSet(raw: RawMajorMinor | undefined): FontSet {
  const latin = extractTypeface(raw?.['a:latin'])
  const ea = raw?.['a:ea'] !== undefined ? extractTypeface(raw['a:ea']) : undefined
  const cs = raw?.['a:cs'] !== undefined ? extractTypeface(raw['a:cs']) : undefined
  return {
    latin,
    ...(ea !== undefined ? { ea } : {}),
    ...(cs !== undefined ? { cs } : {}),
  }
}

function resolveColorSlot(slot: RawColorSlot): HexColor | undefined {
  const srgb = slot['a:srgbClr']
  if (srgb !== undefined && typeof srgb === 'object' && srgb !== null) {
    const val = (srgb as { '@_val'?: string })['@_val']
    if (typeof val === 'string' && val.length > 0) return hexColor(val)
  }
  const sys = slot['a:sysClr']
  if (sys !== undefined && typeof sys === 'object' && sys !== null) {
    const last = (sys as { '@_lastClr'?: string })['@_lastClr']
    if (typeof last === 'string' && last.length > 0) return hexColor(last)
  }
  return undefined
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse the XML content of a DOCX theme part (`word/theme/theme1.xml`).
 *
 * @param xml - UTF-8 text of the theme XML file.
 * @returns   Parsed `Theme` with fontScheme and colorScheme.
 * @throws    DocxParseError on malformed XML or missing `<a:theme>` root.
 */
export function parseTheme(xml: string): Theme {
  assertXmlPartSizeWithinLimit(xml, 'word/theme/theme1.xml')
  let parsed: RawTheme
  try {
    parsed = xmlParser.parse(xml) as RawTheme
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse theme XML: ${msg}`)
  }

  const themeEl = parsed?.['a:theme']
  if (themeEl === undefined || themeEl === null) {
    throw new DocxParseError('Theme XML is missing required <a:theme> root element')
  }

  const elements = themeEl['a:themeElements']

  // ── Font scheme ──────────────────────────────────────────────────────────
  const rawFontScheme = elements?.['a:fontScheme']
  const fontScheme: FontScheme = {
    major: parseFontSet(rawFontScheme?.['a:majorFont']),
    minor: parseFontSet(rawFontScheme?.['a:minorFont']),
  }

  // ── Colour scheme ─────────────────────────────────────────────────────────
  const rawClr = elements?.['a:clrScheme'] as RawColorScheme | undefined
  const colorMap = new Map<string, HexColor>()

  if (rawClr !== undefined && rawClr !== null) {
    for (const slotName of COLOR_SLOT_NAMES) {
      const slot = rawClr[slotName]
      if (slot !== undefined && slot !== null && typeof slot === 'object') {
        // Strip the "a:" namespace prefix to produce keys like "dk1", "lt1"…
        const key = slotName.slice(2)
        const color = resolveColorSlot(slot as RawColorSlot)
        if (color !== undefined) colorMap.set(key, color)
      }
    }
  }

  return { fontScheme, colorScheme: colorMap }
}

/**
 * Resolve a theme font reference (e.g. "minorHAnsi", "majorAscii") to the
 * concrete typeface name declared in the theme's font scheme.
 *
 * Returns `undefined` when the theme is missing or the reference doesn't
 * resolve to a non-empty typeface — callers should fall back to their own
 * default family (typically Calibri for minor, Cambria for major).
 */
export function resolveThemeFont(
  theme: Theme | undefined,
  reference: string | undefined,
): string | undefined {
  if (theme === undefined || reference === undefined) return undefined

  const scheme = reference.startsWith('major') ? theme.fontScheme.major : theme.fontScheme.minor

  if (reference.endsWith('EastAsia')) {
    const value = scheme.ea ?? scheme.latin
    return value !== undefined && value.length > 0 ? value : undefined
  }
  if (reference.endsWith('Bidi')) {
    const value = scheme.cs ?? scheme.latin
    return value !== undefined && value.length > 0 ? value : undefined
  }

  // majorAscii / minorAscii / majorHAnsi / minorHAnsi → latin
  return scheme.latin.length > 0 ? scheme.latin : undefined
}
