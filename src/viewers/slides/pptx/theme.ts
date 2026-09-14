/**
 * OOXML theme color resolution (S3 — `SlideData` styled runs + theme colors).
 *
 * Resolves `a:schemeClr`/`a:srgbClr`/`a:sysClr`, composed with `lumMod`,
 * `lumOff`, `tint`, and `shade` modifiers, against a slide master's theme
 * (`ppt/theme/themeN.xml`) and its `p:clrMap` indirection.
 */

import { getDirectChildren, getFirstByLocalName } from '../shared/xmlUtils'

export type ThemeColors = ReadonlyMap<string, string>
export type ColorMap = ReadonlyMap<string, string>

const SCHEME_SLOT_NAMES = [
  'dk1', 'lt1', 'dk2', 'lt2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
] as const

const DEFAULT_COLOR = '#000000'

/** Parses `<a:clrScheme>` from a theme document into slot -> `#rrggbb`. */
export function parseThemeColors(themeDocument: XMLDocument | null): ThemeColors {
  const colors = new Map<string, string>()
  const scheme = themeDocument ? getFirstByLocalName(themeDocument, 'clrScheme') : null

  if (!scheme) {
    return colors
  }

  for (const slotName of SCHEME_SLOT_NAMES) {
    const slot = getDirectChildren(scheme).find(child => child.localName === slotName)
    if (!slot) {
      continue
    }

    const srgb = getFirstByLocalName(slot, 'srgbClr')
    const sys = getFirstByLocalName(slot, 'sysClr')
    const hex = srgb?.getAttribute('val') ?? sys?.getAttribute('lastClr') ?? sys?.getAttribute('val')

    if (hex) {
      colors.set(slotName, normalizeHex(hex))
    }
  }

  return colors
}

/** Parses a slide master's `<p:clrMap>` attributes (bg1="lt1" tx1="dk1" ...). */
export function parseClrMap(masterDocument: XMLDocument | null): ColorMap {
  const map = new Map<string, string>()
  const clrMap = masterDocument ? getFirstByLocalName(masterDocument, 'clrMap') : null

  if (!clrMap) {
    return map
  }

  for (const attribute of Array.from(clrMap.attributes)) {
    if (attribute.name.startsWith('xmlns')) {
      continue
    }

    map.set(attribute.localName, attribute.value)
  }

  return map
}

function normalizeHex(hex: string): string {
  const cleaned = hex.replace('#', '').toUpperCase()
  return `#${cleaned}`
}

function clamp01(value: number): number {
  return Math.min(1, Math.max(0, value))
}

function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const cleaned = hex.replace('#', '')
  const value = Number.parseInt(cleaned.length === 3
    ? cleaned.split('').map(c => c + c).join('')
    : cleaned, 16)

  return {
    r: (value >> 16) & 0xff,
    g: (value >> 8) & 0xff,
    b: value & 0xff,
  }
}

function rgbToHex(r: number, g: number, b: number): string {
  const channel = (value: number) => Math.round(clamp01(value / 255) * 255).toString(16).padStart(2, '0')
  return `#${channel(r)}${channel(g)}${channel(b)}`.toUpperCase()
}

function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  const rN = r / 255
  const gN = g / 255
  const bN = b / 255
  const max = Math.max(rN, gN, bN)
  const min = Math.min(rN, gN, bN)
  const l = (max + min) / 2

  if (max === min) {
    return { h: 0, s: 0, l }
  }

  const delta = max - min
  const s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min)

  let h = 0
  if (max === rN) {
    h = ((gN - bN) / delta) % 6
  } else if (max === gN) {
    h = (bN - rN) / delta + 2
  } else {
    h = (rN - gN) / delta + 4
  }

  h *= 60
  if (h < 0) {
    h += 360
  }

  return { h, s, l }
}

function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  if (s === 0) {
    const v = Math.round(l * 255)
    return { r: v, g: v, b: v }
  }

  const c = (1 - Math.abs(2 * l - 1)) * s
  const hPrime = h / 60
  const x = c * (1 - Math.abs((hPrime % 2) - 1))
  const m = l - c / 2

  let rgb: [number, number, number]
  if (hPrime < 1) rgb = [c, x, 0]
  else if (hPrime < 2) rgb = [x, c, 0]
  else if (hPrime < 3) rgb = [0, c, x]
  else if (hPrime < 4) rgb = [0, x, c]
  else if (hPrime < 5) rgb = [x, 0, c]
  else rgb = [c, 0, x]

  return {
    r: Math.round((rgb[0] + m) * 255),
    g: Math.round((rgb[1] + m) * 255),
    b: Math.round((rgb[2] + m) * 255),
  }
}

export type ColorModifiers = {
  readonly lumMod?: number
  readonly lumOff?: number
  readonly tint?: number
  readonly shade?: number
  readonly alpha?: number
}

/** Applies OOXML `lumMod`/`lumOff`/`tint`/`shade` (each a 0..1 fraction) to a `#rrggbb` color. */
export function applyColorModifiers(hex: string, modifiers: ColorModifiers): string {
  const { r, g, b } = hexToRgb(hex)
  const hsl = rgbToHsl(r, g, b)
  const { h, s } = hsl
  let { l } = hsl

  if (modifiers.lumMod !== undefined) {
    l *= modifiers.lumMod
  }
  if (modifiers.lumOff !== undefined) {
    l += modifiers.lumOff
  }
  if (modifiers.shade !== undefined) {
    l *= modifiers.shade
  }
  if (modifiers.tint !== undefined) {
    l = l * modifiers.tint + (1 - modifiers.tint)
  }

  l = clamp01(l)
  const rgb = hslToRgb(h, s, l)
  return rgbToHex(rgb.r, rgb.g, rgb.b)
}

/** Reads `val`/1000ths-of-percent attributes shared by lumMod/lumOff/tint/shade/alpha. */
function readPercentAttribute(element: Element | null): number | undefined {
  const raw = element?.getAttribute('val')
  if (raw === null || raw === undefined) {
    return undefined
  }

  const parsed = Number(raw)
  return Number.isFinite(parsed) ? parsed / 100000 : undefined
}

/** Resolves a `schemeClr`/`srgbClr`/`sysClr` element (with its modifier children) to a CSS color. */
export function resolveColorElement(
  colorElement: Element,
  themeColors: ThemeColors,
  clrMap: ColorMap,
): string {
  let base = DEFAULT_COLOR

  if (colorElement.localName === 'srgbClr') {
    base = normalizeHex(colorElement.getAttribute('val') ?? '000000')
  } else if (colorElement.localName === 'sysClr') {
    base = normalizeHex(colorElement.getAttribute('lastClr') ?? colorElement.getAttribute('val') ?? '000000')
  } else if (colorElement.localName === 'schemeClr') {
    const rawVal = colorElement.getAttribute('val') ?? ''
    const resolvedSlot = clrMap.get(rawVal) ?? rawVal
    base = themeColors.get(resolvedSlot) ?? DEFAULT_COLOR
  }

  const modifiers: ColorModifiers = {
    lumMod: readPercentAttribute(getFirstByLocalName(colorElement, 'lumMod')),
    lumOff: readPercentAttribute(getFirstByLocalName(colorElement, 'lumOff')),
    tint: readPercentAttribute(getFirstByLocalName(colorElement, 'tint')),
    shade: readPercentAttribute(getFirstByLocalName(colorElement, 'shade')),
  }

  if (Object.values(modifiers).some(value => value !== undefined)) {
    return applyColorModifiers(base, modifiers)
  }

  return base
}

/** Finds a color element (`srgbClr`/`schemeClr`/`sysClr`) among `parent`'s direct children. */
export function findColorElement(parent: Element): Element | null {
  return getDirectChildren(parent).find(
    child => child.localName === 'srgbClr' || child.localName === 'schemeClr' || child.localName === 'sysClr',
  ) ?? null
}

/**
 * Resolves the color carried by a `solidFill`-shaped container element
 * (the `<a:solidFill>` itself, or an ancestor like `<a:ln>` that holds one).
 */
export function resolveSolidFillColor(
  container: Element | null,
  themeColors: ThemeColors,
  clrMap: ColorMap,
): string | undefined {
  if (!container) {
    return undefined
  }

  const solidFill = container.localName === 'solidFill' ? container : getFirstByLocalName(container, 'solidFill')
  if (!solidFill) {
    return undefined
  }

  const colorElement = findColorElement(solidFill)
  return colorElement ? resolveColorElement(colorElement, themeColors, clrMap) : undefined
}
