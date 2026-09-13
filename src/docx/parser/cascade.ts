/**
 * Atlas — DOCX style cascade resolver (Wave A.4)
 *
 * Resolves effective run/paragraph properties across doc defaults, `basedOn`
 * style chains, linked paragraph/character styles, and direct formatting.
 */

import type {
  Border,
  BorderSet,
  FontSet,
  FrameProps,
  Indent,
  LanguageSet,
  ParaProps,
  RunProps,
  SectionProps,
  Shading,
  Spacing,
  Style,
  TabSet,
  Underline,
} from '../model'
import { DocxParseError } from './unzip'

export function resolveRunProps(
  direct: RunProps | undefined,
  styleId: string | undefined,
  styles: ReadonlyMap<string, Style>,
  docDefaults: { rPr?: RunProps },
): RunProps {
  let resolved = mergeRunProps(undefined, docDefaults.rPr) ?? {}

  const styleChain = resolveStyleChain(styleId, styles)

  for (const style of styleChain) {
    resolved = mergeRunProps(resolved, style.run) ?? resolved
  }

  const terminalStyle = styleChain.length > 0 ? styleChain[styleChain.length - 1] : undefined

  if (terminalStyle?.type === 'paragraph' && terminalStyle.linked !== undefined) {
    const linkedStyle = styles.get(terminalStyle.linked)
    if (linkedStyle?.type === 'character') {
      for (const style of resolveStyleChain(linkedStyle.id, styles)) {
        resolved = mergeRunProps(resolved, style.run) ?? resolved
      }
    }
  }

  return mergeRunProps(resolved, direct) ?? {}
}

export function resolveParaProps(
  direct: ParaProps | undefined,
  styleId: string | undefined,
  styles: ReadonlyMap<string, Style>,
  docDefaults: { pPr?: ParaProps },
): ParaProps {
  let resolved = mergeParaProps(undefined, docDefaults.pPr) ?? {}

  for (const style of resolveStyleChain(styleId, styles)) {
    resolved = mergeParaProps(resolved, style.paragraph) ?? resolved
  }

  return mergeParaProps(resolved, direct) ?? {}
}

function resolveStyleChain(
  styleId: string | undefined,
  styles: ReadonlyMap<string, Style>,
  depth: number = 0,
): ReadonlyArray<Style> {
  if (styleId === undefined || styleId === '') return []

  if (depth > 20) {
    throw new DocxParseError(
      `Style basedOn chain exceeded 20 levels while resolving "${styleId}"`,
    )
  }

  const style = styles.get(styleId)
  if (style === undefined) return []

  const basedOn =
    style.basedOn !== undefined
      ? resolveStyleChain(style.basedOn, styles, depth + 1)
      : []

  return [...basedOn, style]
}

function mergeRunProps(
  base: RunProps | undefined,
  override: RunProps | undefined,
): RunProps | undefined {
  if (base === undefined && override === undefined) return undefined

  const merged = { ...(base ?? {}), ...(override ?? {}) }

  const underline = mergeUnderline(base?.underline, override?.underline)
  const shd = mergeShading(base?.shd, override?.shd)
  const rFonts = mergeFontSet(base?.rFonts, override?.rFonts)
  const lang = mergeLanguageSet(base?.lang, override?.lang)

  if (underline !== undefined) merged.underline = underline
  if (shd !== undefined) merged.shd = shd
  if (rFonts !== undefined) merged.rFonts = rFonts
  if (lang !== undefined) merged.lang = lang

  return merged
}

function mergeParaProps(
  base: ParaProps | undefined,
  override: ParaProps | undefined,
): ParaProps | undefined {
  if (base === undefined && override === undefined) return undefined

  const merged = { ...(base ?? {}), ...(override ?? {}) }

  const numPr = mergeNumPr(base?.numPr, override?.numPr)
  const spacing = mergeSpacing(base?.spacing, override?.spacing)
  const ind = mergeIndent(base?.ind, override?.ind)
  const tabs = mergeTabSet(base?.tabs, override?.tabs)
  const pBdr = mergeBorderSet(base?.pBdr, override?.pBdr)
  const shd = mergeShading(base?.shd, override?.shd)
  const framePr = mergeFrameProps(base?.framePr, override?.framePr)
  const sectPr = mergeSectionProps(base?.sectPr, override?.sectPr)

  if (numPr !== undefined) merged.numPr = numPr
  if (spacing !== undefined) merged.spacing = spacing
  if (ind !== undefined) merged.ind = ind
  if (tabs !== undefined) merged.tabs = tabs
  if (pBdr !== undefined) merged.pBdr = pBdr
  if (shd !== undefined) merged.shd = shd
  if (framePr !== undefined) merged.framePr = framePr
  if (sectPr !== undefined) merged.sectPr = sectPr

  return merged
}

function mergeUnderline(
  base: Underline | undefined,
  override: Underline | undefined,
): Underline | undefined {
  if (base === undefined && override === undefined) return undefined

  const style = override?.style ?? base?.style
  if (style === undefined) return undefined

  return {
    style,
    ...(preferDefined('color', base?.color, override?.color)),
  }
}

function mergeShading(
  base: Shading | undefined,
  override: Shading | undefined,
): Shading | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeFontSet(
  base: FontSet | undefined,
  override: FontSet | undefined,
): FontSet | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeLanguageSet(
  base: LanguageSet | undefined,
  override: LanguageSet | undefined,
): LanguageSet | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeNumPr(
  base: ParaProps['numPr'] | undefined,
  override: ParaProps['numPr'] | undefined,
): ParaProps['numPr'] | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeSpacing(
  base: Spacing | undefined,
  override: Spacing | undefined,
): Spacing | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeIndent(
  base: Indent | undefined,
  override: Indent | undefined,
): Indent | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeTabSet(
  base: TabSet | undefined,
  override: TabSet | undefined,
): TabSet | undefined {
  if (base === undefined && override === undefined) return undefined

  return {
    items: override !== undefined ? override.items : base?.items ?? [],
    ...(preferDefined('defaultTabStop', base?.defaultTabStop, override?.defaultTabStop)),
  }
}

function mergeBorderSet(
  base: BorderSet | undefined,
  override: BorderSet | undefined,
): BorderSet | undefined {
  if (base === undefined && override === undefined) return undefined

  const merged: BorderSet = {
    ...(withDefined('top', mergeBorder(base?.top, override?.top))),
    ...(withDefined('left', mergeBorder(base?.left, override?.left))),
    ...(withDefined('bottom', mergeBorder(base?.bottom, override?.bottom))),
    ...(withDefined('right', mergeBorder(base?.right, override?.right))),
    ...(withDefined('start', mergeBorder(base?.start, override?.start))),
    ...(withDefined('end', mergeBorder(base?.end, override?.end))),
    ...(withDefined('between', mergeBorder(base?.between, override?.between))),
    ...(withDefined('bar', mergeBorder(base?.bar, override?.bar))),
    ...(withDefined('insideH', mergeBorder(base?.insideH, override?.insideH))),
    ...(withDefined('insideV', mergeBorder(base?.insideV, override?.insideV))),
  }

  return hasKeys(merged) ? merged : undefined
}

function mergeBorder(
  base: Border | undefined,
  override: Border | undefined,
): Border | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeFrameProps(
  base: FrameProps | undefined,
  override: FrameProps | undefined,
): FrameProps | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeSectionProps(
  base: SectionProps | undefined,
  override: SectionProps | undefined,
): SectionProps | undefined {
  if (base === undefined && override === undefined) return undefined

  const merged = { ...(base ?? {}), ...(override ?? {}) }

  const pgSz = mergePageSize(base?.pgSz, override?.pgSz)
  const pgMar = mergePageMargins(base?.pgMar, override?.pgMar)
  const cols = mergeSectionColumns(base?.cols, override?.cols)
  const pgNumType = mergePageNumberType(base?.pgNumType, override?.pgNumType)
  const lnNumType = mergeLineNumberType(base?.lnNumType, override?.lnNumType)

  if (pgSz !== undefined) merged.pgSz = pgSz
  if (pgMar !== undefined) merged.pgMar = pgMar
  if (cols !== undefined) merged.cols = cols
  if (pgNumType !== undefined) merged.pgNumType = pgNumType
  if (lnNumType !== undefined) merged.lnNumType = lnNumType

  return merged
}

function mergePageSize(
  base: SectionProps['pgSz'] | undefined,
  override: SectionProps['pgSz'] | undefined,
): SectionProps['pgSz'] | undefined {
  if (base === undefined && override === undefined) return undefined

  const w = override?.w ?? base?.w
  const h = override?.h ?? base?.h

  if (w === undefined || h === undefined) return undefined

  return {
    w,
    h,
    ...(preferDefined('orient', base?.orient, override?.orient)),
  }
}

function mergePageMargins(
  base: SectionProps['pgMar'] | undefined,
  override: SectionProps['pgMar'] | undefined,
): SectionProps['pgMar'] | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeSectionColumns(
  base: SectionProps['cols'] | undefined,
  override: SectionProps['cols'] | undefined,
): SectionProps['cols'] | undefined {
  if (base === undefined && override === undefined) return undefined

  return {
    ...(preferDefined('num', base?.num, override?.num)),
    ...(preferDefined('space', base?.space, override?.space)),
    ...(preferDefined('sep', base?.sep, override?.sep)),
    ...(preferDefined('equalWidth', base?.equalWidth, override?.equalWidth)),
    col: override !== undefined ? override.col : base?.col ?? [],
  }
}

function mergePageNumberType(
  base: SectionProps['pgNumType'] | undefined,
  override: SectionProps['pgNumType'] | undefined,
): SectionProps['pgNumType'] | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeLineNumberType(
  base: SectionProps['lnNumType'] | undefined,
  override: SectionProps['lnNumType'] | undefined,
): SectionProps['lnNumType'] | undefined {
  if (base === undefined && override === undefined) return undefined
  return { ...(base ?? {}), ...(override ?? {}) }
}

function preferDefined<K extends string, V>(
  key: K,
  base: V | undefined,
  override: V | undefined,
){
  if (override !== undefined) {
    return { [key]: override }
  }
  if (base !== undefined) {
    return { [key]: base }
  }
  return {}
}

function withDefined<K extends string, V>(key: K, value: V | undefined) {
  if (value === undefined) return {}
  return { [key]: value }
}

function hasKeys(value: object): boolean {
  return Object.keys(value).length > 0
}
