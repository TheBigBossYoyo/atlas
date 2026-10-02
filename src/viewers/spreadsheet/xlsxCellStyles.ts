/**
 * SHEETFMT-1 — reading the cell formatting an `.xlsx`/`.xlsm` actually carries.
 *
 * Atlas has always rendered every workbook as unstyled text: a file whose header
 * row is bold white-on-blue with borders and a currency format looked exactly
 * like a CSV. The styles were never *lost* — the save path keeps the whole
 * original `styles.xml` and every cell's `s` index (`xlsxPassthrough.ts`) — they
 * were simply never read, because nothing in the pipeline could see them:
 *
 *   - SheetJS's community build does not expose cell styles. With
 *     `cellStyles: true` a cell comes back as `{t,v,w,z,s:{patternType:'none'}}`
 *     — the number format (`z`) is real, `s` is a stub with no font, fill or
 *     border in it. Verified against a styled fixture, not assumed.
 *   - So the only source of truth is the package XML itself: `xl/styles.xml`
 *     for the format tables, and each worksheet's own `<c s="…">` attributes
 *     for which cell uses which.
 *
 * This module reads exactly that, and nothing else. It does not write; applying
 * a format is `xlsxPassthrough.ts`'s job, which already owns the `<xf>`
 * interning machinery this module's output is designed to line up with.
 *
 * **Shape of the result.** A format is NOT stored per cell. A real workbook has
 * a few dozen distinct formats shared across a million cells, which is exactly
 * what OOXML's `s`-index indirection expresses, so this keeps that shape: a
 * `formats` lookup table plus one small integer per cell. A 100k-row sheet costs
 * one number per cell rather than one object, which matters because that is the
 * size this repo's own perf tests use.
 *
 * **Never throws.** A missing, truncated or foreign `styles.xml` returns `null`
 * and the grid renders exactly as it did before. Formatting is a visual
 * enhancement; it must never be the reason a file won't open.
 */
import JSZip from 'jszip'

import { childElements, firstChildElement, parseXmlPart } from '../../office/ooxmlDom'

/**
 * One border edge.
 *
 * OOXML defines 13 border styles (`thin`, `hair`, `dotted`, `dashDotDot`, …).
 * They collapse to a WEIGHT plus a colour here, because that is the honest
 * limit of what a 1-pixel canvas grid line can express: drawing `dashDotDot`
 * as a solid line is already an approximation, and pretending to model the
 * difference between `hair` and `dotted` would be a lie in the type. The
 * original style string is not kept, because the writer re-emits a weight, not
 * the style it was read from — see `docs/KNOWN_LIMITATIONS.md`.
 */
export type CellBorderEdge = {
  readonly weight: 'thin' | 'medium' | 'thick'
  /** `#rrggbb`, or `undefined` for the automatic/theme colour. */
  readonly color: string | undefined
} | undefined

export type CellBorder = {
  readonly top: CellBorderEdge
  readonly right: CellBorderEdge
  readonly bottom: CellBorderEdge
  readonly left: CellBorderEdge
}

export const NO_BORDER: CellBorder = { top: undefined, right: undefined, bottom: undefined, left: undefined }

/** True when any edge has a border. */
export function hasAnyBorder(border: CellBorder): boolean {
  return border.top !== undefined || border.right !== undefined || border.bottom !== undefined || border.left !== undefined
}

/** A cell format, resolved from the `<xf>`/`<font>`/`<fill>`/`<border>` tables into something a renderer can use directly. */
export type ResolvedCellFormat = {
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  readonly strike: boolean
  /** Point size as stored (11 is Excel's default), or `undefined` to inherit. */
  readonly fontSize: number | undefined
  readonly fontName: string | undefined
  /** `#rrggbb`, or `undefined` for the theme/automatic colour. */
  readonly color: string | undefined
  /** Solid-fill background as `#rrggbb`. Pattern fills other than `solid` are ignored rather than approximated. */
  readonly fill: string | undefined
  readonly align: 'left' | 'center' | 'right' | undefined
  readonly valign: 'top' | 'middle' | 'bottom' | undefined
  readonly wrap: boolean
  /**
   * SHEETFMT-3 — the cell's four border edges, each `undefined` for no border.
   *
   * Was a single `bordered: boolean`, which was all the grid could use when it
   * drew no borders at all. Drawing them needs to know WHICH edges and how
   * heavy, so the real shape is read now.
   */
  readonly border: CellBorder
  /** The resolved number-format code (`'0.00%'`, `'dd/mm/yyyy'`, …), or `undefined` for General. */
  readonly numberFormat: string | undefined
}

export const DEFAULT_CELL_FORMAT: ResolvedCellFormat = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  fontSize: undefined,
  fontName: undefined,
  color: undefined,
  fill: undefined,
  align: undefined,
  valign: undefined,
  wrap: false,
  border: NO_BORDER,
  numberFormat: undefined,
}

/** Per-sheet cell formatting: `styleIds[row][col]` indexes into `formats`. `0` is always the default format. */
export type SheetCellStyles = {
  readonly formats: ReadonlyArray<ResolvedCellFormat>
  readonly styleIds: ReadonlyArray<ReadonlyArray<number>>
}

/** Cell formatting for a whole workbook, keyed by worksheet part path (`xl/worksheets/sheet1.xml`). */
export type WorkbookCellStyles = ReadonlyMap<string, SheetCellStyles>

// ---------------------------------------------------------------------------
// styles.xml
// ---------------------------------------------------------------------------

/**
 * The `indexedColors` palette, for a `<color indexed="n">`.
 *
 * ECMA-376 Part 1 §18.8.27 defines these 64 entries as the legacy default
 * palette, used whenever `styles.xml` does not override it with its own
 * `<indexedColors>` (which almost none do). Entries 64/65 are the "system
 * foreground/background" sentinels and deliberately resolve to `undefined`
 * rather than a colour, since what they mean depends on the viewer's own theme
 * — guessing black for 64 would make an automatic-coloured cell unreadable in
 * Atlas's dark themes.
 */
const INDEXED_COLORS: ReadonlyArray<string | undefined> = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
  undefined, undefined,
]

/** A handful of builtin `numFmtId`s that ECMA-376 §18.8.30 fixes by ID and never spells out in `styles.xml`. Mirrors the subset `xlsxPassthrough.ts` keeps for its own auto-typing decision. */
const BUILTIN_NUMFMT_CODES: ReadonlyMap<number, string> = new Map([
  [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'],
  [9, '0%'], [10, '0.00%'], [11, '0.00E+00'],
  [14, 'mm-dd-yy'], [15, 'd-mmm-yy'], [16, 'd-mmm'], [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'], [20, 'h:mm'], [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'], [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'], [40, '#,##0.00;[Red](#,##0.00)'],
  // 41-44 are the accounting formats Excel's own currency buttons produce —
  // among the most common formats in real workbooks, and spelled out nowhere
  // in the file, so leaving them out would report "General" for a column of
  // currency.
  [41, String.raw`_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`],
  [42, String.raw`_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)`],
  [43, String.raw`_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`],
  [44, String.raw`_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`],
  [45, 'mm:ss'], [46, '[h]:mm:ss'], [47, 'mmss.0'],
  [48, '##0.0E+0'], [49, '@'],
])

/**
 * Resolves an OOXML `<color>` to `#rrggbb`.
 *
 * `rgb` is `AARRGGBB` — the alpha is dropped rather than honoured, because a
 * grid cell is opaque and a partly-transparent fill drawn as if it were solid
 * is closer to Excel than one silently skipped. `theme` is deliberately NOT
 * resolved: it indexes the theme part's colour scheme, and picking the wrong
 * entry produces a confidently wrong colour, which is worse than inheriting.
 */
function readColor(el: Element | null): string | undefined {
  if (!el) return undefined
  const rgb = el.getAttribute('rgb')
  if (rgb) {
    const hex = rgb.length === 8 ? rgb.slice(2) : rgb
    if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#${hex.toUpperCase()}`
    return undefined
  }
  const indexed = el.getAttribute('indexed')
  if (indexed !== null) {
    const i = Number(indexed)
    if (Number.isInteger(i) && i >= 0 && i < INDEXED_COLORS.length) return INDEXED_COLORS[i]
  }
  return undefined
}

/** True for `<b/>`, `<b val="1"/>`, `<b val="true"/>`; false for `<b val="0"/>`. Absent means false. */
function readToggle(parent: Element, name: string): boolean {
  const el = firstChildElement(parent, name)
  if (!el) return false
  const val = el.getAttribute('val')
  return val === null || val === '1' || val === 'true'
}

type FontDef = {
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  readonly strike: boolean
  readonly size: number | undefined
  readonly name: string | undefined
  readonly color: string | undefined
}

function readFont(el: Element): FontDef {
  const sizeEl = firstChildElement(el, 'sz')
  const size = sizeEl ? Number(sizeEl.getAttribute('val')) : NaN
  const nameEl = firstChildElement(el, 'name') ?? firstChildElement(el, 'rFont')
  return {
    bold: readToggle(el, 'b'),
    italic: readToggle(el, 'i'),
    // `<u val="none"/>` is an explicit "not underlined", which `readToggle`
    // would report as true (it treats any non-0 val as on).
    underline: (() => {
      const u = firstChildElement(el, 'u')
      if (!u) return false
      return u.getAttribute('val') !== 'none'
    })(),
    strike: readToggle(el, 'strike'),
    size: Number.isFinite(size) && size > 0 ? size : undefined,
    name: nameEl?.getAttribute('val') ?? undefined,
    color: readColor(firstChildElement(el, 'color')),
  }
}

/** A `<fill>`'s solid background, or `undefined` for `none`/`gray125`/any gradient or pattern this does not model. */
function readFill(el: Element): string | undefined {
  const pattern = firstChildElement(el, 'patternFill')
  if (!pattern) return undefined
  const type = pattern.getAttribute('patternType')
  if (type !== 'solid') return undefined
  // In a solid patternFill the VISIBLE colour is fgColor, not bgColor — a
  // detail that is easy to get backwards and produces the wrong colour on
  // exactly the most common case, a plain highlighted cell.
  return readColor(firstChildElement(pattern, 'fgColor')) ?? readColor(firstChildElement(pattern, 'bgColor'))
}

/**
 * Maps an OOXML border style to the weight a canvas line can draw.
 *
 * `double` is deliberately reported as `thick` rather than modelled: drawing
 * two parallel lines inside a grid cell needs more room than a cell border has,
 * and a heavy single line is closer to it than a thin one.
 */
function edgeWeight(style: string | null): 'thin' | 'medium' | 'thick' | null {
  if (!style || style === 'none') return null
  switch (style) {
    case 'hair':
    case 'thin':
    case 'dotted':
    case 'dashed':
    case 'dashDot':
    case 'dashDotDot':
      return 'thin'
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
      return 'medium'
    case 'thick':
    case 'double':
      return 'thick'
    default:
      // An unknown style still means "there IS a border here" — drawing a thin
      // line is far closer than drawing nothing.
      return 'thin'
  }
}

function readEdge(border: Element, name: string): CellBorderEdge {
  const el = firstChildElement(border, name)
  if (!el) return undefined
  const weight = edgeWeight(el.getAttribute('style'))
  if (!weight) return undefined
  return { weight, color: readColor(firstChildElement(el, 'color')) }
}

/** A `<border>`'s four edges. `diagonal` is read and ignored — the grid has no way to draw one. */
function readBorder(el: Element): CellBorder {
  return {
    top: readEdge(el, 'top'),
    right: readEdge(el, 'right'),
    bottom: readEdge(el, 'bottom'),
    left: readEdge(el, 'left'),
  }
}

function readAlign(el: Element | null): Pick<ResolvedCellFormat, 'align' | 'valign' | 'wrap'> {
  if (!el) return { align: undefined, valign: undefined, wrap: false }
  const horizontal = el.getAttribute('horizontal')
  const vertical = el.getAttribute('vertical')
  const wrapText = el.getAttribute('wrapText')
  return {
    // `justify`/`distributed`/`fill`/`centerContinuous` are mapped to the
    // nearest of the three the grid can actually draw rather than dropped.
    align:
      horizontal === 'left' || horizontal === 'justify' || horizontal === 'distributed' || horizontal === 'fill'
        ? 'left'
        : horizontal === 'center' || horizontal === 'centerContinuous'
          ? 'center'
          : horizontal === 'right'
            ? 'right'
            : undefined,
    valign:
      vertical === 'top'
        ? 'top'
        : vertical === 'center'
          ? 'middle'
          : vertical === 'bottom' || vertical === 'justify' || vertical === 'distributed'
            ? 'bottom'
            : undefined,
    wrap: wrapText === '1' || wrapText === 'true',
  }
}

/** The format tables from `xl/styles.xml`, with each `cellXfs` entry already resolved. */
export type StyleTable = {
  /** Resolved formats, parallel to `<cellXfs>`'s own `<xf>` order — so index `n` here IS a cell's `s="n"`. */
  readonly xfs: ReadonlyArray<ResolvedCellFormat>
}

/**
 * Parses `xl/styles.xml` into one resolved format per `<cellXfs>` entry.
 *
 * `applyFont`/`applyFill`/`applyBorder`/`applyAlignment` are deliberately
 * ignored. Excel writes them inconsistently — plenty of real files set
 * `fontId` to a bold font with no `applyFont="1"` and Excel still renders it
 * bold — so honouring them loses formatting that is visibly there. The `xfId`
 * cell-style parent is not followed either: resolving it needs `cellStyleXfs`
 * and matters only for named styles, and a wrong inheritance chain is worse
 * than a flat read.
 */
export function parseStyleTable(stylesXml: string): StyleTable | null {
  let doc: XMLDocument
  try {
    doc = parseXmlPart(stylesXml)
  } catch {
    return null
  }
  const root = doc.documentElement
  const cellXfsEl = firstChildElement(root, 'cellXfs')
  if (!cellXfsEl) return null

  const fontsEl = firstChildElement(root, 'fonts')
  const fonts = fontsEl ? childElements(fontsEl, 'font').map(readFont) : []

  const fillsEl = firstChildElement(root, 'fills')
  const fills = fillsEl ? childElements(fillsEl, 'fill').map(readFill) : []

  const bordersEl = firstChildElement(root, 'borders')
  const borders = bordersEl ? childElements(bordersEl, 'border').map(readBorder) : []

  const numFmtCodes = new Map<number, string>()
  const numFmtsEl = firstChildElement(root, 'numFmts')
  if (numFmtsEl) {
    for (const el of childElements(numFmtsEl, 'numFmt')) {
      const id = Number(el.getAttribute('numFmtId'))
      if (Number.isFinite(id)) numFmtCodes.set(id, el.getAttribute('formatCode') ?? '')
    }
  }

  // `fonts[0]` is the workbook's DEFAULT font, not a style: virtually every
  // `<xf>` in a real file references it, and it almost always states a size
  // ("11") and a name ("Calibri"). Reporting those as if the author had chosen
  // them per cell would make Atlas override its own theme font size and family
  // for every workbook ever opened — including ones with no styling at all.
  // So font 0 is the baseline, and only what DIFFERS from it is a format.
  // Bold/italic/underline/strike stay absolute: a default font that is itself
  // bold means the whole sheet really is bold.
  const baseFont = fonts[0]

  const xfs = childElements(cellXfsEl, 'xf').map((xf): ResolvedCellFormat => {
    const fontId = Number(xf.getAttribute('fontId') ?? '')
    const font = Number.isInteger(fontId) ? fonts[fontId] : undefined
    const fillId = Number(xf.getAttribute('fillId') ?? '')
    const fill = Number.isInteger(fillId) ? fills[fillId] : undefined
    const borderId = Number(xf.getAttribute('borderId') ?? '')
    const border = Number.isInteger(borderId) ? (borders[borderId] ?? NO_BORDER) : NO_BORDER
    const numFmtId = Number(xf.getAttribute('numFmtId') ?? '0')
    const code = Number.isFinite(numFmtId) && numFmtId !== 0
      ? (numFmtCodes.get(numFmtId) ?? BUILTIN_NUMFMT_CODES.get(numFmtId))
      : undefined

    return {
      bold: font?.bold ?? false,
      italic: font?.italic ?? false,
      underline: font?.underline ?? false,
      strike: font?.strike ?? false,
      fontSize: font?.size === baseFont?.size ? undefined : font?.size,
      fontName: font?.name === baseFont?.name ? undefined : font?.name,
      color: font?.color === baseFont?.color ? undefined : font?.color,
      fill,
      border,
      numberFormat: code === '' ? undefined : code,
      ...readAlign(firstChildElement(xf, 'alignment')),
    }
  })

  return { xfs }
}

// ---------------------------------------------------------------------------
// which cell uses which
// ---------------------------------------------------------------------------

/** `A1` -> `{row: 0, col: 0}`, or `null` for anything that is not a plain cell reference. */
function decodeRef(ref: string): { row: number; col: number } | null {
  const match = /^([A-Z]+)(\d+)$/.exec(ref)
  if (!match) return null
  let col = 0
  for (const ch of match[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  const row = Number(match[2])
  if (!Number.isInteger(row) || row < 1) return null
  return { row: row - 1, col: col - 1 }
}

/**
 * Reads each cell's `s` style index out of one worksheet part.
 *
 * Deliberately a regex scan rather than a DOM parse. A worksheet is the one
 * part of a workbook that can be enormous — this repo's own perf tests use a
 * 100k-row sheet — and building a DOM for it here would mean a second full
 * parse of the largest thing in the file purely to read one attribute per
 * cell. The grammar being matched is tiny and fixed: `<c r="A1" s="3" …>`,
 * where both attributes are always simple double-quoted values in every
 * producer's output, and a `<c>` with no `s` is the default format.
 *
 * `offsetRow`/`offsetCol` shift the result into the grid's own coordinates,
 * which start at the sheet's used range rather than at A1 (`sheetToGrid`
 * decodes `!ref` the same way).
 */
export function readSheetStyleIds(
  sheetXml: string,
  rowCount: number,
  colCount: number,
  offsetRow: number,
  offsetCol: number,
): number[][] {
  const styleIds: number[][] = Array.from({ length: rowCount }, () => new Array<number>(colCount).fill(0))

  // One `<c …>` opening tag at a time; `[^>]*` cannot escape the tag because
  // an attribute value may not contain a raw `>` in well-formed XML.
  const cellPattern = /<c\s([^>]*)>/g
  let match: RegExpExecArray | null
  while ((match = cellPattern.exec(sheetXml)) !== null) {
    const attrs = match[1]
    const styleMatch = /\bs="(\d+)"/.exec(attrs)
    if (!styleMatch) continue // no style index: the default format, already 0
    const refMatch = /\br="([A-Z]+\d+)"/.exec(attrs)
    if (!refMatch) continue
    const at = decodeRef(refMatch[1])
    if (!at) continue
    const row = at.row - offsetRow
    const col = at.col - offsetCol
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) continue
    styleIds[row][col] = Number(styleMatch[1])
  }

  return styleIds
}

/**
 * Reads cell formatting for every worksheet in an xlsx/xlsm buffer.
 *
 * Returns `null` — not a partial result, and never a throw — when the package
 * has no readable `styles.xml`, which also covers every non-OOXML format
 * (`.xls`, `.ods`, `.csv`) reaching this by accident. The caller renders
 * unstyled, exactly as before this existed.
 */
export async function readWorkbookCellStyles(
  buffer: ArrayBuffer,
  sheets: ReadonlyArray<{ readonly sourcePath?: string; readonly rowCount: number; readonly colCount: number; readonly offsetRow: number; readonly offsetCol: number }>,
): Promise<{ readonly table: StyleTable; readonly styles: WorkbookCellStyles } | null> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch {
    return null
  }

  const stylesFile = zip.file('xl/styles.xml')
  if (!stylesFile) return null
  let table: StyleTable | null
  try {
    table = parseStyleTable(await stylesFile.async('string'))
  } catch {
    return null
  }
  if (!table) return null

  const styles = new Map<string, SheetCellStyles>()
  for (const sheet of sheets) {
    if (!sheet.sourcePath) continue
    const file = zip.file(sheet.sourcePath)
    if (!file) continue
    try {
      const xml = await file.async('string')
      styles.set(sheet.sourcePath, {
        formats: table.xfs,
        styleIds: readSheetStyleIds(xml, sheet.rowCount, sheet.colCount, sheet.offsetRow, sheet.offsetCol),
      })
    } catch {
      // One unreadable worksheet must not cost the whole workbook its
      // formatting — that sheet simply renders unstyled.
    }
  }

  return { table, styles }
}

/** The format for one cell, or the default when there is no styling for it. */
export function formatAt(styles: SheetCellStyles | undefined, row: number, col: number): ResolvedCellFormat {
  if (!styles) return DEFAULT_CELL_FORMAT
  const id = styles.styleIds[row]?.[col] ?? 0
  return styles.formats[id] ?? DEFAULT_CELL_FORMAT
}
