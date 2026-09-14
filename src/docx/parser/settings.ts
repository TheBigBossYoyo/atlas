/**
 * Atlas — `word/settings.xml` parser (D17 / DXE-11, extended for D11/D24)
 *
 * `word/settings.xml` passes through the save pipeline byte-for-byte
 * (`docx/index.ts`'s `rawArchive` passthrough) because Atlas's `Document`
 * model doesn't own most of what it contains. Atlas reads back a small
 * handful of document-level flags the layout/editor pipeline needs; every
 * other setting stays untouched on save:
 *
 * - `trackChanges` (D17/DXE-11) — seeds the toolbar's Track Changes toggle.
 * - `evenAndOddHeaders` (D11/DXL-09) — whether a section's "even page"
 *   header/footer reference is actually honored; when absent/off, even
 *   pages fall back to the same header/footer as odd pages even if the
 *   document happens to define an `even`-typed reference.
 * - `autoHyphenation` (D24/DXL-18) — parsed and exposed for a future
 *   pattern-based hyphenator, but not yet consumed by the layout pipeline:
 *   no small, permissively-licensed hyphenation pattern set was available
 *   to bundle without adding a new dependency (this worktree cannot run
 *   `npm install`), so DXL-18 was resolved by keeping Atlas's EXISTING
 *   explicit-soft-hyphen (`­`) and discretionary-break support (see
 *   `layout/itemize.ts`'s `SOFT_HYPHEN` handling, present before this task)
 *   as the full extent of hyphenation support — a document that never
 *   authors an explicit soft hyphen sees no automatic mid-word breaks
 *   whether or not this flag is on. Documented scope limitation, not a bug.
 * - `footnotePr`/`endnotePr` (D11 milestone 5) — the document's default
 *   footnote/endnote numbering format and restart rule.
 */
import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

/** OOXML `ST_RestartNumber` (`CT_NumRestart`) — when a footnote/endnote numbering counter resets. */
export type NoteNumberRestart = 'continuous' | 'eachSect' | 'eachPage'

/**
 * A document's default footnote or endnote numbering behavior
 * (`w:footnotePr`/`w:endnotePr` in `word/settings.xml`). `numFmt` is kept as
 * the raw `w:val` string (mirrors `PageNumberType.fmt`'s convention in
 * `model/document.ts`) rather than a closed enum, since OOXML defines many
 * numbering-format values and Atlas's formatter only needs to recognize the
 * common ones — see `layout/noteNumbering.ts`'s `formatSequenceNumber`.
 */
export interface NotePr {
  readonly numFmt?: string
  readonly restart?: NoteNumberRestart
  readonly start?: number
}

export interface SettingsPart {
  readonly trackChanges: boolean
  readonly evenAndOddHeaders?: boolean
  readonly autoHyphenation?: boolean
  readonly footnotePr?: NotePr
  readonly endnotePr?: NotePr
}

type OnOffNode = string | number | { readonly '@_w:val'?: string | number | boolean }

interface RawNumFmtNode {
  readonly '@_w:val'?: string
}

interface RawDecimalNode {
  readonly '@_w:val'?: string | number
}

interface RawNotePrNode {
  readonly 'w:numFmt'?: RawNumFmtNode
  readonly 'w:numStart'?: RawDecimalNode
  readonly 'w:numRestart'?: RawDecimalNode
}

interface RawSettings {
  'w:settings'?: {
    'w:trackChanges'?: OnOffNode
    'w:evenAndOddHeaders'?: OnOffNode
    'w:autoHyphenation'?: OnOffNode
    'w:footnotePr'?: RawNotePrNode
    'w:endnotePr'?: RawNotePrNode
  }
}

function isOffValue(value: string | number | boolean): boolean {
  const normalized = String(value).trim().toLowerCase()
  return normalized === 'false' || normalized === '0' || normalized === 'off'
}

/** Shared on/off-element convention (OOXML `CT_OnOff`): present means on unless `w:val` says otherwise. */
function parseOnOff(node: OnOffNode | undefined): boolean | undefined {
  if (node === undefined) {
    return undefined
  }

  if (typeof node === 'object') {
    const val = node['@_w:val']
    return val === undefined ? true : !isOffValue(val)
  }

  return true
}

function normalizeRestart(value: string | number | undefined): NoteNumberRestart | undefined {
  if (value === 'continuous' || value === 'eachSect' || value === 'eachPage') {
    return value
  }
  return undefined
}

function parseNotePr(node: RawNotePrNode | undefined): NotePr | undefined {
  if (node === undefined) {
    return undefined
  }

  const numFmt = node['w:numFmt']?.['@_w:val']
  const restart = normalizeRestart(node['w:numRestart']?.['@_w:val'])
  const rawStart = node['w:numStart']?.['@_w:val']
  const start = rawStart === undefined ? undefined : Number(rawStart)

  if (numFmt === undefined && restart === undefined && start === undefined) {
    return undefined
  }

  return {
    ...(numFmt !== undefined ? { numFmt } : {}),
    ...(restart !== undefined ? { restart } : {}),
    ...(start !== undefined && !Number.isNaN(start) ? { start } : {}),
  }
}

/**
 * Parses the XML content of `word/settings.xml`.
 *
 * @param xml - UTF-8 text of the settings part.
 * @returns    Document-level flags Atlas's layout/editor pipeline reads
 *   back. Every optional field is omitted (not `false`/`undefined` present)
 *   when the source document doesn't declare it, matching the model's
 *   existing "absent means not specified" convention elsewhere (e.g.
 *   `Comment.resolved`).
 * @throws     `DocxParseError` on malformed XML.
 */
export function parseSettings(xml: string): SettingsPart {
  assertXmlPartSizeWithinLimit(xml, 'word/settings.xml')

  let parsed: RawSettings
  try {
    parsed = xmlParser.parse(xml) as RawSettings
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse word/settings.xml: ${msg}`, 'word/settings.xml')
  }

  const settings = parsed['w:settings']
  const trackChanges = parseOnOff(settings?.['w:trackChanges']) ?? false
  const evenAndOddHeaders = parseOnOff(settings?.['w:evenAndOddHeaders'])
  const autoHyphenation = parseOnOff(settings?.['w:autoHyphenation'])
  const footnotePr = parseNotePr(settings?.['w:footnotePr'])
  const endnotePr = parseNotePr(settings?.['w:endnotePr'])

  return {
    trackChanges,
    ...(evenAndOddHeaders !== undefined ? { evenAndOddHeaders } : {}),
    ...(autoHyphenation !== undefined ? { autoHyphenation } : {}),
    ...(footnotePr !== undefined ? { footnotePr } : {}),
    ...(endnotePr !== undefined ? { endnotePr } : {}),
  }
}
