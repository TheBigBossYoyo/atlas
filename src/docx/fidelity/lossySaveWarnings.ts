/**
 * Atlas — lossy-save detection (round-trip fidelity audit, DXS round 2).
 *
 * Atlas's DOCX model is a curated subset of OOXML: anything a real `.docx`
 * uses that the parser has no field for is invisible to the in-memory
 * document, and a save that regenerates a part *from the model* — as
 * `saveDocx` always does for `word/document.xml`/`word/styles.xml`/
 * `word/numbering.xml` and every header/footer part — emits that part
 * without it. That is silent data loss, indistinguishable from "the user
 * meant to delete this" unless something flags it before the file on disk
 * is overwritten.
 *
 * Rather than maintain a hand-written blocklist of "elements Atlas doesn't
 * support" (which rots the instant a parser gains a new field, in either
 * direction), this does a real, empirical check: it compares every element
 * name present in a model-rewritten part's ORIGINAL xml against the same
 * part's FRESHLY SAVED xml, for a save produced from the document exactly
 * as loaded (no user edits). Since nothing was intentionally changed, the
 * two tag-name multisets should be identical; any element name that drops
 * to zero occurrences is a genuine, previously-invisible fidelity gap —
 * not a proxy for one. This stays correct as parser coverage grows without
 * anyone having to remember to update a list here.
 *
 * Deliberately takes `savedBytes` as a parameter rather than calling
 * `saveDocx` itself: every real call site (a Save action) already has both
 * the loaded `DocxBundle` and the bytes `saveDocx` just produced, so this
 * avoids a second, redundant serialization pass and avoids this module
 * importing back from `../index` (which exports `saveDocx`).
 *
 * Usage (left for whichever layer owns the actual Save UI — this module
 * only detects, it never warns or blocks on its own):
 *
 *   const saved = await saveDocx(bundle)
 *   const warnings = await detectLossySaveWarnings(bundle, saved)
 *   if (warnings.length > 0) {
 *     // surface `describeLossySaveWarnings(warnings)` to the user before
 *     // committing `saved` to disk, e.g. a confirm dialog listing what
 *     // Atlas cannot preserve in this specific document.
 *   }
 */
import { XMLParser } from 'fast-xml-parser'
import JSZip from 'jszip'

import type { DocxBundle } from '../index'

/** One OOXML element that was present in the loaded document but did not survive an unedited save. */
export interface LossySaveWarning {
  /** The part it disappeared from, e.g. `word/document.xml`. */
  readonly part: string
  /** The qualified element name, e.g. `w:pgBorders`. */
  readonly tag: string
  /** How many occurrences existed in the original part. */
  readonly occurrences: number
}

// Parts `saveDocx` always regenerates from the in-memory model (never a raw
// passthrough) for which "no user edit happened, so nothing should
// disappear" holds unconditionally. `word/comments.xml`/`footnotes.xml`/
// `endnotes.xml` are deliberately excluded: `saveDocx` removes them outright
// when the model has zero comments/notes (an intentional cleanup — see
// `docx/index.ts`'s "Wave 1 follow-up" comments), which would read as false
// positives here.
const ALWAYS_REWRITTEN_PARTS: ReadonlyArray<string> = [
  'word/document.xml',
  'word/styles.xml',
  'word/numbering.xml',
]

/**
 * Pure "optional grouping container" elements: every property they could
 * carry lives in their own optional attributes/children, and Atlas's parser
 * already extracts all of those individually (see the doc comment next to
 * each corresponding model field). When an occurrence of one of these tags
 * in the original XML is entirely empty — no attributes, no children — its
 * disappearance on save is provably semantics-preserving (Word itself
 * treats an empty `<w:pgNumType/>` identically to no `w:pgNumType` at all):
 * confirmed by generating a real corpus fixture with the `docx` npm package
 * and observing exactly this shape. This is deliberately not "every
 * self-closing element" — a genuinely meaningful toggle element like
 * `<w:b/>` or `<w:rtlGutter/>` is *also* attribute-less and child-less, but
 * its mere presence is its whole meaning, so it must never be added here.
 * Only add a tag to this list after verifying (the way this list's current
 * three entries were) that Atlas's parser has a dedicated field for
 * everything that tag can ever carry.
 */
const SAFE_WHEN_EMPTY_TAGS: ReadonlySet<string> = new Set([
  'w:rPrDefault',
  'w:pPrDefault',
  'w:pgNumType',
])

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text' })

interface TagStats {
  count: number
  /** True only if every occurrence seen so far had zero attributes and zero children. */
  allEmpty: boolean
}

function isEmptyOccurrence(value: unknown): boolean {
  if (value === undefined || value === null) return true
  if (typeof value === 'object' && !Array.isArray(value)) {
    return Object.keys(value as Record<string, unknown>).length === 0
  }
  // A bare string/number/boolean value means an element with no attributes
  // and no child elements, holding at most (possibly empty) text.
  return typeof value !== 'object'
}

function countTags(xml: string, into: Map<string, TagStats>): void {
  let tree: unknown
  try {
    tree = xmlParser.parse(xml)
  } catch {
    // Malformed XML is `validateDocxPackage`'s job to catch, not this
    // module's — skip counting rather than throw from a detector.
    return
  }
  walk(tree, into)
}

function walk(node: unknown, into: Map<string, TagStats>): void {
  if (node === null || typeof node !== 'object') {
    return
  }
  if (Array.isArray(node)) {
    for (const item of node) walk(item, into)
    return
  }
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key.startsWith('@_') || key === '#text' || key === '#comment') {
      continue
    }
    const occurrences = Array.isArray(value) ? value : [value]
    const existing = into.get(key) ?? { count: 0, allEmpty: true }
    for (const occurrence of occurrences) {
      existing.count += 1
      existing.allEmpty = existing.allEmpty && isEmptyOccurrence(occurrence)
    }
    into.set(key, existing)
    walk(value, into)
  }
}

async function readPart(zip: JSZip, path: string): Promise<string | undefined> {
  const entry = zip.file(path)
  if (entry === null) {
    return undefined
  }
  return entry.async('string')
}

/**
 * Compares `bundle`'s originally-loaded parts against `savedBytes` (the
 * output of `saveDocx(bundle)`, called with no intervening edits) and
 * reports every element name that was present in a model-rewritten part
 * and is entirely absent from that same part after the round trip.
 */
export async function detectLossySaveWarnings(
  bundle: DocxBundle,
  savedBytes: Uint8Array,
): Promise<ReadonlyArray<LossySaveWarning>> {
  if (bundle.rawArchive === undefined) {
    return []
  }

  const savedZip = await JSZip.loadAsync(savedBytes)
  const partsToCheck = [...ALWAYS_REWRITTEN_PARTS, ...headerFooterPartPaths(bundle)]

  const warnings: LossySaveWarning[] = []
  for (const partPath of partsToCheck) {
    const originalBytes = bundle.rawArchive.get(partPath)
    if (originalBytes === undefined) {
      continue
    }
    const originalXml = new TextDecoder().decode(originalBytes)
    const savedXml = await readPart(savedZip, partPath)
    if (savedXml === undefined) {
      // The whole part vanished — a bigger problem than any one element,
      // and one `validateDocxPackage`/the relationship-integrity checks
      // already catch structurally; not this module's concern.
      continue
    }

    const originalCounts = new Map<string, TagStats>()
    countTags(originalXml, originalCounts)
    const savedCounts = new Map<string, TagStats>()
    countTags(savedXml, savedCounts)

    for (const [tag, stats] of originalCounts) {
      if (savedCounts.has(tag)) {
        continue
      }
      if (SAFE_WHEN_EMPTY_TAGS.has(tag) && stats.allEmpty) {
        continue
      }
      warnings.push({ part: partPath, tag, occurrences: stats.count })
    }
  }

  return warnings
}

/** Every header/footer part the loaded document's own relationships point at. */
function headerFooterPartPaths(bundle: DocxBundle): ReadonlyArray<string> {
  const paths: string[] = []
  for (const rel of bundle.relationships ?? []) {
    if (rel.type.endsWith('/header') || rel.type.endsWith('/footer')) {
      paths.push(`word/${rel.target.replace(/^\//, '')}`)
    }
  }
  return paths
}

/** Renders warnings as short, user-facing lines, one per (part, tag), grouped by part. */
export function describeLossySaveWarnings(warnings: ReadonlyArray<LossySaveWarning>): ReadonlyArray<string> {
  return warnings.map(
    (warning) =>
      `${warning.part}: ${warning.tag} (${warning.occurrences} occurrence${warning.occurrences === 1 ? '' : 's'}) is not supported and will be removed`,
  )
}
