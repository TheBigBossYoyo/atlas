/**
 * Atlas — lossy-save detection (round-trip fidelity audit, DXS round 2).
 *
 * Atlas's DOCX model is a curated subset of OOXML: anything a real `.docx`
 * uses that the parser has no field for is invisible to the in-memory
 * document, and a save that regenerates a part *from the model* — as
 * `saveDocx` always does for `word/document.xml`/`word/styles.xml`/
 * `word/numbering.xml` and every header/footer part, and does whenever the
 * model's collection is non-empty for `word/comments.xml`/`footnotes.xml`/
 * `endnotes.xml` — emits that part without it. That is silent data loss,
 * indistinguishable from "the user meant to delete this" unless something
 * flags it before the file on disk is overwritten.
 *
 * Rather than maintain a hand-written blocklist of "elements Atlas doesn't
 * support" (which rots the instant a parser gains a new field, in either
 * direction), this does a real, empirical check: it compares every element
 * name (and, since DXS round-2 follow-up "B4", every attribute VALUE that
 * name carries) present in a model-rewritten part's ORIGINAL xml against
 * the same part's FRESHLY SAVED xml, for a save produced from the document
 * exactly as loaded (no user edits). Since nothing was intentionally
 * changed, the two tag-name multisets — and each tag's per-attribute value
 * multisets — should be identical; any element name whose occurrence count
 * drops (to zero, or merely below what it was), or any attribute value that
 * disappears from every surviving occurrence of its tag, is a genuine,
 * previously-invisible fidelity gap — not a proxy for one. This stays
 * correct as parser coverage grows without anyone having to remember to
 * update a list here.
 *
 * B4's own motivating gap: a count-only-to-zero check sees a `w:rFonts`
 * that's still there (because it also carries literal `w:ascii`/etc.
 * attributes) as nothing having happened at all, even though its
 * `w:asciiTheme` value silently vanished (exactly DOCX-1's shape) — or sees
 * 12 `w:tab` become 1 as "the tag is still present" — nothing but the
 * attribute-value and count-DELTA checks below can catch either. Both are
 * still deliberately conservative: a tag's occurrence count is only ever
 * flagged for DECREASING (growth — the far more common real-edit shape —
 * is never reported, see the loop below), and a handful of attributes whose
 * value is renumbering/bookkeeping rather than content are excluded by name
 * (`IDENTITY_ATTRIBUTE_NAMES`/`TAG_SCOPED_IDENTITY_ATTRIBUTES`/
 * `isNamespacePlumbingAttribute`) so legitimate re-serialization churn
 * — reordered attributes, renumbered ids, normalized namespace
 * declarations — never warns.
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

/**
 * `'element'` — an element name (`tag`) dropped in occurrence count between
 * the original part and the saved one: to zero (fully gone — `savedOccurrences`
 * is `undefined`) or merely fewer (`savedOccurrences` holds what survived).
 * `'attribute'` — the element itself is still there the same number of
 * times, but one value that `attribute` held on it in the original no
 * longer appears on ANY occurrence of `tag` in the saved part.
 */
export type LossySaveWarningKind = 'element' | 'attribute'

/** One OOXML element (or one of its attribute values) that was present in the loaded document but did not survive an unedited save. */
export interface LossySaveWarning {
  /** The part it disappeared from, e.g. `word/document.xml`. */
  readonly part: string
  /** The qualified element name, e.g. `w:pgBorders`. */
  readonly tag: string
  readonly kind: LossySaveWarningKind
  /**
   * How many occurrences (of the element, for `kind: 'element'`; of this
   * specific attribute value, for `kind: 'attribute'`) existed in the
   * original part.
   */
  readonly occurrences: number
  /** `kind: 'element'` only, and only when some (not all) occurrences survived: how many are left in the saved part. */
  readonly savedOccurrences?: number
  /** `kind: 'attribute'` only: the attribute name, e.g. `w:asciiTheme`. */
  readonly attribute?: string
}

// Parts `saveDocx` always regenerates from the in-memory model (never a raw
// passthrough) for which "no user edit happened, so nothing should
// disappear" holds unconditionally.
const ALWAYS_REWRITTEN_PARTS: ReadonlyArray<string> = [
  'word/document.xml',
  'word/styles.xml',
  'word/numbering.xml',
]

/**
 * `word/comments.xml`/`word/footnotes.xml`/`word/endnotes.xml` — checked
 * separately from `ALWAYS_REWRITTEN_PARTS` because, unlike those three,
 * `saveDocx` does NOT unconditionally regenerate these from the model: it
 * only does so when the corresponding model collection is non-empty (see
 * `docx/index.ts`'s "Wave 1 follow-up" / "5. Comments" / "6. Footnotes /
 * endnotes" steps), and for comments specifically, an emptied model
 * actively *removes* the part (`removePartRegistration`) rather than
 * leaving it as a raw-archive passthrough.
 *
 * That asymmetry is exactly why these three were left out of this
 * detector entirely at first: naively diffing "original part" against
 * "saved part" reads as a false positive the moment the part legitimately
 * disappears — a document with zero footnotes has no `footnotes.xml` to
 * begin with, and a user who deletes every comment should never be told
 * their comments were "lost".
 *
 * It turns out both of those "legitimately absent" shapes are already
 * covered by this module's two existing guards below, with no part-type-
 * specific logic needed:
 *   - No original part at all (the overwhelmingly common "never had one"
 *     case: `bundle.rawArchive.get(partPath) === undefined`) — the loop
 *     `continue`s before ever comparing anything.
 *   - An original part that legitimately vanishes from the saved output
 *     (comments.xml removed because the model went to zero comments;
 *     footnotes.xml/endnotes.xml left as an untouched raw-archive
 *     passthrough — byte-identical, hence tag-identical — because the
 *     model's collection was already empty) — `readPart(savedZip, ...)`
 *     returns `undefined` and the loop `continue`s there instead.
 * Verified empirically against the full fixture corpus (see this module's
 * own test file) and with dedicated regression tests for both shapes
 * before adding these three part paths here.
 */
const NOTES_AND_COMMENTS_PARTS: ReadonlyArray<string> = [
  'word/comments.xml',
  'word/footnotes.xml',
  'word/endnotes.xml',
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
  // B4 — `a:srcRect` (CT_RelativeRect, an image's crop rectangle) can only
  // ever carry its four optional `l`/`t`/`r`/`b` attributes, each defaulting
  // to 0 (no crop) when absent, and never any children — confirmed both
  // from the schema and empirically: the `image-crop-rotation-flip` corpus
  // fixture's own source XML has a real `<a:srcRect l="…" t="…" r="…"
  // b="…"/>` immediately followed by a second, entirely redundant bare
  // `<a:srcRect/>` (a `docx`-npm-package generation artifact), and Atlas's
  // save correctly keeps only the first — a same-tag COUNT drop (2 -> 1)
  // that carries zero information loss, exactly this list's own criterion.
  'a:srcRect',
])

/**
 * Attribute names whose VALUE is bookkeeping/identity churn a save may
 * legitimately rewrite without losing anything a user could notice — as
 * opposed to `w:asciiTheme`/`w:outline`/etc., whose value (or presence) IS
 * the content. Excluded from attribute-VALUE comparison only: the
 * attribute's mere presence/absence still feeds the ordinary element-level
 * checks above exactly as before (an `<w:id w:val="…"/>` content-control
 * element disappearing is still reported — this only silences "the id
 * NUMBER changed").
 *
 * Each entry below is deliberate, not a blanket "ids don't matter" rule:
 *   - `w:id` — the cross-reference key for a bookmark pair, a footnote/
 *     endnote/comment reference<->definition link, or a tracked-change
 *     (`w:ins`/`w:del`) marker. The number itself carries no document
 *     content; only internal consistency between the two ends of the link
 *     does, which this element-name-flattened, single-part diff was never
 *     positioned to verify anyway. Atlas currently round-trips these
 *     verbatim (parsed straight off the source attribute — see
 *     parser/document.ts), so today this exclusion costs nothing in
 *     practice; it documents that a future serializer free to renumber them
 *     (the way it already does for `wp:docPr/@id` below) would lose nothing
 *     a user can see.
 *   - `r:id` / `r:embed` — a relationship-table reference (hyperlink,
 *     image, header/footer). The string is an index into that part's own
 *     `.rels` file, not content; a resolvable-but-different id points at
 *     the exact same target.
 *   - `w:rsidR`/`w:rsidRDefault`/`w:rsidP`/`w:rsidRPr`/`w:rsidDel`/
 *     `w:rsidRoot`/`w:rsidTr` — Word's own per-editing-session revision-save
 *     ID. Real Word regenerates these on essentially every save; they group
 *     edits for Word's own conflict UI and carry no document content ever.
 *   - `w14:paraId`/`w14:textId` — Word 2010+'s per-paragraph/run tracking
 *     ids (co-authoring/commenting anchors), the same bookkeeping category
 *     as `w:rsid*` above.
 */
const IDENTITY_ATTRIBUTE_NAMES: ReadonlySet<string> = new Set([
  'w:id',
  'r:id',
  'r:embed',
  'w:rsidR',
  'w:rsidRDefault',
  'w:rsidP',
  'w:rsidRPr',
  'w:rsidDel',
  'w:rsidRoot',
  'w:rsidTr',
  'w14:paraId',
  'w14:textId',
])

/**
 * `(tag, attribute)` pairs excluded from value comparison for a reason
 * specific to that one element, expressed as `${tag}/@${attribute}`. Each
 * one below was confirmed by reading the exact writer code that produces
 * it, not guessed from a false-positive and papered over — see this
 * module's test file / this task's own report for the corpus fixtures that
 * would otherwise have false-positived on each.
 *
 *   - `wp:docPr/@id` — `buildDocPrNode` (documentWriter.ts) hardcodes this
 *     to the literal string `"1"` unconditionally; Atlas never reads the
 *     source drawing's own docPr id back for anything. Diffing it would
 *     warn on essentially every image/shape in every document that isn't
 *     the package's first drawing, for a value nothing depends on.
 *   - `w:vMerge/@w:val` — `buildTableCellMergeElement` (documentWriter.ts)
 *     deliberately emits a bare `<w:vMerge/>` (no `w:val` at all) for
 *     `'continue'`, per CT_VMerge's own schema default (`w:val` absent
 *     means `continue`) — the identical shorthand OOXML itself defines,
 *     not data loss.
 *   - `pic:cNvPr/@name` / `pic:cNvPr/@descr` — `buildGraphicNode`
 *     (documentWriter.ts) hardcodes `pic:cNvPr`'s own `name`/`descr` to
 *     `""` unconditionally; the doc comment right above that call explains
 *     why: a drawing's name/title/description live on `wp:docPr` instead
 *     (`buildDocPrNode`, which DOES emit `drawing.name`/`drawing.description`
 *     verbatim), and `pic:cNvPr`'s own copy — real Word/DOCX-generator
 *     output typically mirrors the same string on both elements — is
 *     treated as the redundant one. The same information is still checked
 *     (and would still be caught if lost) via `wp:docPr`'s own `name`/
 *     `descr` attributes, which this exclusion does NOT cover.
 */
const TAG_SCOPED_IDENTITY_ATTRIBUTES: ReadonlySet<string> = new Set([
  'wp:docPr/@id',
  'w:vMerge/@w:val',
  'pic:cNvPr/@name',
  'pic:cNvPr/@descr',
])

/**
 * Strips a qualified attribute name's namespace prefix (`w:parentId` ->
 * `parentId`), for `survivingValueCount`'s prefix-insensitive matching.
 */
function localAttributeName(name: string): string {
  const colonIndex = name.indexOf(':')
  return colonIndex === -1 ? name : name.slice(colonIndex + 1)
}

/**
 * OOXML's boolean type (`ST_OnOff`) accepts six interchangeable spellings
 * for the same two values — `1`/`true`/`on` and `0`/`false`/`off` — and
 * real `.docx` producers use whichever they please (this corpus's own
 * fixtures write `w:tblLook`'s six flags and `w:cols/@w:sep` as `"true"`/
 * `"false"`, and `a:xfrm/@flipH` the same way). `buildOnOffAttribute`
 * (documentWriter.ts) always normalizes Atlas's own OUTPUT to `1`/`0`
 * regardless of which spelling the source used — confirmed the same
 * "true" -> "1" rewrite happens for a genuinely boolean value with no
 * modeling gap at all (`w:rtl/@w:val`, `a:xfrm/@flipH`), not just for ones
 * this module already excludes for other reasons. Applied unconditionally
 * to every attribute value (not tag/attribute-scoped): this is a document-
 * wide OOXML lexical convention, not one element's own writer's quirk, so
 * scoping it to today's known offenders would just mean rediscovering the
 * identical false positive the next time it shows up on an attribute this
 * module hasn't seen yet. The six spellings only ever canonicalize to one
 * of exactly two buckets, so this can never mask a value genuinely
 * flipping between two DIFFERENT semantic values (`"1"` and `"5"` stay
 * distinct; only `"1"`/`"true"`/`"on"` ever collapse together).
 */
const ON_OFF_CANONICAL: ReadonlyMap<string, string> = new Map([
  ['1', 'true'],
  ['true', 'true'],
  ['on', 'true'],
  ['0', 'false'],
  ['false', 'false'],
  ['off', 'false'],
])

/**
 * `w:jc`/`w:lvlJc`'s shared `w:val` (`ST_Jc`) accepts the legacy `left`/
 * `right` as synonyms for the ECMA-376-preferred `start`/`end` —
 * `parseJustifyContent` (parser/document.ts, parser/numbering.ts,
 * parser/styles.ts — three independent copies, same mapping) canonicalizes
 * every source spelling before it ever reaches the model, so Atlas's own
 * writer only ever emits the canonical form; a source paragraph/numbering
 * level explicitly using `"left"`/`"right"` (the overwhelmingly common
 * spelling in real-world `.docx` files) would otherwise false-positive on
 * every single one. Scoped to exactly the two tags `parseJustifyContent`'s
 * callers parse, both via their own `w:val` — deliberately NOT a blanket
 * "left means start everywhere" rule; an unrelated attribute that happens
 * to also spell a value `"left"`/`"right"` for a non-alignment purpose
 * (e.g. a border side) is untouched.
 */
const JC_CANONICAL: ReadonlyMap<string, string> = new Map([
  ['left', 'start'],
  ['start', 'start'],
  ['right', 'end'],
  ['end', 'end'],
])
const JC_TAGS: ReadonlySet<string> = new Set(['w:jc', 'w:lvlJc'])

/** Canonicalizes an attribute value for COMPARISON purposes only (reported warnings always keep the original, un-canonicalized value) — see `ON_OFF_CANONICAL`/`JC_CANONICAL`'s own doc comments for what this does and does not fold together. */
function canonicalizeAttributeValue(tag: string, attribute: string, value: string): string {
  const onOff = ON_OFF_CANONICAL.get(value)
  if (onOff !== undefined) {
    return onOff
  }
  if (attribute === 'w:val' && JC_TAGS.has(tag)) {
    const jc = JC_CANONICAL.get(value)
    if (jc !== undefined) {
      return jc
    }
  }
  return value
}

/**
 * How many of `tag`'s surviving (saved) occurrences carried a value
 * canonically equal to `originalValue`'s — matching `attrName` prefix-
 * insensitively (`localAttributeName`) so a source attribute re-emitted
 * under a corrected/different namespace prefix (Atlas does this for
 * `w:comment/@w:parentId` -> `@w15:parentId`, normalizing a source
 * document's own non-conformant prefix to the schema-correct one — see
 * `parser/comments.ts`/`commentsWriter.ts`) isn't reported as the value
 * having disappeared.
 */
function survivingValueCount(
  savedAttributes: ReadonlyMap<string, Map<string, number>>,
  tag: string,
  attrName: string,
  originalValue: string,
): number {
  const targetLocalName = localAttributeName(attrName)
  const targetCanonical = canonicalizeAttributeValue(tag, attrName, originalValue)
  let total = 0
  for (const [savedName, savedValues] of savedAttributes) {
    if (localAttributeName(savedName) !== targetLocalName) {
      continue
    }
    for (const [savedValue, count] of savedValues) {
      if (canonicalizeAttributeValue(tag, attrName, savedValue) === targetCanonical) {
        total += count
      }
    }
  }
  return total
}

/**
 * Namespace-declaration and markup-compatibility plumbing: never document
 * content under any circumstances, and (see `buildDocumentAttributes`/
 * `partWriterSupport.ts`) Atlas already deliberately normalizes both —
 * unioning the source's own namespace prefixes / `mc:Ignorable` tokens with
 * a fixed baseline set — so the exact resulting string can legitimately
 * differ from the source (different token order, or extra baseline
 * prefixes the source didn't itself declare) without anything being lost.
 */
function isNamespacePlumbingAttribute(name: string): boolean {
  return name === 'mc:Ignorable' || name === 'xmlns' || name.startsWith('xmlns:')
}

function isIdentityAttribute(tag: string, attribute: string): boolean {
  return (
    IDENTITY_ATTRIBUTE_NAMES.has(attribute)
    || isNamespacePlumbingAttribute(attribute)
    || TAG_SCOPED_IDENTITY_ATTRIBUTES.has(`${tag}/@${attribute}`)
  )
}

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text' })

interface TagStats {
  count: number
  /**
   * How many of `count` occurrences had zero attributes and zero children
   * (see `isEmptyOccurrence`). For a `SAFE_WHEN_EMPTY_TAGS` member, this
   * many occurrences are provably safe to lose — whether that's ALL of
   * them (the original, full-removal-only check) or only SOME (the
   * count-drop check below, added for `a:srcRect`'s own shape: one real
   * occurrence plus one redundant empty one, only the second of which may
   * legitimately vanish).
   */
  emptyCount: number
  /**
   * Attribute (qualified) name -> value -> how many of this tag's
   * occurrences (anywhere in the part — this diff is flat/element-name-
   * keyed throughout, not positional) carried exactly that value. An
   * occurrence lacking the attribute contributes nothing here.
   */
  attributes: Map<string, Map<string, number>>
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
    const existing = into.get(key) ?? { count: 0, emptyCount: 0, attributes: new Map<string, Map<string, number>>() }
    for (const occurrence of occurrences) {
      existing.count += 1
      if (isEmptyOccurrence(occurrence)) {
        existing.emptyCount += 1
      }
      recordAttributes(existing.attributes, occurrence)
    }
    into.set(key, existing)
    walk(value, into)
  }
}

/** Records `occurrence`'s own `@_`-prefixed attributes into `into` (see `TagStats.attributes`'s doc comment) — never the attributes of its children, which `walk`'s own recursion records against their own tag name instead. */
function recordAttributes(into: Map<string, Map<string, number>>, occurrence: unknown): void {
  if (occurrence === null || typeof occurrence !== 'object' || Array.isArray(occurrence)) {
    return
  }
  for (const [key, value] of Object.entries(occurrence as Record<string, unknown>)) {
    if (!key.startsWith('@_')) {
      continue
    }
    const name = key.slice('@_'.length)
    // fast-xml-parser (no `parseAttributeValue`) already gives us the raw
    // attribute string; `String(...)` only guards the pathological case of
    // a non-string/number value slipping through.
    const valueString = String(value)
    const values = into.get(name) ?? new Map<string, number>()
    values.set(valueString, (values.get(valueString) ?? 0) + 1)
    into.set(name, values)
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
  const partsToCheck = [
    ...ALWAYS_REWRITTEN_PARTS,
    ...NOTES_AND_COMMENTS_PARTS,
    ...headerFooterPartPaths(bundle),
  ]

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
      // How many of this tag's ORIGINAL occurrences are provably safe to
      // lose because they were entirely empty (see `TagStats.emptyCount`'s
      // doc comment) — 0 unless `tag` is on the curated `SAFE_WHEN_EMPTY_TAGS`
      // allowlist, in which case it's every empty occurrence there was, not
      // just "all or nothing": that lets `a:srcRect`'s one real occurrence
      // plus one redundant empty one (count-drop 2 -> 1) clear the bar below
      // without also exempting a tag whose mere PRESENCE is its entire
      // meaning (a toggle like `w:b`) from ever being checked — such a tag
      // is simply never added to that allowlist in the first place (see its
      // own doc comment).
      const safeToLoseCount = SAFE_WHEN_EMPTY_TAGS.has(tag) ? stats.emptyCount : 0
      const meaningfulOriginalCount = stats.count - safeToLoseCount

      const savedStats = savedCounts.get(tag)
      if (savedStats === undefined) {
        if (meaningfulOriginalCount > 0) {
          warnings.push({ part: partPath, tag, kind: 'element', occurrences: stats.count })
        }
        // The element is entirely gone — nothing left to compare a
        // specific attribute VALUE against; the line above already says
        // everything there is to say about this tag in this part.
        continue
      }

      // Count-drop: fewer occurrences of a tag than the original had is
      // exactly the "12 `w:tab` became 1" shape a presence-only check can't
      // see — the tag survives, so `savedCounts.has(tag)` above is true,
      // but real occurrences of it quietly vanished. An INCREASE is not
      // reported: nothing here promises the compared save had no
      // intervening edits (the real Save-flow callsite diffs the
      // as-opened bytes against the just-saved ones across a whole
      // editing session, not a no-op save), and a user adding content is
      // indistinguishable from — and vastly more likely than — a bug that
      // manufactures new elements, so warning on growth would be pure
      // false-positive risk for no realistic true-positive benefit.
      if (savedStats.count < meaningfulOriginalCount) {
        warnings.push({
          part: partPath,
          tag,
          kind: 'element',
          occurrences: stats.count,
          savedOccurrences: savedStats.count,
        })
      }

      if (meaningfulOriginalCount === 0) {
        // Every original occurrence was a provably-safe-to-lose empty one
        // (this tag's full-removal case, generalized) — nothing it carried
        // could have had a meaningful attribute value to begin with.
        continue
      }

      // Attribute-value loss: the element survived (possibly at a smaller
      // count, reported above) but a value one of its occurrences carried
      // for `attrName` in the original no longer appears on ANY of its
      // surviving occurrences — DOCX-1's exact shape (`w:rFonts` kept its
      // literal `w:ascii`/etc. attributes, so the tag/count checks above
      // both see nothing wrong, while `w:asciiTheme`'s value silently
      // disappeared). Checking "does this exact (canonicalized) value still
      // exist somewhere on this tag" rather than positionally matching
      // occurrence-for-occurrence is deliberate: it can't tell two
      // occurrences' values being swapped from nothing happening at all
      // (a real, accepted precision-over-recall trade-off — see this
      // module's test file and this task's own report for where that line
      // was drawn), but it also means legitimate churn — re-serialization
      // visiting occurrences in a different order, or ONE occurrence's
      // value changing because an edit actually touched it while every
      // other occurrence's original value is still present elsewhere on
      // the same tag — never warns.
      for (const [attrName, originalValues] of stats.attributes) {
        if (isIdentityAttribute(tag, attrName)) {
          continue
        }
        for (const [value, count] of originalValues) {
          const survivingCount = survivingValueCount(savedStats.attributes, tag, attrName, value)
          if (survivingCount === 0) {
            warnings.push({
              part: partPath,
              tag,
              kind: 'attribute',
              attribute: attrName,
              occurrences: count,
            })
          }
        }
      }
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

function occurrencePhrase(count: number): string {
  return `${count} occurrence${count === 1 ? '' : 's'}`
}

/** Renders warnings as short, developer-facing lines, one per (part, tag [, attribute]). Not for end-user display — see `categorizeLossySaveWarnings` for that. */
export function describeLossySaveWarnings(warnings: ReadonlyArray<LossySaveWarning>): ReadonlyArray<string> {
  return warnings.map((warning) => {
    if (warning.kind === 'attribute') {
      return `${warning.part}: ${warning.tag}/@${warning.attribute} (${occurrencePhrase(warning.occurrences)}) lost its value and will not be preserved`
    }
    if (warning.savedOccurrences !== undefined) {
      return `${warning.part}: ${warning.tag} (${occurrencePhrase(warning.occurrences)} dropped to ${warning.savedOccurrences}) is not fully supported and some will be removed`
    }
    return `${warning.part}: ${warning.tag} (${occurrencePhrase(warning.occurrences)}) is not supported and will be removed`
  })
}

// ---------------------------------------------------------------------------
// End-user-facing categorization (Save flow — see `DocxViewer.tsx`)
//
// A non-technical user has no use for "w:sdt" or "mc:AlternateContent" —
// they know their document has a dropdown/date-picker field, or a shape/
// text box, not an XML element. This groups a warning's raw tag into the
// coarse category the Save flow has plain-language copy for, so the UI
// layer never has to name an element itself.
// ---------------------------------------------------------------------------

export type LossySaveWarningCategory = 'content-control' | 'shape-fallback' | 'other'

/**
 * Every element name a `w:sdt` (content control) wrapper or its properties
 * can introduce — i.e. exactly what disappears when `documentWriter.ts`'s
 * wrapper-region passthrough (see `WrapperPassthrough`'s doc comment on
 * `../model/document.ts`) doesn't fire for it: an edit touched the
 * control's own content, or it sits somewhere the parser doesn't capture a
 * region for at all (a table cell, header, or footer).
 */
const CONTENT_CONTROL_TAGS: ReadonlySet<string> = new Set([
  'w:sdt',
  'w:sdtPr',
  'w:sdtEndPr',
  'w:sdtContent',
  'w:id',
  'w:alias',
  'w:tag',
  'w:lock',
  'w:placeholder',
  'w:showingPlcHdr',
  'w:dataBinding',
  'w:text',
  'w:comboBox',
  'w:dropDownList',
  'w:date',
  'w:docPartObj',
  'w:docPartList',
  'w:citation',
  'w:group',
  'w:picture',
  'w:checkbox',
  'w15:color',
  'w15:appearance',
])

/** Every element name an `mc:AlternateContent` wrapper (a shape/text box's modern-vs-legacy-VML pair) can introduce, for the same reason as `CONTENT_CONTROL_TAGS`. */
const SHAPE_FALLBACK_TAGS: ReadonlySet<string> = new Set([
  'mc:AlternateContent',
  'mc:Choice',
  'mc:Fallback',
  'w:pict',
  'v:shape',
  'v:shapetype',
  'v:rect',
  'v:roundrect',
  'v:oval',
  'v:line',
  'v:textbox',
  'v:imagedata',
])

function categorizeLossySaveWarningTag(tag: string): LossySaveWarningCategory {
  if (CONTENT_CONTROL_TAGS.has(tag)) {
    return 'content-control'
  }
  if (SHAPE_FALLBACK_TAGS.has(tag)) {
    return 'shape-fallback'
  }
  return 'other'
}

/**
 * Groups `warnings` into the coarse, user-facing categories the Save flow
 * shows plain-language copy for, instead of naming XML elements. An empty
 * result means nothing worth telling the user was lost.
 */
export function categorizeLossySaveWarnings(
  warnings: ReadonlyArray<LossySaveWarning>,
): ReadonlySet<LossySaveWarningCategory> {
  return new Set(warnings.map((warning) => categorizeLossySaveWarningTag(warning.tag)))
}
