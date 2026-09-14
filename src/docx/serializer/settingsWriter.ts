/**
 * Atlas — `word/settings.xml` writer (D17 / DXE-11)
 *
 * `word/settings.xml` passes through the save pipeline byte-for-byte
 * (`docx/index.ts`'s `rawArchive` passthrough) for everything except the one
 * setting the editor's Track Changes toggle owns. Rather than parsing the
 * whole part into a generic tree and rebuilding it — risking dropping or
 * reordering settings Atlas doesn't model and Word is known to be less
 * forgiving about in this specific part than in most others — this performs
 * a narrow, targeted string edit that inserts, normalizes, or removes just
 * the `<w:trackChanges/>` element and leaves every other byte untouched.
 */
import { WORD_NAMESPACE, XML_DECLARATION } from './partWriterSupport'

const TRACK_CHANGES_ELEMENT_RE =
  /<w:trackChanges\b[^>]*\/>|<w:trackChanges\b[^>]*>[\s\S]*?<\/w:trackChanges>/
const SETTINGS_OPEN_TAG_RE = /<w:settings\b[^>]*>/

/**
 * A conservative subset of the `CT_Settings` element sequence (ECMA-376
 * §17.15.1) that reliably precedes `w:trackChanges`: document-level display
 * and proofing flags Word writes, when present, near the top of the file.
 * Used only to find a schema-order-safe insertion point for a brand-new
 * `<w:trackChanges/>` — always inserting as the very first child (this
 * module's original approach) is only schema-valid when the source document
 * has none of these earlier elements, which holds for small hand-built
 * fixtures but not for real Word output: inserting `<w:trackChanges/>`
 * before an already-present `<w:zoom/>`, for example, produces an
 * out-of-sequence `settings.xml` a strict OOXML consumer may flag for
 * repair. Intentionally NOT exhaustive — only entries whose relative
 * position is unambiguous are listed; anything else falls through to the
 * open-tag fallback below, same as before this list existed.
 */
const ELEMENTS_BEFORE_TRACK_CHANGES: ReadonlyArray<string> = [
  'writeProtection',
  'view',
  'zoom',
  'removePersonalInformation',
  'doNotDisplayPageBoundaries',
  'displayBackgroundShape',
  'embedTrueTypeFonts',
  'embedSystemFonts',
  'saveSubsetFonts',
  'mirrorMargins',
  'hideSpellingErrors',
  'hideGrammaticalErrors',
  'proofState',
  'attachedTemplate',
  'linkStyles',
  'documentType',
  'mailMerge',
  'revisionView',
]

/**
 * The end offset of the last (rightmost) occurrence, at or after
 * `searchStart`, of any element in `ELEMENTS_BEFORE_TRACK_CHANGES` — or
 * `null` if the document has none of them. All of these elements are
 * schema-defined as attribute-only/empty (no element children), so a
 * non-nesting self-closing-or-simple-content match is safe here.
 */
function findLastKnownPrecedingElementEnd(xml: string, searchStart: number): number | null {
  let latestEnd: number | null = null

  for (const tag of ELEMENTS_BEFORE_TRACK_CHANGES) {
    const pattern = new RegExp(`<w:${tag}\\b[^>]*/>|<w:${tag}\\b[^>]*>[\\s\\S]*?<\\/w:${tag}>`, 'g')
    pattern.lastIndex = searchStart
    let match = pattern.exec(xml)
    while (match !== null) {
      const end = match.index + match[0].length
      if (latestEnd === null || end > latestEnd) {
        latestEnd = end
      }
      match = pattern.exec(xml)
    }
  }

  return latestEnd
}

/**
 * Returns `originalXml` patched so it reflects `trackChanges`, or `undefined`
 * when there is nothing to write (no original part, and Track Changes is
 * off — the default state a document with no settings.xml already implies).
 *
 * @param originalXml - The source document's `word/settings.xml` text, or
 *   `undefined` when the document never had one.
 * @param trackChanges - The editor's current Track Changes toggle state.
 */
export function writeSettingsXml(
  originalXml: string | undefined,
  trackChanges: boolean,
): string | undefined {
  if (originalXml === undefined) {
    return trackChanges
      ? `${XML_DECLARATION}<w:settings xmlns:w="${WORD_NAMESPACE}"><w:trackChanges/></w:settings>`
      : undefined
  }

  const hasElement = TRACK_CHANGES_ELEMENT_RE.test(originalXml)

  if (!trackChanges) {
    return hasElement ? originalXml.replace(TRACK_CHANGES_ELEMENT_RE, '') : originalXml
  }

  if (hasElement) {
    // Normalizes a `w:val="false"`/`"0"` form to bare presence; a no-op
    // (returns byte-identical text) when it was already bare.
    return originalXml.replace(TRACK_CHANGES_ELEMENT_RE, '<w:trackChanges/>')
  }

  const openTag = SETTINGS_OPEN_TAG_RE.exec(originalXml)
  if (openTag === null) {
    // Doesn't look like a well-formed <w:settings> root — leave it exactly
    // as loaded rather than risk corrupting an unexpected shape.
    return originalXml
  }

  const afterOpenTag = openTag.index + openTag[0].length
  const insertAt = findLastKnownPrecedingElementEnd(originalXml, afterOpenTag) ?? afterOpenTag
  return `${originalXml.slice(0, insertAt)}<w:trackChanges/>${originalXml.slice(insertAt)}`
}
