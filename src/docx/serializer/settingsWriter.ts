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

  const insertAt = openTag.index + openTag[0].length
  return `${originalXml.slice(0, insertAt)}<w:trackChanges/>${originalXml.slice(insertAt)}`
}
