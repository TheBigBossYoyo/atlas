/**
 * Atlas — docProps/core.xml writer (D19 / DXS-13)
 *
 * `docProps/core.xml`'s `dcterms:modified`/`cp:lastModifiedBy` previously
 * passed straight through `saveDocx`'s raw-archive passthrough untouched —
 * a document saved by Atlas kept whatever modified date/author its
 * *original* author's copy of Word last wrote, which is wrong (and, for
 * `dcterms:modified` specifically, actively misleading: it's the field
 * Explorer/Word "last modified" columns and search indexers read).
 *
 * Deliberately does a targeted string-level rewrite of just these two
 * elements rather than parsing the whole part into an object model and
 * rebuilding it: `docProps/core.xml` can carry properties Atlas has no
 * model for at all (`dc:title`, `dc:subject`, `dc:creator`, `cp:keywords`,
 * `dc:description`, `cp:category`, `cp:contentStatus`, `cp:version`,
 * `dcterms:created`, custom `cp:` extensions, …) — a parse-then-rebuild
 * would need to either model all of them or silently drop whichever ones
 * it doesn't, whereas a narrow find-and-replace leaves every other
 * property byte-for-byte untouched by construction.
 */

const CORE_PROPERTIES_NAMESPACES =
  'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
  + 'xmlns:dc="http://purl.org/dc/elements/1.1/" '
  + 'xmlns:dcterms="http://purl.org/dc/terms/" '
  + 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
  + 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

/**
 * Formats a `Date` as W3CDTF (`YYYY-MM-DDTHH:MM:SSZ`) — the profile
 * `dcterms:created`/`dcterms:modified` use, matching what Word itself
 * writes (whole seconds, no milliseconds).
 */
function toW3cdtf(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

/**
 * Rewrites `docProps/core.xml`'s `dcterms:modified` (always) and
 * `cp:lastModifiedBy` (when `lastModifiedBy` is given) to reflect this
 * save, leaving every other property untouched. Builds a minimal valid
 * `docProps/core.xml` from scratch when `existingXml` is `undefined` (a
 * source package with no `docProps/core.xml` part at all — legal per the
 * OPC spec, if unusual for a real Word-authored file).
 */
export function updateCorePropsXml(
  existingXml: string | undefined,
  lastModifiedBy: string | undefined,
  now: Date = new Date(),
): string {
  const base = existingXml ?? buildDefaultCorePropsXml()
  const withModified = replaceOrInsertElement(base, 'dcterms:modified', toW3cdtf(now), {
    'xsi:type': 'dcterms:W3CDTF',
  })

  return lastModifiedBy === undefined
    ? withModified
    : replaceOrInsertElement(withModified, 'cp:lastModifiedBy', lastModifiedBy)
}

function buildDefaultCorePropsXml(): string {
  return `${XML_DECLARATION}<cp:coreProperties ${CORE_PROPERTIES_NAMESPACES}></cp:coreProperties>`
}

/**
 * Replaces `<tagName ...>text</tagName>` (or a self-closing
 * `<tagName .../>`) with a freshly-built element carrying `textContent`
 * and `attributes`, or inserts one just before the document's closing root
 * tag when `tagName` isn't present at all.
 */
function replaceOrInsertElement(
  xml: string,
  tagName: string,
  textContent: string,
  attributes: Readonly<Record<string, string>> = {},
): string {
  const attributeString = Object.entries(attributes)
    .map(([name, value]) => ` ${name}="${escapeXmlAttribute(value)}"`)
    .join('')
  const newElement = `<${tagName}${attributeString}>${escapeXmlText(textContent)}</${tagName}>`

  const escapedTag = escapeRegExp(tagName)
  const pairedPattern = new RegExp(`<${escapedTag}\\b[^>]*>[\\s\\S]*?</${escapedTag}>`)
  if (pairedPattern.test(xml)) {
    return xml.replace(pairedPattern, newElement)
  }

  const selfClosingPattern = new RegExp(`<${escapedTag}\\b[^>]*/>`)
  if (selfClosingPattern.test(xml)) {
    return xml.replace(selfClosingPattern, newElement)
  }

  const closingRootTag = xml.match(/<\/[\w:.-]+>\s*$/)
  if (closingRootTag === null || closingRootTag.index === undefined) {
    // Defensive fallback for a malformed/unexpected shape — appending keeps
    // the value present rather than silently dropping it.
    return `${xml}${newElement}`
  }

  return `${xml.slice(0, closingRootTag.index)}${newElement}${xml.slice(closingRootTag.index)}`
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, '&quot;')
}
