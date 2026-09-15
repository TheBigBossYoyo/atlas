/**
 * Atlas — minimal `docProps/core.xml` reader for AUTHOR/TITLE fields
 * (DEFER-5 / DXS-20)
 *
 * `docPropsWriter.ts` treats `docProps/core.xml` as an opaque passthrough
 * string it only ever regenerates two elements of (`dcterms:modified`/
 * `cp:lastModifiedBy`, DXS-13) — there's no structured model for the rest
 * of it (by design; see that file's doc comment). AUTHOR/TITLE field
 * evaluation needs `dc:creator`/`dc:title` as plain strings, which doesn't
 * justify a real parser: a plain extraction regex is enough, and keeps
 * this module's only DOCX-package dependency a raw string the caller
 * already has (`bundle.rawArchive.get('docProps/core.xml')`).
 */
export interface CoreProps {
  readonly author?: string
  readonly title?: string
}

const CREATOR_RE = /<dc:creator>([\s\S]*?)<\/dc:creator>/
const TITLE_RE = /<dc:title>([\s\S]*?)<\/dc:title>/

export function parseCoreProps(xml: string | undefined): CoreProps {
  if (xml === undefined) {
    return {}
  }

  const author = CREATOR_RE.exec(xml)?.[1]
  const title = TITLE_RE.exec(xml)?.[1]

  return {
    ...(author !== undefined ? { author: decodeXmlEntities(author) } : {}),
    ...(title !== undefined ? { title: decodeXmlEntities(title) } : {}),
  }
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}
