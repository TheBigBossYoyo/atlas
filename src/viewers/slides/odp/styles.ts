/**
 * S2 — ODP master pages / `styles.xml` resolution, plus the generic named-style
 * lookup (`style:style@style:parent-style-name` chains) shared by text (S3/S4),
 * shape fill/border (S8), and page background resolution.
 */

import { getDirectChildren, getElementsByLocalName, getFirstByLocalName, parseXml } from '../shared/xmlUtils'
import { parseOdfLengthToPixels } from '../shared/units'

const DEFAULT_SLIDE_WIDTH = 1280
const DEFAULT_SLIDE_HEIGHT = 720

export type OdpDocuments = {
  readonly contentDocument: XMLDocument
  /** `null` when the archive has no `styles.xml` (some minimal/legacy exports). */
  readonly stylesDocument: XMLDocument | null
}

export async function loadOdpDocuments(
  readText: (path: string) => Promise<string | null>,
): Promise<OdpDocuments> {
  const contentXml = await readText('content.xml')
  if (!contentXml) {
    throw new Error('Missing archive entry: content.xml')
  }

  const stylesXml = await readText('styles.xml')

  return {
    contentDocument: parseXml(contentXml),
    stylesDocument: stylesXml ? parseXml(stylesXml) : null,
  }
}

/**
 * Every named `<style:style>` in a document's `automatic-styles`/`styles`, keyed by name.
 * `<style:default-style>` (family-wide defaults with no name) is out of scope here.
 */
function collectNamedStyles(root: XMLDocument | Element | null, target: Map<string, Element>): void {
  if (!root) {
    return
  }

  for (const style of getElementsByLocalName(root, 'style')) {
    const name = style.getAttribute('style:name') ?? style.getAttribute('name')
    if (name && !target.has(name)) {
      target.set(name, style)
    }
  }
}

export type OdpStyleIndex = {
  readonly byName: ReadonlyMap<string, Element>
  readonly listStyles: ReadonlyMap<string, Element>
  readonly masterPages: ReadonlyMap<string, Element>
  readonly pageLayouts: ReadonlyMap<string, Element>
}

/** Builds one style lookup spanning both `content.xml` and `styles.xml` (producers split styles across either). */
export function buildStyleIndex(docs: OdpDocuments): OdpStyleIndex {
  const byName = new Map<string, Element>()
  collectNamedStyles(docs.contentDocument, byName)
  collectNamedStyles(docs.stylesDocument, byName)

  const listStyles = new Map<string, Element>()
  for (const root of [docs.contentDocument, docs.stylesDocument]) {
    if (!root) continue
    for (const listStyle of getElementsByLocalName(root, 'list-style')) {
      const name = listStyle.getAttribute('style:name')
      if (name && !listStyles.has(name)) {
        listStyles.set(name, listStyle)
      }
    }
  }

  const masterPages = new Map<string, Element>()
  const pageLayouts = new Map<string, Element>()
  for (const root of [docs.stylesDocument, docs.contentDocument]) {
    if (!root) continue
    for (const masterPage of getElementsByLocalName(root, 'master-page')) {
      const name = masterPage.getAttribute('style:name')
      if (name && !masterPages.has(name)) {
        masterPages.set(name, masterPage)
      }
    }
    for (const layout of getElementsByLocalName(root, 'page-layout')) {
      const name = layout.getAttribute('style:name')
      if (name && !pageLayouts.has(name)) {
        pageLayouts.set(name, layout)
      }
    }
  }

  return { byName, listStyles, masterPages, pageLayouts }
}

/** Walks `style:parent-style-name` from `styleName` up to its root ancestor, self first. */
export function resolveStyleChain(styleName: string | null, index: OdpStyleIndex): Element[] {
  const chain: Element[] = []
  let current = styleName
  const seen = new Set<string>()

  while (current && !seen.has(current)) {
    seen.add(current)
    const style = index.byName.get(current)
    if (!style) {
      break
    }

    chain.push(style)
    current = style.getAttribute('style:parent-style-name')
  }

  return chain
}

/** The first `<style:PROPERTY-NAME>` child found while walking a style's parent chain. */
export function findInheritedProperties(chain: ReadonlyArray<Element>, propertyTag: string): Element | null {
  for (const style of chain) {
    const properties = getFirstByLocalName(style, propertyTag)
    if (properties) {
      return properties
    }
  }

  return null
}

export type SlidePageSize = { readonly width: number; readonly height: number }

function readPageLayoutSize(layout: Element | null): SlidePageSize | null {
  const properties = layout ? getFirstByLocalName(layout, 'page-layout-properties') : null
  if (!properties) {
    return null
  }

  const width = parseOdfLengthToPixels(properties.getAttribute('fo:page-width'))
  const height = parseOdfLengthToPixels(properties.getAttribute('fo:page-height'))
  return width !== undefined && height !== undefined ? { width, height } : null
}

/**
 * Resolves the presentation's slide size from the first page's master page's
 * page layout (`styles.xml`), falling back to `content.xml`'s own automatic
 * `page-layout-properties` only when no master/layout chain is found.
 */
export function resolveSlideSize(
  firstPage: Element | null,
  index: OdpStyleIndex,
  contentDocument: XMLDocument,
): SlidePageSize {
  const masterPageName = firstPage?.getAttribute('draw:master-page-name')
  const masterPage = masterPageName ? index.masterPages.get(masterPageName) : undefined
  const layoutName = masterPage?.getAttribute('style:page-layout-name')
  const layout = layoutName ? index.pageLayouts.get(layoutName) : undefined

  return (
    readPageLayoutSize(layout ?? null)
    ?? readPageLayoutSize(getFirstByLocalName(contentDocument, 'page-layout'))
    ?? { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT }
  )
}

/** A page's own master page element, resolved via `draw:master-page-name` — used for background/title inheritance. */
export function resolveMasterPage(page: Element, index: OdpStyleIndex): Element | null {
  const name = page.getAttribute('draw:master-page-name')
  return name ? index.masterPages.get(name) ?? null : null
}

/** Direct `draw:frame`/other drawing shapes on a master page (S1/S8 background + placeholder inheritance). */
export function masterPageShapes(masterPage: Element | null): Element[] {
  return masterPage ? getDirectChildren(masterPage) : []
}
