/** S4 — bullet/numbering resolution (`a:buChar`/`a:buAutoNum`/`a:buNone`), with the same
 * slide -> layout -> master -> txStyles inheritance chain used for run formatting. */

import { getDirectChildren, getFirstByLocalName } from '../shared/xmlUtils'
import type { SlideBullet } from '../../shared/SlideDeck.types'
import { findMatchingPlaceholder, type LayoutChain, type PlaceholderRef } from './layout'

type BulletSource = { readonly char?: string; readonly numbered?: boolean; readonly none: boolean }

/** Reads a paragraph-properties-shaped element's own `a:buChar`/`a:buAutoNum`/`a:buNone`, if any. */
function readOwnBulletSource(pPr: Element | null): BulletSource | null {
  if (!pPr) {
    return null
  }

  if (getFirstByLocalName(pPr, 'buNone')) {
    return { none: true }
  }

  const buChar = getFirstByLocalName(pPr, 'buChar')
  if (buChar) {
    const char = buChar.getAttribute('char')
    return char ? { char, none: false } : null
  }

  if (getFirstByLocalName(pPr, 'buAutoNum')) {
    return { numbered: true, none: false }
  }

  return null
}

function findLvlPPr(container: Element, level: number, tag: 'lstStyle' | 'title' | 'body' | 'other'): Element | null {
  const root = tag === 'lstStyle'
    ? getFirstByLocalName(getFirstByLocalName(container, 'txBody') ?? container, 'lstStyle')
    : container

  if (!root) {
    return null
  }

  return getDirectChildren(root).find(child => child.localName === `lvl${level + 1}pPr`) ?? null
}

function txStylesLvlPPr(
  masterDocument: XMLDocument | null,
  family: 'titleStyle' | 'bodyStyle' | 'otherStyle',
  level: number,
): Element | null {
  if (!masterDocument) {
    return null
  }

  const txStyles = getFirstByLocalName(masterDocument, 'txStyles')
  const styleRoot = txStyles ? getFirstByLocalName(txStyles, family) : null
  if (!styleRoot) {
    return null
  }

  return getDirectChildren(styleRoot).find(child => child.localName === `lvl${level + 1}pPr`) ?? null
}

function familyFor(ref: PlaceholderRef | null): 'titleStyle' | 'bodyStyle' | 'otherStyle' {
  if (!ref) {
    return 'otherStyle'
  }

  if (ref.type === 'title' || ref.type === 'ctrTitle') {
    return 'titleStyle'
  }

  if (ref.type === 'body' || ref.type === 'subTitle' || ref.type === null) {
    return 'bodyStyle'
  }

  return 'otherStyle'
}

/**
 * Resolves a paragraph's bullet, walking the paragraph's own `a:pPr`, then its
 * shape's `a:lstStyle` level, then the matching layout/master placeholder's
 * level, then the master's `p:txStyles` level. `a:buNone` at any point stops
 * inheritance immediately (a paragraph explicitly opted out of bullets).
 */
export function resolveBullet(
  paragraphPPr: Element | null,
  level: number,
  ownShape: Element,
  ref: PlaceholderRef | null,
  chain: LayoutChain,
): SlideBullet | undefined {
  const sources: Array<BulletSource | null> = [readOwnBulletSource(paragraphPPr)]

  const ownLvl = findLvlPPr(ownShape, level, 'lstStyle')
  sources.push(readOwnBulletSource(ownLvl))

  const layoutMatch = ref ? findMatchingPlaceholder(chain.layoutDocument, ref) : null
  if (layoutMatch) {
    sources.push(readOwnBulletSource(findLvlPPr(layoutMatch, level, 'lstStyle')))
  }

  const masterMatch = ref ? findMatchingPlaceholder(chain.masterDocument, ref) : null
  if (masterMatch) {
    sources.push(readOwnBulletSource(findLvlPPr(masterMatch, level, 'lstStyle')))
  }

  sources.push(readOwnBulletSource(txStylesLvlPPr(chain.masterDocument, familyFor(ref), level)))

  const resolved = sources.find(source => source !== null) ?? null
  if (!resolved || resolved.none) {
    return undefined
  }

  return { level, char: resolved.char, numbered: resolved.numbered }
}
