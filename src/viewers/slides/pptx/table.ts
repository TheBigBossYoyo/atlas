/** S6 — table rendering: `a:tbl` -> a 2D grid of cells with their own resolved run formatting. */

import { getDirectChildren, getElementsByLocalName, getFirstByLocalName } from '../shared/xmlUtils'
import type { SlideTableCell } from '../../shared/SlideDeck.types'
import type { LayoutChain } from './layout'
import { parseTextBody } from './text'

const TABLE_GRAPHIC_DATA_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'

/** The `a:tbl` inside a `p:graphicFrame`, if its `a:graphicData@uri` marks it as a table. */
export function findTableElement(graphicFrame: Element): Element | null {
  const graphic = getFirstByLocalName(graphicFrame, 'graphic')
  const graphicData = graphic ? getFirstByLocalName(graphic, 'graphicData') : null
  if (graphicData?.getAttribute('uri') !== TABLE_GRAPHIC_DATA_URI) {
    return null
  }

  return getFirstByLocalName(graphicData, 'tbl')
}

function parseCell(tc: Element, chain: LayoutChain): SlideTableCell {
  const parsed = parseTextBody(tc, null, chain)
  return { text: parsed.text, runs: parsed.paragraphs.flatMap(paragraph => paragraph.runs) }
}

/** Walks `a:tr`/`a:tc` into a row-major grid, resolving each cell's text the same way a text box would. */
export function parseTable(tbl: Element, chain: LayoutChain): ReadonlyArray<ReadonlyArray<SlideTableCell>> {
  return getDirectChildren(tbl, 'tr').map(tr =>
    getDirectChildren(tr, 'tc').map(tc => parseCell(tc, chain)),
  )
}

/** Marks every `<a:t>` inside a table as consumed, so the stray-text fallback doesn't double-render cells. */
export function markTableTextNodesConsumed(tbl: Element, consumed: Set<Element>): void {
  for (const textNode of getElementsByLocalName(tbl, 't')) {
    consumed.add(textNode)
  }
}
