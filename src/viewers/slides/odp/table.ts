/** S6 — ODP table rendering: `table:table` -> a 2D grid of cells with resolved run formatting. */

import { getDirectChildren } from '../shared/xmlUtils'
import type { SlideTableCell } from '../../shared/SlideDeck.types'
import type { OdpStyleIndex } from './styles'
import { parseOdpTextBody } from './text'

function parseCell(cell: Element, index: OdpStyleIndex): SlideTableCell {
  const parsed = parseOdpTextBody(cell, index)
  return { text: parsed.text, runs: parsed.paragraphs.flatMap(paragraph => paragraph.runs) }
}

/** Walks `table:table-row`/`table:table-cell` into a row-major grid. */
export function parseOdpTable(table: Element, index: OdpStyleIndex): ReadonlyArray<ReadonlyArray<SlideTableCell>> {
  return getDirectChildren(table, 'table-row').map(row =>
    getDirectChildren(row, 'table-cell').map(cell => parseCell(cell, index)),
  )
}
