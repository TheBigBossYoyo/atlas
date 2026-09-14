/** S6/S11 — classifies a `p:graphicFrame` (table / chart / SmartArt / OLE) and builds its `SlideShape`. */

import { getFirstByLocalName } from '../shared/xmlUtils'
import type { SlideShape, SlideTransform } from '../../shared/SlideDeck.types'
import type { LayoutChain } from './layout'
import { findTableElement, markTableTextNodesConsumed, parseTable } from './table'

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const OLE_URI = 'http://schemas.openxmlformats.org/presentationml/2006/ole'

function graphicDataUri(graphicFrame: Element): string | null {
  const graphic = getFirstByLocalName(graphicFrame, 'graphic')
  const graphicData = graphic ? getFirstByLocalName(graphic, 'graphicData') : null
  return graphicData?.getAttribute('uri') ?? null
}

/**
 * Builds the `SlideShape` for a `p:graphicFrame`: a real table grid (S6), a
 * labeled placeholder for chart/SmartArt/OLE content Atlas cannot render
 * (S11), or `null` for an unrecognized/empty frame.
 */
export function buildGraphicFrameShape(
  graphicFrame: Element,
  id: string,
  transform: SlideTransform,
  chain: LayoutChain,
  consumedTextNodes: Set<Element>,
): SlideShape | null {
  const tbl = findTableElement(graphicFrame)
  if (tbl) {
    markTableTextNodesConsumed(tbl, consumedTextNodes)
    return { kind: 'table', id, transform, rows: parseTable(tbl, chain) }
  }

  const uri = graphicDataUri(graphicFrame)
  if (uri === CHART_URI) {
    return { kind: 'unsupported', id, transform, label: 'Chart not supported' }
  }

  if (uri === DIAGRAM_URI) {
    return { kind: 'unsupported', id, transform, label: 'SmartArt not supported' }
  }

  if (uri === OLE_URI) {
    return { kind: 'unsupported', id, transform, label: 'Embedded object not supported' }
  }

  return null
}
