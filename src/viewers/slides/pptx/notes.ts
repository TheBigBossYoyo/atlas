/** S12 — speaker notes: a slide's `notesSlide` relationship, minus its footer/number/thumbnail placeholders. */

import type { CancelSignal, ZipArchive } from '../shared/xmlUtils'
import { getDirectChildren, getElementsByLocalName, getFirstByLocalName, parseXml, readZipText } from '../shared/xmlUtils'
import { findRelationshipTarget, parseRelationships } from './relationships'

const SKIPPED_PLACEHOLDER_TYPES = new Set(['sldImg', 'sldNum', 'dt', 'ftr'])

/** Mirrors S5's document-order walk so a manual line break in notes text becomes a real `\n`. */
function extractParagraphText(paragraph: Element): string {
  let text = ''

  for (const child of getDirectChildren(paragraph)) {
    if (child.localName === 'r' || child.localName === 'fld') {
      text += getFirstByLocalName(child, 't')?.textContent ?? ''
    } else if (child.localName === 'br') {
      text += '\n'
    }
  }

  return text
}

function extractShapeNotesText(shape: Element): string {
  const txBody = getFirstByLocalName(shape, 'txBody')
  if (!txBody) {
    return ''
  }

  return getDirectChildren(txBody, 'p')
    .map(extractParagraphText)
    .join('\n')
    .trim()
}

/** Resolves and parses a slide's notes page, returning trimmed body text or `undefined` when there is none. */
export async function resolveSlideNotes(
  zip: ZipArchive,
  slidePath: string,
  signal: CancelSignal,
): Promise<string | undefined> {
  const slideRels = await parseRelationships(zip, slidePath, signal)
  const notesPath = findRelationshipTarget(slideRels, '/notesSlide')
  if (!notesPath || signal.cancelled) {
    return undefined
  }

  const xml = await readZipText(zip, notesPath, signal)
  if (!xml) {
    return undefined
  }

  const notesDocument = parseXml(xml)
  const sections: string[] = []

  for (const shape of getElementsByLocalName(notesDocument, 'sp')) {
    const ph = getFirstByLocalName(shape, 'ph')
    const type = ph?.getAttribute('type')
    if (type && SKIPPED_PLACEHOLDER_TYPES.has(type)) {
      continue
    }

    const text = extractShapeNotesText(shape)
    if (text) {
      sections.push(text)
    }
  }

  const combined = sections.join('\n\n').trim()
  return combined.length > 0 ? combined : undefined
}
