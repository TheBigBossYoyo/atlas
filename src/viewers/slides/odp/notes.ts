/** S12 — speaker notes: a `draw:page`'s `presentation:notes` child, minus its footer/page-number placeholders. */

import { getElementsByLocalName, getFirstByLocalName } from '../shared/xmlUtils'
import type { OdpStyleIndex } from './styles'
import { parseOdpTextBody } from './text'

const SKIPPED_PRESENTATION_CLASSES = new Set(['page-number', 'date-time', 'header', 'footer'])

/** Resolves a slide's notes page text, or `undefined` when it has none / is empty. */
export function resolveOdpNotes(page: Element, index: OdpStyleIndex): string | undefined {
  const notesPage = getFirstByLocalName(page, 'notes')
  if (!notesPage) {
    return undefined
  }

  const sections: string[] = []

  for (const frame of getElementsByLocalName(notesPage, 'frame')) {
    const presentationClass = frame.getAttribute('presentation:class')
    if (presentationClass && SKIPPED_PRESENTATION_CLASSES.has(presentationClass)) {
      continue
    }

    const textBox = getFirstByLocalName(frame, 'text-box')
    if (!textBox) {
      continue
    }

    const { text } = parseOdpTextBody(textBox, index)
    if (text) {
      sections.push(text)
    }
  }

  const combined = sections.join('\n\n').trim()
  return combined.length > 0 ? combined : undefined
}
