/**
 * Atlas — plain-text extraction for a paragraph (DEFER-5 / DXS-20)
 *
 * A small, self-contained "what would a reader see as this paragraph's
 * text" helper for `toc.ts`'s heading-entry text — deliberately not
 * reusing `editor/Find.ts`'s internal `paragraphText` (built around that
 * module's own `RunSlice` search-index abstraction) to keep this module's
 * dependencies to just the model.
 */
import type { HyperlinkChild, Paragraph, ParagraphChild, Run, RunChild } from '../model'

export function paragraphPlainText(paragraph: Paragraph): string {
  return paragraph.children.map(paragraphChildText).join('')
}

function paragraphChildText(child: ParagraphChild): string {
  switch (child.kind) {
    case 'run':
      return runText(child)
    case 'hyperlink':
      return child.children.map(hyperlinkChildText).join('')
    case 'field':
      // A field's cached display content IS its visible text.
      return child.result.map(paragraphChildText).join('')
    case 'ins-revision':
      return (child.children as ReadonlyArray<ParagraphChild>).map(paragraphChildText).join('')
    case 'del-revision':
      // Tracked-deleted text isn't part of the document's visible content.
      return ''
    default:
      return ''
  }
}

function hyperlinkChildText(child: HyperlinkChild): string {
  switch (child.kind) {
    case 'run':
      return runText(child)
    case 'field':
      return child.result.map(paragraphChildText).join('')
    default:
      return ''
  }
}

function runText(run: Run): string {
  return run.children.map(runChildText).join('')
}

function runChildText(child: RunChild): string {
  switch (child.kind) {
    case 'text':
      return child.value
    case 'tab':
      return '\t'
    default:
      return ''
  }
}
