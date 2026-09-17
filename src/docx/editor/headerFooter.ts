/**
 * D29 — editing a document's headers and footers.
 *
 * Headers and footers live in their own parts, outside the body the caret
 * moves through, and the save path already writes them back from the model
 * (`writeHeaderXml`/`writeFooterXml`). What was missing was any way to change
 * them, so a page number or a title typed in Word could be read but never
 * corrected here.
 *
 * Editing is per part and plain-text: each line becomes one paragraph, and
 * each paragraph keeps the properties (and its first run's formatting) of the
 * paragraph that was there before, so an edited header still looks like the
 * document's own header. Anything richer — tables, fields, images inside a
 * header — is left untouched only if the user does not edit that part; this
 * is a deliberate, documented limit rather than a silent one (the UI says
 * "text only").
 */
import type { Block, Document, Footer, Header, Paragraph, Run } from '../model'

export type HeaderFooterKind = 'header' | 'footer'

export type HeaderFooterPart = {
  readonly kind: HeaderFooterKind
  readonly id: string
  /** `default`, `first` or `even`, as referenced by the section that uses it. */
  readonly type: string
  /** Current text, one line per paragraph. */
  readonly text: string
  /** True when the part holds something this plain-text editor would drop (a table, an image, a field). */
  readonly hasRichContent: boolean
}

function paragraphText(paragraph: Paragraph): string {
  let text = ''
  for (const child of paragraph.children) {
    if (child.kind === 'run') {
      for (const item of child.children) {
        if (item.kind === 'text') text += item.value
        else if (item.kind === 'tab') text += '\t'
        else if (item.kind === 'break') text += '\n'
      }
    } else if (child.kind === 'hyperlink') {
      for (const item of child.children) {
        if (item.kind !== 'run') continue
        for (const runChild of item.children) {
          if (runChild.kind === 'text') text += runChild.value
        }
      }
    }
  }
  return text
}

export function blocksToText(blocks: ReadonlyArray<Block>): string {
  return blocks
    .filter((block): block is Paragraph => block.kind === 'paragraph')
    .map(paragraphText)
    .join('\n')
}

function hasRichContent(blocks: ReadonlyArray<Block>): boolean {
  return blocks.some((block) => {
    if (block.kind !== 'paragraph') return true
    return block.children.some(
      (child) => child.kind !== 'run' || child.children.some((item) => item.kind === 'unknown'),
    )
  })
}

/** Every header/footer part the document's sections actually reference, in reading order, without duplicates. */
export function listHeaderFooterParts(document: Document): ReadonlyArray<HeaderFooterPart> {
  const parts: HeaderFooterPart[] = []
  const seen = new Set<string>()

  for (const section of document.sections) {
    const references: ReadonlyArray<{ kind: HeaderFooterKind; id: string; type: string }> = [
      ...(section.props.headerReference ?? []).map((reference) => ({
        kind: 'header' as const,
        id: reference.id,
        type: reference.type,
      })),
      ...(section.props.footerReference ?? []).map((reference) => ({
        kind: 'footer' as const,
        id: reference.id,
        type: reference.type,
      })),
    ]

    for (const reference of references) {
      const key = `${reference.kind}:${reference.id}`
      if (seen.has(key)) continue
      const part: Header | Footer | undefined =
        reference.kind === 'header' ? document.headers.get(reference.id) : document.footers.get(reference.id)
      if (!part) continue
      seen.add(key)
      parts.push({
        kind: reference.kind,
        id: reference.id,
        type: reference.type,
        text: blocksToText(part.blocks),
        hasRichContent: hasRichContent(part.blocks),
      })
    }
  }

  return parts
}

/** Rebuilds a run with new text, keeping its formatting. */
function runWithText(template: Run | undefined, text: string): Run {
  return {
    kind: 'run',
    ...(template?.props !== undefined ? { props: template.props } : {}),
    children: [{ kind: 'text', value: text }],
  }
}

function paragraphWithText(template: Paragraph | undefined, text: string): Paragraph {
  const firstRun = template?.children.find((child): child is Run => child.kind === 'run')
  return {
    kind: 'paragraph',
    ...(template?.props !== undefined ? { props: template.props } : {}),
    children: text === '' ? [] : [runWithText(firstRun, text)],
  }
}

/** Replaces a header/footer part's paragraphs with `text` (one paragraph per line). */
export function setHeaderFooterText(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
  text: string,
): Document {
  const source = kind === 'header' ? document.headers.get(id) : document.footers.get(id)
  if (!source) return document
  if (blocksToText(source.blocks) === text) return document

  const templates = source.blocks.filter((block): block is Paragraph => block.kind === 'paragraph')
  const blocks: Block[] = text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line, index) => paragraphWithText(templates[index] ?? templates[templates.length - 1], line))

  if (kind === 'header') {
    const headers = new Map(document.headers)
    headers.set(id, { ...source, blocks } as Header)
    return { ...document, headers }
  }

  const footers = new Map(document.footers)
  footers.set(id, { ...source, blocks } as Footer)
  return { ...document, footers }
}
