/**
 * D29 — editing a document's headers and footers.
 *
 * Headers and footers live in their own parts, outside the body the caret
 * moves through, and the save path already writes them back from the model
 * (`writeHeaderXml`/`writeFooterXml`). What was missing was any way to change
 * them, so a page number or a title typed in Word could be read but never
 * corrected here.
 *
 * D29 follow-up — the first cut edited a whole part as one block of text (one
 * line per paragraph), rebuilding every paragraph from scratch on save. That
 * silently dropped anything a paragraph held besides plain runs: images,
 * tables, PAGE/NUMPAGES fields, and any formatting beyond the first run's.
 *
 * The model here is per PARAGRAPH, not per part: a paragraph made up
 * entirely of plain runs (text/tab/break — no drawings, fields, hyperlinks,
 * footnote/comment references, or raw unknown XML) gets one editable field;
 * anything else — a table, or a paragraph holding a drawing/field/hyperlink/
 * etc — is shown as a read-only, labelled placeholder ("[Image]", "[Table]",
 * "[Page number]") and is never rewritten. Within an editable paragraph, an
 * edit is applied as a prefix/suffix diff against the paragraph's OWN runs
 * (see `paragraphWithEditedText`): text outside the changed span keeps its
 * exact original run object (and so its exact formatting), and only the
 * touched span is rebuilt, using the formatting of the run right before it —
 * the same "diff, don't flatten" approach the multi-run case in the plain
 * body editor relies on, scoped down to one paragraph's own runs.
 *
 * Whole paragraphs can still be added (appended, seeded with the last
 * editable paragraph's own properties) or removed (only a paragraph this
 * editor could have produced itself, and never the part's last remaining
 * block) — see `buildInsertHeaderFooterParagraph`/`buildRemoveHeaderFooterParagraph`.
 */
import type { Block, Document, Field, Footer, Header, Paragraph, ParagraphChild, Run, RunChild } from '../model'
import type { ReplaceHeaderFooterBlocksCommand } from './commandTypes'

export type HeaderFooterKind = 'header' | 'footer'

/** One row of the editing panel: either a plain-text paragraph (editable),
 * or a read-only summary of a block this editor leaves untouched. */
export type HeaderFooterRow =
  | {
      readonly kind: 'text'
      /** Index into the part's own `blocks` array — stable identity for edits/removal. */
      readonly blockIndex: number
      readonly text: string
    }
  | {
      readonly kind: 'placeholder'
      readonly blockIndex: number
      /** Human-readable summary, e.g. "[Image]", "[Table]", "Page [Page number]". */
      readonly label: string
    }

export type HeaderFooterPart = {
  readonly kind: HeaderFooterKind
  readonly id: string
  /** `default`, `first` or `even`, as referenced by the section that uses it. */
  readonly type: string
  readonly rows: ReadonlyArray<HeaderFooterRow>
}

// ---------------------------------------------------------------------------
// Reading — plain-text detection and placeholder labelling
// ---------------------------------------------------------------------------

function isPlainRunChild(child: RunChild): boolean {
  return child.kind === 'text' || child.kind === 'tab' || child.kind === 'break'
}

/** A paragraph this editor can safely rewrite from a single string: every
 * child is a run, and every run holds only text/tab/break — no drawings,
 * fields, footnote/comment references, or raw unknown XML. */
function isPlainTextParagraph(paragraph: Paragraph): boolean {
  return paragraph.children.every((child) => child.kind === 'run' && child.children.every(isPlainRunChild))
}

function fieldLabel(field: Field): string {
  switch (field.fieldType) {
    case 'PAGE':
      return '[Page number]'
    case 'NUMPAGES':
      return '[Page count]'
    default:
      return '[Field]'
  }
}

function runChildLabel(child: RunChild): string {
  switch (child.kind) {
    case 'text':
      return child.value
    case 'tab':
      return '\t'
    case 'break':
      return '\n'
    case 'drawing':
      return '[Image]'
    case 'comment-reference':
      return '[Comment]'
    case 'footnote-reference':
      return '[Footnote]'
    case 'endnote-reference':
      return '[Endnote]'
    case 'unknown':
      return '[Other content]'
  }
}

function runLabel(run: Run): string {
  return run.children.map(runChildLabel).join('')
}

function paragraphChildLabel(child: ParagraphChild): string {
  switch (child.kind) {
    case 'run':
      return runLabel(child)
    case 'hyperlink':
      return child.children
        .map((grandchild) => (grandchild.kind === 'run' ? runLabel(grandchild) : grandchild.kind === 'field' ? fieldLabel(grandchild) : ''))
        .join('')
    case 'field':
      return fieldLabel(child)
    case 'ins-revision':
    case 'del-revision':
      return child.children.map(runLabel).join('')
    case 'bookmark':
    case 'comment-range':
      return ''
    case 'comment-reference':
      return '[Comment]'
    case 'footnote-reference':
      return '[Footnote]'
    case 'endnote-reference':
      return '[Endnote]'
    case 'unknown':
      return '[Other content]'
  }
}

/** Read-only summary for a block this editor won't rewrite. */
function blockLabel(block: Block): string {
  if (block.kind === 'table') {
    return '[Table]'
  }
  if (block.kind === 'unknown') {
    return '[Unknown content]'
  }
  const label = block.children.map(paragraphChildLabel).join('').trim()
  return label.length > 0 ? label : '[Empty paragraph]'
}

function partRows(blocks: ReadonlyArray<Block>): ReadonlyArray<HeaderFooterRow> {
  return blocks.map((block, blockIndex) => {
    if (block.kind === 'paragraph' && isPlainTextParagraph(block)) {
      return { kind: 'text', blockIndex, text: paragraphPlainText(block) }
    }
    return { kind: 'placeholder', blockIndex, label: blockLabel(block) }
  })
}

function paragraphPlainText(paragraph: Paragraph): string {
  let text = ''
  for (const child of paragraph.children) {
    if (child.kind !== 'run') continue
    for (const item of child.children) {
      if (item.kind === 'text') text += item.value
      else if (item.kind === 'tab') text += '\t'
      else if (item.kind === 'break') text += '\n'
    }
  }
  return text
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
        rows: partRows(part.blocks),
      })
    }
  }

  return parts
}

// ---------------------------------------------------------------------------
// Writing — structural block replace, and the per-paragraph text diff
// ---------------------------------------------------------------------------

function getPart(document: Document, kind: HeaderFooterKind, id: string): Header | Footer | undefined {
  return kind === 'header' ? document.headers.get(id) : document.footers.get(id)
}

function withPartBlocks(document: Document, kind: HeaderFooterKind, id: string, blocks: ReadonlyArray<Block>): Document {
  const source = getPart(document, kind, id)
  if (!source) return document

  if (kind === 'header') {
    const headers = new Map(document.headers)
    headers.set(id, { ...(source as Header), blocks })
    return { ...document, headers }
  }

  const footers = new Map(document.footers)
  footers.set(id, { ...(source as Footer), blocks })
  return { ...document, footers }
}

/** Pure structural splice — the low-level operation `replace-header-footer-blocks`
 * applies; see that command's doc comment in `commandTypes.ts`. */
export function replaceHeaderFooterBlocks(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
  at: number,
  count: number,
  blocks: ReadonlyArray<Block>,
): Document {
  const source = getPart(document, kind, id)
  if (!source) return document
  const nextBlocks = source.blocks.slice(0, at).concat(blocks, source.blocks.slice(at + count))
  return withPartBlocks(document, kind, id, nextBlocks)
}

function runPlainText(run: Run): string {
  let text = ''
  for (const item of run.children) {
    if (item.kind === 'text') text += item.value
    else if (item.kind === 'tab') text += '\t'
    else if (item.kind === 'break') text += '\n'
  }
  return text
}

function textToRunChildren(text: string): RunChild[] {
  const children: RunChild[] = []
  let buffer = ''
  for (const char of text) {
    if (char === '\t' || char === '\n') {
      if (buffer.length > 0) {
        children.push({ kind: 'text', value: buffer })
        buffer = ''
      }
      children.push(char === '\t' ? { kind: 'tab' } : { kind: 'break' })
    } else {
      buffer += char
    }
  }
  if (buffer.length > 0 || children.length === 0) {
    children.push({ kind: 'text', value: buffer })
  }
  return children
}

/** Rebuilds a run with new text, keeping its formatting (props/rsid). */
function runWithPlainText(template: Run | undefined, text: string): Run {
  return {
    kind: 'run',
    ...(template?.props !== undefined ? { props: template.props } : {}),
    children: textToRunChildren(text),
  }
}

/**
 * Rewrites a plain-text paragraph's own runs to match `newText`, keeping
 * every run the edit didn't touch as the exact same object (so its
 * formatting survives untouched), and rebuilding only the smallest span the
 * old and new text actually disagree on — a common-prefix/common-suffix diff
 * against the paragraph's own concatenated text. The replacement text for
 * that span takes the formatting of whichever run it lands in: the run being
 * replaced, when there's existing text being overwritten, or — when nothing
 * is being removed at all (a pure insertion at some caret position) — the
 * run immediately to the left of the caret, matching Word's own "typed text
 * inherits the formatting immediately to its left" behavior. Returns the
 * same `paragraph` object when the text didn't actually change.
 */
function paragraphWithEditedText(paragraph: Paragraph, newText: string): Paragraph {
  // isPlainTextParagraph guarantees every child is a Run.
  const runs = paragraph.children as ReadonlyArray<Run>
  const oldText = runs.map(runPlainText).join('')
  if (oldText === newText) return paragraph

  const maxPrefix = Math.min(oldText.length, newText.length)
  let prefixLen = 0
  while (prefixLen < maxPrefix && oldText[prefixLen] === newText[prefixLen]) prefixLen++

  const maxSuffix = maxPrefix - prefixLen
  let suffixLen = 0
  while (
    suffixLen < maxSuffix &&
    oldText[oldText.length - 1 - suffixLen] === newText[newText.length - 1 - suffixLen]
  ) {
    suffixLen++
  }

  const changeStart = prefixLen
  const changeOldEnd = oldText.length - suffixLen
  const insertedText = newText.slice(prefixLen, newText.length - suffixLen)

  // Formatting for the inserted text is resolved once, up front, against the
  // ORIGINAL run spans — not tracked incrementally while splicing below,
  // which conflated "the run just before the change" with "the run the
  // change replaces" and picked the wrong one whenever a replacement
  // happened to start exactly on a run boundary.
  const formatTemplate = findFormatTemplate(runs, changeStart, changeStart === changeOldEnd)

  const nextRuns: Run[] = []
  let offset = 0
  let insertedHere = false

  for (const run of runs) {
    const text = runPlainText(run)
    const runStart = offset
    const runEnd = offset + text.length
    offset = runEnd

    if (runEnd <= changeStart || runStart >= changeOldEnd) {
      // Entirely outside the changed span: reused verbatim.
      nextRuns.push(run)
      continue
    }

    // Overlaps the changed span: keep whichever of its own text falls
    // outside [changeStart, changeOldEnd), and drop the rest — the
    // replacement text is spliced in once, at the first overlapping run.
    const before = text.slice(0, Math.max(0, changeStart - runStart))
    const after = text.slice(Math.max(0, changeOldEnd - runStart))

    if (before.length > 0) {
      nextRuns.push(runWithPlainText(run, before))
    }
    if (!insertedHere) {
      if (insertedText.length > 0) {
        nextRuns.push(runWithPlainText(formatTemplate, insertedText))
      }
      insertedHere = true
    }
    if (after.length > 0) {
      nextRuns.push(runWithPlainText(run, after))
    }
  }

  if (!insertedHere && insertedText.length > 0) {
    // Pure append past the last run (or the paragraph had no runs at all).
    nextRuns.push(runWithPlainText(formatTemplate, insertedText))
  }

  return { ...paragraph, children: nextRuns }
}

/**
 * The run whose formatting the text landing at `changeStart` should take.
 * For a replacement (`isPureInsertion` false), that's the run `changeStart`
 * falls inside (the first run being overwritten). For a pure insertion (a
 * caret position, nothing removed), that's the run ending exactly there —
 * i.e. immediately to the caret's left — falling back to whatever run
 * contains it when `changeStart` doesn't land on a boundary. Undefined only
 * when the paragraph has no runs at all (an empty paragraph gaining its
 * first text), in which case the new run carries no explicit formatting.
 */
function findFormatTemplate(runs: ReadonlyArray<Run>, changeStart: number, isPureInsertion: boolean): Run | undefined {
  let offset = 0
  let runBeforeStart: Run | undefined
  let runContaining: Run | undefined

  for (const run of runs) {
    const text = runPlainText(run)
    const runStart = offset
    const runEnd = offset + text.length
    offset = runEnd

    if (runEnd === changeStart) runBeforeStart = run
    if (runContaining === undefined && runStart <= changeStart && changeStart < runEnd) runContaining = run
  }

  if (isPureInsertion) {
    return runBeforeStart ?? runContaining ?? runs[runs.length - 1]
  }
  return runContaining ?? runBeforeStart ?? runs[runs.length - 1]
}

/**
 * Builds the command that rewrites one plain-text paragraph's text, or
 * `null` when there is nothing to do: the part/block doesn't exist, the
 * block isn't a plain-text paragraph (it holds a drawing/field/table/etc —
 * this editor never touches those), or the text didn't actually change.
 */
export function buildHeaderFooterTextEdit(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
  blockIndex: number,
  text: string,
): ReplaceHeaderFooterBlocksCommand | null {
  const source = getPart(document, kind, id)
  const block = source?.blocks[blockIndex]
  if (block === undefined || block.kind !== 'paragraph' || !isPlainTextParagraph(block)) {
    return null
  }

  const normalized = text.replace(/\r\n?/g, '\n')
  const nextParagraph = paragraphWithEditedText(block, normalized)
  if (nextParagraph === block) {
    return null
  }

  return { kind: 'replace-header-footer-blocks', target: kind, id, at: blockIndex, count: 1, blocks: [nextParagraph] }
}

/** Document-returning convenience over `buildHeaderFooterTextEdit`, mainly for tests. */
export function setHeaderFooterBlockText(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
  blockIndex: number,
  text: string,
): Document {
  const command = buildHeaderFooterTextEdit(document, kind, id, blockIndex, text)
  return command === null ? document : replaceHeaderFooterBlocks(document, kind, id, command.at, command.count, command.blocks)
}

/** Appends a new, empty plain-text paragraph, seeded with the properties of
 * the part's last plain-text paragraph (if any) so it looks at home. `null`
 * when the part doesn't exist. */
export function buildInsertHeaderFooterParagraph(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
): ReplaceHeaderFooterBlocksCommand | null {
  const source = getPart(document, kind, id)
  if (!source) return null

  const templates = source.blocks.filter((block): block is Paragraph => block.kind === 'paragraph' && isPlainTextParagraph(block))
  const template = templates[templates.length - 1]
  const paragraph: Paragraph = {
    kind: 'paragraph',
    ...(template?.props !== undefined ? { props: template.props } : {}),
    children: [],
  }

  return { kind: 'replace-header-footer-blocks', target: kind, id, at: source.blocks.length, count: 0, blocks: [paragraph] }
}

/** Removes a plain-text paragraph. `null` (a no-op) when the part doesn't
 * exist, `blockIndex` isn't a plain-text paragraph, or it's the part's only
 * remaining block — a header/footer part is never emptied out entirely from
 * this panel. */
export function buildRemoveHeaderFooterParagraph(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
  blockIndex: number,
): ReplaceHeaderFooterBlocksCommand | null {
  const source = getPart(document, kind, id)
  if (!source || source.blocks.length <= 1) return null
  const block = source.blocks[blockIndex]
  if (block === undefined || block.kind !== 'paragraph' || !isPlainTextParagraph(block)) return null

  return { kind: 'replace-header-footer-blocks', target: kind, id, at: blockIndex, count: 1, blocks: [] }
}
