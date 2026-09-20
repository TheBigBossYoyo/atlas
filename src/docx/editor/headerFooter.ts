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
 * D29 follow-up 2 (mixed paragraphs) — the per-paragraph model above still
 * couldn't touch the single most common real-world header/footer shape: a
 * paragraph that mixes plain text WITH a drawing/field/hyperlink/etc, like
 * "Chapter title .......... Page X of Y" (text, a tab, a PAGE field, more
 * text, a NUMPAGES field) or a logo image followed by a title. Such a
 * paragraph used to be ONE placeholder, atom and all, with no way to edit
 * even the parts that are plain text.
 *
 * The model is now per SEGMENT within a paragraph, not just per paragraph:
 * `segmentParagraphChildren` walks a paragraph's own `children` and groups
 * them into alternating spans —
 *   - a run of one or more consecutive plain runs (text/tab/break only) is
 *     one EDITABLE text segment, addressed by its position among a
 *     paragraph's own text segments (`segmentIndex`, 0-based, left to
 *     right);
 *   - anything else — a drawing, a field, a hyperlink, a footnote/comment
 *     reference, an existing tracked-change wrapper, or raw unknown XML —
 *     is one READ-ONLY atom, labelled for the panel ("[Image]", "[Page
 *     number]", "[Link: …]", …) and never rewritten.
 * `w:bookmarkStart`/`w:bookmarkEnd` and comment-range markers are also
 * atoms structurally (they still occupy their own position in `children`,
 * so editing text on either side of one can't accidentally merge across
 * it), but carry no visible label — the panel never shows a chip for one,
 * matching how they've always rendered invisibly in this editor's labels.
 *
 * Three paragraph shapes fall out of this:
 *   - every child is part of one editable run of text (no atoms at all) —
 *     unchanged from before, one editable field, row kind `'text'`;
 *   - no editable text anywhere (every child is an atom) — unchanged from
 *     before, one read-only placeholder, row kind `'placeholder'`;
 *   - a genuine mix — a NEW row kind, `'mixed'`, one editable text input per
 *     text segment interleaved with a read-only chip per (non-hidden) atom.
 *
 * Editing one text segment (`buildHeaderFooterSegmentEdit`) reuses exactly
 * the same prefix/suffix diff `paragraphWithEditedText` already used for a
 * whole plain paragraph (factored out as `editRunsForText`), just scoped to
 * that segment's own slice of `children` instead of the whole array — text
 * outside the touched span keeps its exact original run object and
 * formatting, only the touched span is rebuilt, and every atom (and every
 * OTHER text segment) is spliced back in completely untouched, at its exact
 * original position. A field's own runs (`w:fldSimple`, or the
 * `w:fldChar`/`w:instrText` triple) are captured verbatim in `Field.raw` and
 * this editor never reads or rewrites that field node at all, so they can
 * never come apart from being edited around.
 *
 * Deleting ALL the text in a segment removes that segment's run(s) from
 * `children` entirely — the same thing already happened when the sole text
 * segment of a plain paragraph was emptied out (see `editRunsForText`: an
 * edit that inserts nothing simply pushes nothing for that span). If the
 * user empties out every text segment around an atom that way, the
 * paragraph is left holding only atoms, with no editable text anywhere —
 * on the next read it reclassifies as an ordinary `'placeholder'` row, same
 * as a paragraph that was always atom-only. The atom itself is never lost:
 * removing the whole line (if it's otherwise unwanted) is still available
 * via the panel's own remove-paragraph button.
 *
 * A run whose OWN children mix plain content with a drawing/field/etc in
 * the same `<w:r>` (legal per the schema, but not real Word/LibreOffice
 * output — those always give a drawing or field its own run) is treated as
 * ONE atom for the whole run rather than split further, labelled with the
 * concatenation of its children's own labels (`runLabel`) — safer than
 * guessing how to split formatting/rPr across a run boundary schemas don't
 * actually require, for a shape no real-world document produces.
 *
 * Whole paragraphs can still be added (appended, seeded with the last
 * editable paragraph's own properties) or removed (only a paragraph this
 * editor could have produced itself, and never the part's last remaining
 * block) — see `buildInsertHeaderFooterParagraph`/`buildRemoveHeaderFooterParagraph`.
 */
import type { Block, Document, Field, Footer, Header, Hyperlink, Paragraph, ParagraphChild, Run, RunChild } from '../model'
import type { ReplaceHeaderFooterBlocksCommand } from './commandTypes'

export type HeaderFooterKind = 'header' | 'footer'

/** One piece of a `'mixed'` row: either an editable span of plain text
 * (`segmentIndex` is this text segment's 0-based position among the
 * paragraph's OWN text segments, left to right — the identifier
 * `buildHeaderFooterSegmentEdit` takes to know which one to rewrite), or a
 * read-only atom label for a drawing/field/hyperlink/etc this editor leaves
 * exactly as it is. A bookmark/comment-range boundary is an atom
 * structurally but is never surfaced here (see the module doc comment) —
 * it never appears in this array at all. */
export type HeaderFooterSegment =
  | { readonly kind: 'text'; readonly segmentIndex: number; readonly text: string }
  | { readonly kind: 'atom'; readonly label: string }

/** One row of the editing panel: a plain-text paragraph (one editable
 * field), a paragraph mixing text with atoms (one editable field per text
 * segment, interleaved with read-only atom chips), or a read-only summary
 * of a block this editor leaves untouched entirely. */
export type HeaderFooterRow =
  | {
      readonly kind: 'text'
      /** Index into the part's own `blocks` array — stable identity for edits/removal. */
      readonly blockIndex: number
      readonly text: string
    }
  | {
      readonly kind: 'mixed'
      readonly blockIndex: number
      readonly segments: ReadonlyArray<HeaderFooterSegment>
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
      return '[Total pages]'
    case 'DATE':
      return '[Date]'
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

/** A run this editor can rewrite as plain text on its own: every child is
 * text/tab/break, matching `isPlainTextParagraph`'s per-run test. A run
 * mixing plain content with a drawing/field/etc (legal, but not real
 * Word/LibreOffice output) is NOT plain — see the module doc comment on why
 * it's treated as one atom rather than split further. */
function isPlainRun(run: Run): boolean {
  return run.children.every(isPlainRunChild)
}

/** The visible label for an atom paragraph child — everything a `'mixed'`
 * row's segments can hold besides an editable text span. `hidden` is true
 * only for a bookmark/comment-range boundary: it still occupies its own
 * position (so text on either side of it is never fused into one editable
 * segment), but the panel never renders a chip for it, matching how it has
 * always contributed nothing to this editor's labels. */
function atomChildLabel(child: ParagraphChild): { readonly label: string; readonly hidden: boolean } {
  switch (child.kind) {
    case 'run':
      return { label: runLabel(child), hidden: false }
    case 'hyperlink':
      return { label: hyperlinkAtomLabel(child), hidden: false }
    case 'field':
      return { label: fieldLabel(child), hidden: false }
    case 'bookmark':
    case 'comment-range':
      return { label: '', hidden: true }
    case 'comment-reference':
      return { label: '[Comment]', hidden: false }
    case 'footnote-reference':
      return { label: '[Footnote]', hidden: false }
    case 'endnote-reference':
      return { label: '[Endnote]', hidden: false }
    case 'ins-revision':
    case 'del-revision':
      return { label: '[Tracked change]', hidden: false }
    case 'unknown':
      return { label: '[Other content]', hidden: false }
  }
}

function hyperlinkAtomLabel(hyperlink: Hyperlink): string {
  const text = hyperlink.children
    .map((child) => (child.kind === 'run' ? runLabel(child) : child.kind === 'field' ? fieldLabel(child) : ''))
    .join('')
    .trim()
  return text.length > 0 ? `[Link: ${text}]` : '[Link]'
}

/** One contiguous span of a paragraph's own `children`: either a maximal
 * run of consecutive plain runs (fused into one editable text segment, the
 * same "fuse adjacent runs, diff the concatenated text" approach the
 * whole-paragraph case already used), or a single non-plain child treated
 * as one read-only atom. `startChildIndex`/`endChildIndex` are the (start
 * inclusive, end exclusive) slice of `children` this segment owns — the
 * exact range `paragraphWithEditedSegment` splices a rebuilt slice back
 * into. */
type ParagraphSegment =
  | { readonly kind: 'text'; readonly startChildIndex: number; readonly endChildIndex: number; readonly text: string }
  | {
      readonly kind: 'atom'
      readonly startChildIndex: number
      readonly endChildIndex: number
      readonly label: string
      readonly hidden: boolean
    }

function segmentParagraphChildren(children: ReadonlyArray<ParagraphChild>): ReadonlyArray<ParagraphSegment> {
  const segments: ParagraphSegment[] = []
  let index = 0

  while (index < children.length) {
    const child = children[index]

    if (child.kind === 'run' && isPlainRun(child)) {
      let end = index
      let text = ''
      while (end < children.length) {
        const candidate = children[end]
        if (candidate.kind !== 'run' || !isPlainRun(candidate)) break
        text += runPlainText(candidate)
        end++
      }
      segments.push({ kind: 'text', startChildIndex: index, endChildIndex: end, text })
      index = end
      continue
    }

    const { label, hidden } = atomChildLabel(child)
    segments.push({ kind: 'atom', startChildIndex: index, endChildIndex: index + 1, label, hidden })
    index++
  }

  return segments
}

/**
 * The panel-facing segment list for a `'mixed'` row, or `null` when the
 * paragraph has no editable text anywhere (every child is an atom) — that
 * paragraph stays a `'placeholder'` row instead, exactly as before this
 * feature existed. Hidden atoms (bookmarks/comment-ranges) are dropped from
 * the returned list entirely; visible text segments are numbered by their
 * order among ALL text segments (hidden atoms don't consume a number, but
 * they do still split two text runs they sit between into separate
 * segments — see the module doc comment).
 */
function mixedRowSegments(paragraph: Paragraph): ReadonlyArray<HeaderFooterSegment> | null {
  const segments = segmentParagraphChildren(paragraph.children)
  if (!segments.some((segment) => segment.kind === 'text')) {
    return null
  }

  const rows: HeaderFooterSegment[] = []
  let segmentIndex = 0
  for (const segment of segments) {
    if (segment.kind === 'text') {
      rows.push({ kind: 'text', segmentIndex, text: segment.text })
      segmentIndex++
    } else if (!segment.hidden) {
      rows.push({ kind: 'atom', label: segment.label })
    }
  }
  return rows
}

function partRows(blocks: ReadonlyArray<Block>): ReadonlyArray<HeaderFooterRow> {
  return blocks.map((block, blockIndex) => {
    if (block.kind === 'paragraph') {
      if (isPlainTextParagraph(block)) {
        return { kind: 'text', blockIndex, text: paragraphPlainText(block) }
      }
      const segments = mixedRowSegments(block)
      if (segments !== null) {
        return { kind: 'mixed', blockIndex, segments }
      }
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
 * Rewrites a plain-text paragraph's own runs to match `newText` — a thin
 * wrapper over `editRunsForText` scoped to the WHOLE paragraph (every child
 * is guaranteed to be a `Run` by `isPlainTextParagraph`). Returns the same
 * `paragraph` object when the text didn't actually change.
 */
function paragraphWithEditedText(paragraph: Paragraph, newText: string): Paragraph {
  const runs = paragraph.children as ReadonlyArray<Run>
  const nextRuns = editRunsForText(runs, newText)
  if (nextRuns === runs) return paragraph
  return { ...paragraph, children: nextRuns }
}

/**
 * Rewrites one paragraph's own text SEGMENT — `children.slice(startChildIndex,
 * endChildIndex)`, a maximal run of consecutive plain runs identified by
 * `segmentParagraphChildren` — to match `newText`, splicing the rebuilt runs
 * back into the paragraph's full `children` array. Every child outside that
 * slice (earlier/later text segments, and every atom — a drawing, a field,
 * a hyperlink, anything else this editor doesn't touch) is carried through
 * completely untouched, at its exact original position. Returns `null` when
 * the text within the slice didn't actually change.
 */
function paragraphWithEditedSegment(
  paragraph: Paragraph,
  startChildIndex: number,
  endChildIndex: number,
  newText: string,
): Paragraph | null {
  const slice = paragraph.children.slice(startChildIndex, endChildIndex) as ReadonlyArray<Run>
  const nextSlice = editRunsForText(slice, newText)
  if (nextSlice === slice) return null

  const children = [
    ...paragraph.children.slice(0, startChildIndex),
    ...nextSlice,
    ...paragraph.children.slice(endChildIndex),
  ]
  return { ...paragraph, children }
}

/**
 * The shared core both `paragraphWithEditedText` (a whole plain paragraph)
 * and `paragraphWithEditedSegment` (one text segment of a mixed paragraph)
 * rebuild from: keeps every run the edit didn't touch as the exact same
 * object (so its formatting survives untouched), and rebuilds only the
 * smallest span the old and new text actually disagree on — a
 * common-prefix/common-suffix diff against `runs`' own concatenated text.
 * The replacement text for that span takes the formatting of whichever run
 * it lands in: the run being replaced, when there's existing text being
 * overwritten, or — when nothing is being removed at all (a pure insertion
 * at some caret position) — the run immediately to the left of the caret,
 * matching Word's own "typed text inherits the formatting immediately to
 * its left" behavior. Returns the same `runs` array (by reference) when the
 * text didn't actually change, so callers can cheaply detect a no-op.
 */
function editRunsForText(runs: ReadonlyArray<Run>, newText: string): ReadonlyArray<Run> {
  const oldText = runs.map(runPlainText).join('')
  if (oldText === newText) return runs

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

  return nextRuns
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

/**
 * Builds the command that rewrites one TEXT SEGMENT of a `'mixed'` row's
 * paragraph — the counterpart to `buildHeaderFooterTextEdit` for a
 * paragraph that mixes plain text with a drawing/field/hyperlink/etc (see
 * the module doc comment). `segmentIndex` addresses the segment the same
 * way `HeaderFooterSegment.segmentIndex` does: 0-based, left to right,
 * counting only text segments. `null` when there's nothing to do: the
 * part/block doesn't exist, the block is a fully plain paragraph (use
 * `buildHeaderFooterTextEdit` for that — this function deliberately never
 * touches it, so the two stay each other's exact complement) or not a
 * paragraph at all, `segmentIndex` is out of range, or the text didn't
 * actually change.
 */
export function buildHeaderFooterSegmentEdit(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
  blockIndex: number,
  segmentIndex: number,
  text: string,
): ReplaceHeaderFooterBlocksCommand | null {
  const source = getPart(document, kind, id)
  const block = source?.blocks[blockIndex]
  if (block === undefined || block.kind !== 'paragraph' || isPlainTextParagraph(block)) {
    return null
  }

  const textSegments = segmentParagraphChildren(block.children).filter(
    (segment): segment is Extract<ParagraphSegment, { kind: 'text' }> => segment.kind === 'text',
  )
  const target = textSegments[segmentIndex]
  if (target === undefined) {
    return null
  }

  const normalized = text.replace(/\r\n?/g, '\n')
  const nextParagraph = paragraphWithEditedSegment(block, target.startChildIndex, target.endChildIndex, normalized)
  if (nextParagraph === null) {
    return null
  }

  return { kind: 'replace-header-footer-blocks', target: kind, id, at: blockIndex, count: 1, blocks: [nextParagraph] }
}

/** Document-returning convenience over `buildHeaderFooterSegmentEdit`, mainly for tests. */
export function setHeaderFooterSegmentText(
  document: Document,
  kind: HeaderFooterKind,
  id: string,
  blockIndex: number,
  segmentIndex: number,
  text: string,
): Document {
  const command = buildHeaderFooterSegmentEdit(document, kind, id, blockIndex, segmentIndex, text)
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
