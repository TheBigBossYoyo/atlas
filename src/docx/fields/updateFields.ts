/**
 * Atlas — "Update field(s)" command (DEFER-5 / DXS-20)
 *
 * Walks the whole document — body sections, plus headers/footers/
 * footnotes/endnotes — recalculating every field's cached display text via
 * `evaluateFieldText` and replacing `Field.result`/clearing `Field.raw`
 * wherever evaluation actually produced a new value. A field evaluation
 * leaves unchanged (see `evaluateFieldText`'s doc comment) is left
 * completely untouched, `raw` included, so it still round-trips
 * byte-faithfully on save.
 *
 * `TOC` fields are deliberately skipped here — a whole table of contents
 * regenerating is a many-entry operation with its own switches/generator,
 * not a single recalculated text value; see `./toc.ts`'s
 * `updateTableOfContents` for that.
 *
 * `PAGE`/`NUMPAGES` are also skipped inside a header/footer/footnote/
 * endnote: a header can be rendered on many different pages, so "the"
 * page number isn't a single fixed value the way it is for a body
 * paragraph (which lives on exactly one page) — evaluating it once here
 * would bake in a page number that's only correct for whichever instance
 * happened to be used, silently wrong everywhere else. Every other field
 * type doesn't have that per-render-instance problem and is evaluated
 * normally wherever it appears.
 */
import type {
  Block,
  Document,
  Field,
  HyperlinkChild,
  ParagraphChild,
  RunProps,
  TableChild,
  TableRowChild,
} from '../model'

import { evaluateFieldText } from './evaluate'
import type { FieldEvaluationContext } from './types'

export interface UpdateFieldsResult {
  readonly document: Document
  /** How many fields' cached results actually changed — 0 means nothing needed updating (every field was already current, unevaluable, or a TOC). */
  readonly updatedCount: number
}

export function updateFields(document: Document, context: FieldEvaluationContext): UpdateFieldsResult {
  let updatedCount = 0
  const onUpdated = (): void => {
    updatedCount += 1
  }

  const sections = document.sections.map((section, sectionIndex) => ({
    ...section,
    blocks: updateBlocks(section.blocks, context, true, sectionIndex, onUpdated),
  }))

  const document_ = {
    ...document,
    sections,
    headers: mapMapValues(document.headers, (header) => ({
      ...header,
      blocks: updateBlocks(header.blocks, context, false, undefined, onUpdated),
    })),
    footers: mapMapValues(document.footers, (footer) => ({
      ...footer,
      blocks: updateBlocks(footer.blocks, context, false, undefined, onUpdated),
    })),
    footnotes: mapMapValues(document.footnotes, (footnote) => ({
      ...footnote,
      blocks: updateBlocks(footnote.blocks, context, false, undefined, onUpdated),
    })),
    endnotes: mapMapValues(document.endnotes, (endnote) => ({
      ...endnote,
      blocks: updateBlocks(endnote.blocks, context, false, undefined, onUpdated),
    })),
  }

  return { document: document_, updatedCount }
}

function mapMapValues<K, V>(map: ReadonlyMap<K, V>, transform: (value: V) => V): ReadonlyMap<K, V> {
  const result = new Map<K, V>()
  for (const [key, value] of map) {
    result.set(key, transform(value))
  }
  return result
}

function updateBlocks(
  blocks: ReadonlyArray<Block>,
  context: FieldEvaluationContext,
  allowPageFields: boolean,
  sectionIndex: number | undefined,
  onUpdated: () => void,
): ReadonlyArray<Block> {
  return blocks.map((block, blockIndex) => {
    if (block.kind === 'paragraph') {
      const paragraphPath = sectionIndex !== undefined ? [sectionIndex, blockIndex] : []
      return {
        ...block,
        children: block.children.map((child) =>
          updateParagraphChild(child, context, allowPageFields, paragraphPath, onUpdated),
        ),
      }
    }

    if (block.kind === 'table') {
      return {
        ...block,
        rows: block.rows.map((row) => updateTableRow(row, context, onUpdated)),
      }
    }

    return block
  })
}

function updateTableRow(
  row: TableChild,
  context: FieldEvaluationContext,
  onUpdated: () => void,
): TableChild {
  if (row.kind !== 'table-row') {
    return row
  }
  return {
    ...row,
    cells: row.cells.map((cell) => updateTableCell(cell, context, onUpdated)),
  }
}

function updateTableCell(
  cell: TableRowChild,
  context: FieldEvaluationContext,
  onUpdated: () => void,
): TableRowChild {
  if (cell.kind !== 'table-cell') {
    return cell
  }
  // A table cell's own blocks have no meaningful top-level "section index"
  // of their own for PAGE-field addressing purposes (a cell's paragraphs
  // aren't direct children of a section) — force `allowPageFields` false so
  // a PAGE/NUMPAGES field nested in a table cell is left unevaluated (a
  // documented minor scope limit; PAGE inside a table cell is rare) rather
  // than evaluating it against a meaningless empty path.
  return {
    ...cell,
    blocks: updateBlocks(cell.blocks, context, false, undefined, onUpdated),
  }
}

function updateParagraphChild(
  child: ParagraphChild,
  context: FieldEvaluationContext,
  allowPageFields: boolean,
  paragraphPath: ReadonlyArray<number>,
  onUpdated: () => void,
): ParagraphChild {
  switch (child.kind) {
    case 'field':
      return updateField(child, context, allowPageFields, paragraphPath, onUpdated)
    case 'hyperlink':
      return {
        ...child,
        children: child.children.map((hyperlinkChild) =>
          updateHyperlinkChild(hyperlinkChild, context, allowPageFields, paragraphPath, onUpdated),
        ),
      }
    case 'ins-revision':
    case 'del-revision': {
      const children = (child.children as ReadonlyArray<ParagraphChild>).map((revisionChild) =>
        updateParagraphChild(revisionChild, context, allowPageFields, paragraphPath, onUpdated),
      )
      return { ...child, children } as typeof child
    }
    default:
      return child
  }
}

function updateHyperlinkChild(
  child: HyperlinkChild,
  context: FieldEvaluationContext,
  allowPageFields: boolean,
  paragraphPath: ReadonlyArray<number>,
  onUpdated: () => void,
): HyperlinkChild {
  return child.kind === 'field' ? updateField(child, context, allowPageFields, paragraphPath, onUpdated) : child
}

function updateField(
  field: Field,
  context: FieldEvaluationContext,
  allowPageFields: boolean,
  paragraphPath: ReadonlyArray<number>,
  onUpdated: () => void,
): Field {
  if (field.fieldType === 'TOC') {
    return field
  }
  if ((field.fieldType === 'PAGE' || field.fieldType === 'NUMPAGES') && !allowPageFields) {
    return field
  }
  // `w:fldLock` (`Field.locked`) means the author explicitly froze this
  // field's cached value (e.g. a signing date that must stop tracking
  // "today") — real Word's own Update Field(s)/F9 never recalculates a
  // locked field, so honoring that here isn't optional: silently
  // overwriting a value the document itself marks as locked would corrupt
  // exactly the content the lock exists to protect.
  if (field.locked === true) {
    return field
  }

  const newText = evaluateFieldText(field, paragraphPath, context)
  if (newText === undefined) {
    return field
  }

  onUpdated()
  const props = firstRunProps(field.result)
  const result: ReadonlyArray<ParagraphChild> = [
    {
      kind: 'run',
      ...(props !== undefined ? { props } : {}),
      children: [{ kind: 'text', value: newText }],
    },
  ]

  return { ...field, result, raw: undefined }
}

function firstRunProps(result: ReadonlyArray<ParagraphChild>): RunProps | undefined {
  for (const child of result) {
    if (child.kind === 'run') {
      return child.props
    }
  }
  return undefined
}
