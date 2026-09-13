import type {
  Block,
  Comment,
  CommentRange,
  CommentReference,
  Document,
  Hyperlink,
  HyperlinkChild,
  Paragraph,
  ParagraphChild,
  Run,
  RunChild,
  Section,
  TableChild,
  TableRowChild,
  TextNode,
} from '../model'
import type { Position, Range } from './commandTypes'
import { normalizeRange } from './Selection'

type RunLocation =
  | {
      readonly kind: 'paragraph'
      readonly childIndex: number
      readonly run: Run
    }
  | {
      readonly kind: 'hyperlink'
      readonly childIndex: number
      readonly runIndex: number
      readonly hyperlink: Hyperlink
      readonly run: Run
    }

type InsertableNode = ParagraphChild | HyperlinkChild

export function addCommentToDocument(
  document: Document,
  range: Range,
  text: string,
  author: string,
): { readonly document: Document; readonly commentId: string } {
  const normalized = normalizeRange(range)
  const commentId = nextCommentId(document)
  const withEnd = insertCommentBoundary(document, normalized.end, commentId, 'end-with-reference')
  const withStart = insertCommentBoundary(withEnd, normalized.start, commentId, 'start')
  const comments = new Map(withStart.comments)
  comments.set(
    commentId,
    createComment({
      id: commentId,
      author,
      text,
    }),
  )

  return {
    document: {
      ...withStart,
      comments,
    },
    commentId,
  }
}

export function replyToComment(document: Document, parentId: string, text: string, author: string): Document {
  const commentId = nextCommentId(document)
  const comments = new Map(document.comments)
  comments.set(
    commentId,
    createComment({
      id: commentId,
      author,
      text,
      parentId,
    }),
  )

  return {
    ...document,
    comments,
  }
}

export function deleteCommentFromDocument(document: Document, commentId: string): Document {
  const idsToDelete = collectCommentDescendants(document.comments, commentId)
  const comments = new Map(document.comments)

  for (const id of idsToDelete) {
    comments.delete(id)
  }

  return {
    ...document,
    comments,
    sections: document.sections.map((section) => stripCommentIdsFromSection(section, idsToDelete)),
  }
}

function createComment(input: {
  readonly id: string
  readonly author: string
  readonly text: string
  readonly parentId?: string
}): Comment {
  return {
    kind: 'comment',
    id: input.id,
    author: input.author,
    initials: buildInitials(input.author),
    date: new Date().toISOString(),
    ...(input.parentId !== undefined ? { parentId: input.parentId } : {}),
    body: createCommentBody(input.text),
  }
}

function buildInitials(author: string): string {
  const parts = author.trim().split(/\s+/).filter(Boolean)
  if (parts.length === 0) {
    return 'AT'
  }
  if (parts.length === 1) {
    return parts[0].slice(0, 2).toUpperCase()
  }
  return `${parts[0][0] ?? ''}${parts[parts.length - 1][0] ?? ''}`.toUpperCase()
}

function createCommentBody(text: string): ReadonlyArray<Paragraph> {
  const normalized = text.replace(/\r\n/g, '\n').trim()
  if (normalized.length === 0) {
    return [createParagraph('')]
  }

  return normalized.split(/\n{2,}/).map((paragraphText) => createParagraph(paragraphText.replace(/\n/g, ' ')))
}

function createParagraph(text: string): Paragraph {
  return {
    kind: 'paragraph',
    children:
      text.length > 0
        ? [
            {
              kind: 'run',
              children: [{ kind: 'text', value: text }],
            },
          ]
        : [],
  }
}

function nextCommentId(document: Document): string {
  let maxNumericId = -1

  for (const id of document.comments.keys()) {
    const numeric = Number(id)
    if (Number.isFinite(numeric)) {
      maxNumericId = Math.max(maxNumericId, numeric)
    }
  }

  return String(maxNumericId + 1)
}

function collectCommentDescendants(comments: ReadonlyMap<string, Comment>, commentId: string): Set<string> {
  const result = new Set<string>([commentId])
  let changed = true

  while (changed) {
    changed = false
    for (const [id, comment] of comments) {
      if (comment.parentId !== undefined && result.has(comment.parentId) && !result.has(id)) {
        result.add(id)
        changed = true
      }
    }
  }

  return result
}

function insertCommentBoundary(
  document: Document,
  position: Position,
  commentId: string,
  mode: 'start' | 'end-with-reference',
): Document {
  return updateParagraphAtPath(document, position.paragraphPath, (paragraph) =>
    insertIntoParagraph(paragraph, position, commentId, mode),
  )
}

function updateParagraphAtPath(
  document: Document,
  paragraphPath: ReadonlyArray<number>,
  update: (paragraph: Paragraph) => Paragraph,
): Document {
  const sectionIndex = paragraphPath[0]
  const section = document.sections[sectionIndex]
  if (section === undefined) {
    return document
  }

  const nextBlocks = updateBlocksAtPath(section.blocks, paragraphPath.slice(1), update)
  if (nextBlocks === section.blocks) {
    return document
  }

  const nextSections = document.sections.slice()
  nextSections[sectionIndex] = {
    ...section,
    blocks: nextBlocks,
  }

  return {
    ...document,
    sections: nextSections,
  }
}

function updateBlocksAtPath(
  blocks: ReadonlyArray<Block>,
  path: ReadonlyArray<number>,
  update: (paragraph: Paragraph) => Paragraph,
): ReadonlyArray<Block> {
  const blockIndex = path[0]
  const block = blocks[blockIndex]
  if (block === undefined) {
    return blocks
  }

  if (path.length === 1) {
    if (block.kind !== 'paragraph') {
      return blocks
    }

    const nextParagraph = update(block)
    if (nextParagraph === block) {
      return blocks
    }

    const nextBlocks = blocks.slice()
    nextBlocks[blockIndex] = nextParagraph
    return nextBlocks
  }

  if (block.kind !== 'table') {
    return blocks
  }

  const rowIndex = path[1]
  const cellIndex = path[2]
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return blocks
  }

  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return blocks
  }

  const nextCellBlocks = updateBlocksAtPath(cell.blocks, path.slice(3), update)
  if (nextCellBlocks === cell.blocks) {
    return blocks
  }

  const nextRows = block.rows.slice()
  const nextCells = row.cells.slice()
  nextCells[cellIndex] = {
    ...cell,
    blocks: nextCellBlocks,
  }
  nextRows[rowIndex] = {
    ...row,
    cells: nextCells,
  }

  const nextBlocks = blocks.slice()
  nextBlocks[blockIndex] = {
    ...block,
    rows: nextRows,
  }
  return nextBlocks
}

function insertIntoParagraph(
  paragraph: Paragraph,
  position: Position,
  commentId: string,
  mode: 'start' | 'end-with-reference',
): Paragraph {
  const location = findRunLocation(paragraph, position.runIndex)
  if (location === null) {
    return paragraph
  }

  const nodes =
    mode === 'start'
      ? [createCommentRange(commentId, 'start')]
      : [createCommentRange(commentId, 'end'), createCommentReferenceRun(commentId)]

  if (location.kind === 'paragraph') {
    return {
      ...paragraph,
      children: insertNodesAroundRun(paragraph.children, location.childIndex, location.run, position.charOffset, nodes),
    }
  }

  const nextChildren = paragraph.children.slice()
  const nextHyperlinkChildren = insertNodesAroundRun(
    location.hyperlink.children,
    location.runIndex,
    location.run,
    position.charOffset,
    nodes,
  )

  nextChildren[location.childIndex] = {
    ...location.hyperlink,
    children: nextHyperlinkChildren,
  }

  return {
    ...paragraph,
    children: nextChildren,
  }
}

function findRunLocation(paragraph: Paragraph, runIndex: number): RunLocation | null {
  let currentRunIndex = 0

  for (let childIndex = 0; childIndex < paragraph.children.length; childIndex += 1) {
    const child = paragraph.children[childIndex]
    if (child.kind === 'run') {
      if (currentRunIndex === runIndex) {
        return { kind: 'paragraph', childIndex, run: child }
      }
      currentRunIndex += 1
      continue
    }

    if (child.kind === 'hyperlink') {
      for (let runChildIndex = 0; runChildIndex < child.children.length; runChildIndex += 1) {
        const hyperlinkChild = child.children[runChildIndex]
        if (hyperlinkChild.kind !== 'run') {
          continue
        }
        if (currentRunIndex === runIndex) {
          return {
            kind: 'hyperlink',
            childIndex,
            runIndex: runChildIndex,
            hyperlink: child,
            run: hyperlinkChild,
          }
        }
        currentRunIndex += 1
      }
    }
  }

  return null
}

function insertNodesAroundRun<T extends InsertableNode>(
  children: ReadonlyArray<T>,
  targetIndex: number,
  run: Run,
  charOffset: number,
  nodes: ReadonlyArray<InsertableNode>,
): ReadonlyArray<T> {
  const runLength = getRunTextLength(run)
  const nextNodes = nodes as ReadonlyArray<T>

  if (charOffset <= 0) {
    return [...children.slice(0, targetIndex), ...nextNodes, ...children.slice(targetIndex)]
  }

  if (charOffset >= runLength) {
    return [...children.slice(0, targetIndex + 1), ...nextNodes, ...children.slice(targetIndex + 1)]
  }

  const split = splitRun(run, charOffset)
  const replacement: T[] = []
  if (split.before !== null) {
    replacement.push(split.before as T)
  }
  replacement.push(...nextNodes)
  if (split.after !== null) {
    replacement.push(split.after as T)
  }

  return [...children.slice(0, targetIndex), ...replacement, ...children.slice(targetIndex + 1)]
}

function getRunTextLength(run: Run): number {
  return run.children.reduce((total, child) => total + getRunChildLength(child), 0)
}

function getRunChildLength(child: RunChild): number {
  if (child.kind === 'text') {
    return child.value.length
  }
  if (child.kind === 'tab' || child.kind === 'break') {
    return 1
  }
  return 0
}

function splitRun(run: Run, charOffset: number): { readonly before: Run | null; readonly after: Run | null } {
  const beforeChildren: RunChild[] = []
  const afterChildren: RunChild[] = []
  let remaining = charOffset

  for (const child of run.children) {
    const length = getRunChildLength(child)
    if (remaining <= 0) {
      afterChildren.push(child)
      continue
    }

    if (length === 0) {
      beforeChildren.push(child)
      continue
    }

    if (remaining >= length) {
      beforeChildren.push(child)
      remaining -= length
      continue
    }

    if (child.kind === 'text') {
      const beforeText = child.value.slice(0, remaining)
      const afterText = child.value.slice(remaining)
      if (beforeText.length > 0) {
        beforeChildren.push(createTextNode(beforeText, child.preserveSpace))
      }
      if (afterText.length > 0) {
        afterChildren.push(createTextNode(afterText, child.preserveSpace))
      }
    } else {
      beforeChildren.push(child)
    }

    remaining = 0
  }

  return {
    before: beforeChildren.length > 0 ? { ...run, children: beforeChildren } : null,
    after: afterChildren.length > 0 ? { ...run, children: afterChildren } : null,
  }
}

function createTextNode(value: string, preserveSpace?: TextNode['preserveSpace']): TextNode {
  return {
    kind: 'text',
    value,
    ...(preserveSpace !== undefined ? { preserveSpace } : {}),
  }
}

function createCommentRange(id: string, boundary: CommentRange['boundary']): CommentRange {
  return {
    kind: 'comment-range',
    id,
    boundary,
  }
}

function createCommentReferenceRun(id: string): Run {
  const reference: CommentReference = {
    kind: 'comment-reference',
    id,
  }

  return {
    kind: 'run',
    children: [reference],
  }
}

function stripCommentIdsFromSection(section: Section, ids: ReadonlySet<string>): Section {
  return {
    ...section,
    blocks: section.blocks.map((block) => stripCommentIdsFromBlock(block, ids)),
  }
}

function stripCommentIdsFromBlock(block: Block, ids: ReadonlySet<string>): Block {
  if (block.kind === 'paragraph') {
    return stripCommentIdsFromParagraph(block, ids)
  }

  if (block.kind === 'table') {
    return {
      ...block,
      rows: block.rows.map((row) => stripCommentIdsFromRow(row, ids)),
    }
  }

  return block
}

function stripCommentIdsFromRow(row: TableChild, ids: ReadonlySet<string>): TableChild {
  if (row.kind !== 'table-row') {
    return row
  }

  return {
    ...row,
    cells: row.cells.map((cell) => stripCommentIdsFromCell(cell, ids)),
  }
}

function stripCommentIdsFromCell(cell: TableRowChild, ids: ReadonlySet<string>): TableRowChild {
  if (cell.kind !== 'table-cell') {
    return cell
  }

  return {
    ...cell,
    blocks: cell.blocks.map((block) => stripCommentIdsFromBlock(block, ids)),
  }
}

function stripCommentIdsFromParagraph(paragraph: Paragraph, ids: ReadonlySet<string>): Paragraph {
  return {
    ...paragraph,
    children: paragraph.children.flatMap((child) => stripCommentIdsFromParagraphChild(child, ids)),
  }
}

function stripCommentIdsFromParagraphChild(
  child: ParagraphChild,
  ids: ReadonlySet<string>,
): ReadonlyArray<ParagraphChild> {
  if ((child.kind === 'comment-range' || child.kind === 'comment-reference') && ids.has(child.id)) {
    return []
  }

  if (child.kind === 'run') {
    const nextChildren = child.children.filter(
      (runChild) => !(runChild.kind === 'comment-reference' && ids.has(runChild.id)),
    )
    return nextChildren.length === child.children.length ? [child] : [{ ...child, children: nextChildren }]
  }

  if (child.kind === 'hyperlink') {
    return [
      {
        ...child,
        children: child.children.flatMap((hyperlinkChild): ReadonlyArray<HyperlinkChild> => {
          if (
            (hyperlinkChild.kind === 'comment-range' || hyperlinkChild.kind === 'comment-reference') &&
            ids.has(hyperlinkChild.id)
          ) {
            return []
          }

          if (hyperlinkChild.kind === 'run') {
            const nextChildren = hyperlinkChild.children.filter(
              (runChild) => !(runChild.kind === 'comment-reference' && ids.has(runChild.id)),
            )
            return nextChildren.length === hyperlinkChild.children.length
              ? [hyperlinkChild]
              : [{ ...hyperlinkChild, children: nextChildren }]
          }

          return [hyperlinkChild]
        }),
      },
    ]
  }

  return [child]
}
