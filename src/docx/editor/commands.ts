import type {
  Block,
  Document,
  ParaProps,
  Paragraph,
  ParagraphChild,
  Run,
  RunProps,
  Section,
  Table,
  TableCell,
  TableRow,
  TextNode,
} from '../model'

import type {
  ApplyParaFormatCommand,
  ApplyRunFormatCommand,
  Command,
  DeleteRangeCommand,
  Position,
  Range,
} from './commandTypes'

type ResolvedParagraphPath = {
  readonly sectionIndex: number
  readonly blockPath: ReadonlyArray<number>
}

type EditableRun = {
  readonly run: Run
  readonly text: string
}

type BlocksUpdater = (
  blocks: ReadonlyArray<Block>,
  blockIndex: number,
) => ReadonlyArray<Block> | null

type RunRange = {
  readonly startOffset: number
  readonly endOffset: number
}

type RevisionResolutionCommand = Extract<
  Command,
  { kind: 'accept-revision' | 'reject-revision' }
>

type RevisionTarget = {
  readonly paragraphPath: ReadonlyArray<number>
  readonly childIndex: number
}

export function applyCommand(
  doc: Document,
  cmd: Command,
): {
  document: Document
  inverse: Command
} {
  switch (cmd.kind) {
    case 'insert-text':
      return applyInsertText(doc, cmd)
    case 'delete-range':
      return applyDeleteRange(doc, cmd)
    case 'insert-paragraph-break':
      return applyInsertParagraphBreak(doc, cmd)
    case 'apply-run-format':
      return applyRunFormat(doc, cmd)
    case 'apply-para-format':
      return applyParaFormat(doc, cmd)
    case 'apply-style':
      return applyStyle(doc, cmd)
    case 'insert-table': {
      // TODO: replace the no-op inverse stub once InsertTable is implemented.
      void createNoOpInsertText(cmd.at)
      throw new Error('not yet implemented')
    }
    case 'insert-hyperlink': {
      // TODO: replace the no-op inverse stub once InsertHyperlink is implemented.
      void createNoOpInsertText(cmd.range.anchor)
      throw new Error('not yet implemented')
    }
    case 'insert-list': {
      // TODO: replace the no-op inverse stub once InsertList is implemented.
      void createNoOpInsertText(createPosition(cmd.paragraphPaths[0] ?? [], 0, 0))
      throw new Error('not yet implemented')
    }
    case 'change-list-level': {
      // TODO: replace the no-op inverse stub once ChangeListLevel is implemented.
      void createNoOpInsertText(createPosition(cmd.paragraphPath, 0, 0))
      throw new Error('not yet implemented')
    }
    case 'accept-revision':
      return applyRevisionResolution(doc, cmd, 'accept')
    case 'reject-revision':
      return applyRevisionResolution(doc, cmd, 'reject')
    case 'accept-all-revisions':
      return applyAllRevisions(doc, 'accept')
    case 'reject-all-revisions':
      return applyAllRevisions(doc, 'reject')
  }
}

function applyRevisionResolution(
  doc: Document,
  cmd: RevisionResolutionCommand,
  mode: 'accept' | 'reject',
): {
  document: Document
  inverse: Command
} {
  const target = resolveRevisionTarget(doc, cmd)
  const paragraph = requireParagraph(doc, target.paragraphPath)
  const child = paragraph.children[target.childIndex]
  if (child === undefined || (child.kind !== 'ins-revision' && child.kind !== 'del-revision')) {
    throw new Error('Revision child not found at index')
  }

  const nextChildren = resolveRevisionChild(paragraph.children, target.childIndex, child, mode)
  const nextParagraph = cloneParagraph(paragraph, nextChildren)
  const nextDocument = replaceParagraphOrThrow(doc, target.paragraphPath, nextParagraph)

  return {
    document: nextDocument,
    inverse: createNoOpInsertText(createPosition(target.paragraphPath, 0, 0)),
  }
}

function resolveRevisionChild(
  children: ReadonlyArray<ParagraphChild>,
  childIndex: number,
  revision: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>,
  mode: 'accept' | 'reject',
): ReadonlyArray<ParagraphChild> {
  const keep =
    (revision.kind === 'ins-revision' && mode === 'accept') ||
    (revision.kind === 'del-revision' && mode === 'reject')

  const replacement: ReadonlyArray<ParagraphChild> = keep ? getRevisionChildren(revision) : []
  return freezeArray([
    ...children.slice(0, childIndex),
    ...replacement,
    ...children.slice(childIndex + 1),
  ])
}

function resolveRevisionTarget(
  doc: Document,
  cmd: RevisionResolutionCommand,
): RevisionTarget {
  if ('id' in cmd) {
    const resolved = findRevisionTargetById(doc, cmd.id)
    if (resolved === null) {
      throw new Error(`Revision not found for id ${cmd.id}`)
    }
    return resolved
  }

  return {
    paragraphPath: clonePath(cmd.paragraphPath),
    childIndex: cmd.childIndex,
  }
}

function findRevisionTargetById(
  doc: Document,
  id: string,
): RevisionTarget | null {
  for (let sectionIndex = 0; sectionIndex < doc.sections.length; sectionIndex += 1) {
    const section = doc.sections[sectionIndex]
    if (section === undefined) {
      continue
    }

    const target = findRevisionTargetInBlocks(section.blocks, [sectionIndex], id)
    if (target !== null) {
      return target
    }
  }

  return null
}

function findRevisionTargetInBlocks(
  blocks: ReadonlyArray<Block>,
  pathPrefix: ReadonlyArray<number>,
  id: string,
): RevisionTarget | null {
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
    const block = blocks[blockIndex]
    if (block === undefined) {
      continue
    }

    const blockPath = [...pathPrefix, blockIndex]
    if (block.kind === 'paragraph') {
      const childIndex = block.children.findIndex(
        (child) =>
          (child.kind === 'ins-revision' || child.kind === 'del-revision') &&
          child.id === id,
      )

      if (childIndex >= 0) {
        return {
          paragraphPath: freezeArray(blockPath),
          childIndex,
        }
      }

      continue
    }

    if (block.kind !== 'table') {
      continue
    }

    for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex += 1) {
      const row = block.rows[rowIndex]
      if (row === undefined || row.kind !== 'table-row') {
        continue
      }

      for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
        const cell = row.cells[cellIndex]
        if (cell === undefined || cell.kind !== 'table-cell') {
          continue
        }

        const target = findRevisionTargetInBlocks(
          cell.blocks,
          [...blockPath, rowIndex, cellIndex],
          id,
        )
        if (target !== null) {
          return target
        }
      }
    }
  }

  return null
}

function getRevisionChildren(
  revision: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>,
): ReadonlyArray<ParagraphChild> {
  return revision.children as ReadonlyArray<ParagraphChild>
}

function applyAllRevisions(
  doc: Document,
  mode: 'accept' | 'reject',
): {
  document: Document
  inverse: Command
} {
  let nextDocument = doc

  for (let sectionIdx = 0; sectionIdx < nextDocument.sections.length; sectionIdx += 1) {
    nextDocument = walkBlocksResolveRevisions(nextDocument, sectionIdx, mode)
  }

  return {
    document: nextDocument,
    inverse: { kind: mode === 'accept' ? 'reject-all-revisions' : 'accept-all-revisions' },
  }
}

function walkBlocksResolveRevisions(
  doc: Document,
  sectionIndex: number,
  mode: 'accept' | 'reject',
): Document {
  const section = doc.sections[sectionIndex]
  if (section === undefined) {
    return doc
  }

  let mutated = false
  const nextBlocks = section.blocks.map((block): Block => {
    if (block.kind !== 'paragraph') {
      return block
    }

    const resolved = resolveAllRevisionsInChildren(block.children, mode)
    if (resolved === block.children) {
      return block
    }
    mutated = true
    return cloneParagraph(block, resolved)
  })

  if (!mutated) {
    return doc
  }

  const nextSection = cloneSection(section, freezeArray(nextBlocks))
  const nextSections = replaceArrayItem(doc.sections, sectionIndex, nextSection)
  return cloneDocument(doc, nextSections)
}

function resolveAllRevisionsInChildren(
  children: ReadonlyArray<ParagraphChild>,
  mode: 'accept' | 'reject',
): ReadonlyArray<ParagraphChild> {
  let mutated = false
  const next: ParagraphChild[] = []

  for (const child of children) {
    if (child.kind !== 'ins-revision' && child.kind !== 'del-revision') {
      next.push(child)
      continue
    }

    mutated = true
    const keep =
      (child.kind === 'ins-revision' && mode === 'accept') ||
      (child.kind === 'del-revision' && mode === 'reject')
    if (keep) {
      next.push(...child.children)
    }
  }

  return mutated ? freezeArray(next) : children
}

export function findParagraph(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
): Paragraph | null {
  const resolvedPath = resolveParagraphPath(doc, paragraphPath)
  if (resolvedPath === null) {
    return null
  }

  const section = doc.sections[resolvedPath.sectionIndex]
  if (section === undefined) {
    return null
  }

  return findParagraphInBlocks(section.blocks, resolvedPath.blockPath)
}

export function replaceParagraph(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  newPara: Paragraph,
): Document {
  return (
    updateDocumentAtParagraphPath(doc, paragraphPath, (blocks, blockIndex) => {
      const block = blocks[blockIndex]
      if (block === undefined || block.kind !== 'paragraph') {
        return null
      }

      return replaceArrayItem(blocks, blockIndex, newPara)
    }) ?? doc
  )
}

function applyInsertText(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-text' }>,
): {
  document: Document
  inverse: Command
} {
  if (cmd.text.length === 0) {
    return {
      document: doc,
      inverse: createDeleteRangeCommand(cmd.at, cmd.at),
    }
  }

  const paragraph = requireParagraph(doc, cmd.at.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const originalRuns = editableRuns.map((entry) => entry.run)

  if (editableRuns.length === 0) {
    if (cmd.at.runIndex !== 0 || cmd.at.charOffset !== 0) {
      throw new Error('InsertText position is outside the paragraph')
    }

    const insertedRun = createRunWithText(undefined, cmd.text)
    const nextParagraph = cloneParagraph(paragraph, [insertedRun])
    const nextDocument = replaceParagraphOrThrow(doc, cmd.at.paragraphPath, nextParagraph)
    const start = createPosition(cmd.at.paragraphPath, 0, 0)
    const end = createPosition(cmd.at.paragraphPath, 0, cmd.text.length)

    return {
      document: nextDocument,
      inverse: createDeleteRangeCommand(start, end),
    }
  }

  const target = resolveInsertTarget(editableRuns, cmd.at)
  const nextRuns = originalRuns.slice()
  const current = editableRuns[target.runIndex]
  const nextText =
    current.text.slice(0, target.charOffset) +
    cmd.text +
    current.text.slice(target.charOffset)

  nextRuns[target.runIndex] = createRunLike(current.run, nextText, current.run.props)

  const nextParagraph = cloneParagraph(paragraph, nextRuns)
  const nextDocument = replaceParagraphOrThrow(doc, cmd.at.paragraphPath, nextParagraph)
  const start = createPosition(cmd.at.paragraphPath, target.runIndex, target.charOffset)
  const end = createPosition(cmd.at.paragraphPath, target.runIndex, target.charOffset + cmd.text.length)

  return {
    document: nextDocument,
    inverse: createDeleteRangeCommand(start, end),
  }
}

function applyDeleteRange(
  doc: Document,
  cmd: DeleteRangeCommand,
): {
  document: Document
  inverse: Command
} {
  const range = normalizeRange(doc, cmd.range)
  if (comparePositions(doc, range.anchor, range.focus) === 0) {
    return {
      document: doc,
      inverse: createNoOpInsertText(range.anchor),
    }
  }

  if (sameParagraphPath(doc, range.anchor.paragraphPath, range.focus.paragraphPath)) {
    return applyDeleteWithinParagraph(doc, range)
  }

  return applyDeleteParagraphBreak(doc, range)
}

function applyDeleteWithinParagraph(
  doc: Document,
  range: Range,
): {
  document: Document
  inverse: Command
} {
  const paragraph = requireParagraph(doc, range.anchor.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)

  if (range.anchor.runIndex !== range.focus.runIndex) {
    throw new Error('DeleteRange across multiple runs is not yet implemented')
  }

  const run = editableRuns[range.anchor.runIndex]
  if (run === undefined) {
    throw new Error('DeleteRange position is outside the paragraph')
  }

  if (
    range.anchor.charOffset < 0 ||
    range.focus.charOffset < range.anchor.charOffset ||
    range.focus.charOffset > run.text.length
  ) {
    throw new Error('DeleteRange position is outside the run')
  }

  const deletedText = run.text.slice(range.anchor.charOffset, range.focus.charOffset)
  const nextText =
    run.text.slice(0, range.anchor.charOffset) + run.text.slice(range.focus.charOffset)

  const nextRuns = editableRuns.map((entry) => entry.run).slice()
  nextRuns[range.anchor.runIndex] = createRunLike(run.run, nextText, run.run.props)

  const nextParagraph = cloneParagraph(paragraph, nextRuns)
  const nextDocument = replaceParagraphOrThrow(doc, range.anchor.paragraphPath, nextParagraph)

  return {
    document: nextDocument,
    inverse: {
      kind: 'insert-text',
      at: createPosition(
        range.anchor.paragraphPath,
        range.anchor.runIndex,
        range.anchor.charOffset,
      ),
      text: deletedText,
    },
  }
}

function applyDeleteParagraphBreak(
  doc: Document,
  range: Range,
): {
  document: Document
  inverse: Command
} {
  const startPath = resolveParagraphPath(doc, range.anchor.paragraphPath)
  const endPath = resolveParagraphPath(doc, range.focus.paragraphPath)
  if (startPath === null || endPath === null || !isNextSiblingPath(startPath, endPath)) {
    throw new Error('DeleteRange across multiple paragraphs is not yet implemented')
  }

  const firstParagraph = requireParagraph(doc, range.anchor.paragraphPath)
  const secondParagraph = requireParagraph(doc, range.focus.paragraphPath)
  const firstRuns = requireEditableRuns(firstParagraph)
  const secondRuns = requireEditableRuns(secondParagraph)

  if (!isParagraphEndPosition(firstRuns, range.anchor) || !isParagraphStartPosition(range.focus)) {
    throw new Error('DeleteRange only supports deleting a paragraph break at paragraph boundaries')
  }

  const mergedRuns = mergeAdjacentRuns([
    ...firstRuns.map((entry) => entry.run),
    ...secondRuns.map((entry) => entry.run),
  ])
  const mergedParagraph = cloneParagraph(firstParagraph, mergedRuns)
  const nextDocument = mergeParagraphWithNext(doc, range.anchor.paragraphPath, mergedParagraph)
  const inverseAt = createParagraphEndPosition(range.anchor.paragraphPath, firstRuns)

  return {
    document: nextDocument,
    inverse: {
      kind: 'insert-paragraph-break',
      at: inverseAt,
    },
  }
}

function applyInsertParagraphBreak(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-paragraph-break' }>,
): {
  document: Document
  inverse: Command
} {
  const paragraph = requireParagraph(doc, cmd.at.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const split = splitRunsAtPosition(editableRuns, cmd.at)

  const firstParagraph = cloneParagraph(paragraph, split.beforeRuns)
  const secondParagraph = cloneParagraph(paragraph, split.afterRuns)
  const nextDocument = insertParagraphAfter(doc, cmd.at.paragraphPath, firstParagraph, secondParagraph)
  const nextPath = incrementParagraphPath(cmd.at.paragraphPath)

  return {
    document: nextDocument,
    inverse: createDeleteRangeCommand(
      createParagraphEndPosition(cmd.at.paragraphPath, getEditableRunsOrEmpty(firstParagraph)),
      createPosition(nextPath, 0, 0),
    ),
  }
}

function applyRunFormat(
  doc: Document,
  cmd: ApplyRunFormatCommand,
): {
  document: Document
  inverse: Command
} {
  const range = normalizeRange(doc, cmd.range)
  if (comparePositions(doc, range.anchor, range.focus) === 0 || Object.keys(cmd.format).length === 0) {
    return {
      document: doc,
      inverse: createNoOpRunFormat(range),
    }
  }

  if (!sameParagraphPath(doc, range.anchor.paragraphPath, range.focus.paragraphPath)) {
    throw new Error('ApplyRunFormat across multiple paragraphs is not yet implemented')
  }

  const paragraph = requireParagraph(doc, range.anchor.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const offsets = getRangeOffsets(paragraph, range)
  const inverseFormat = collectInverseRunFormat(editableRuns, offsets, cmd.format)
  const formattedRuns = buildFormattedRuns(editableRuns, offsets, cmd.format)
  const mergedRuns = mergeAdjacentRuns(formattedRuns)
  const nextParagraph = cloneParagraph(paragraph, mergedRuns)
  const nextDocument = replaceParagraphOrThrow(doc, range.anchor.paragraphPath, nextParagraph)
  const nextEditableRuns = requireEditableRuns(nextParagraph)
  const nextRange = createRange(
    offsetToPosition(range.anchor.paragraphPath, nextEditableRuns, offsets.startOffset, 'forward'),
    offsetToPosition(range.anchor.paragraphPath, nextEditableRuns, offsets.endOffset, 'backward'),
  )

  return {
    document: nextDocument,
    inverse: {
      kind: 'apply-run-format',
      range: nextRange,
      format: inverseFormat,
    },
  }
}

function applyParaFormat(
  doc: Document,
  cmd: ApplyParaFormatCommand,
): {
  document: Document
  inverse: Command
} {
  if (cmd.paragraphPaths.length === 0 || Object.keys(cmd.format).length === 0) {
    return {
      document: doc,
      inverse: {
        kind: 'apply-para-format',
        paragraphPaths: cmd.paragraphPaths.map(clonePath),
        format: {},
      },
    }
  }

  const paragraphs = cmd.paragraphPaths.map((paragraphPath) => requireParagraph(doc, paragraphPath))
  const inverseFormat = collectInverseParagraphFormat(paragraphs, cmd.format)

  let nextDocument = doc
  for (const paragraphPath of cmd.paragraphPaths) {
    const paragraph = requireParagraph(nextDocument, paragraphPath)
    const nextParagraph = cloneParagraph(
      paragraph,
      paragraph.children,
      applyPropsPatch(paragraph.props, cmd.format),
    )
    nextDocument = replaceParagraphOrThrow(nextDocument, paragraphPath, nextParagraph)
  }

  return {
    document: nextDocument,
    inverse: {
      kind: 'apply-para-format',
      paragraphPaths: cmd.paragraphPaths.map(clonePath),
      format: inverseFormat,
    },
  }
}

function applyStyle(
  doc: Document,
  cmd: Extract<Command, { kind: 'apply-style' }>,
): {
  document: Document
  inverse: Command
} {
  const paragraph = requireParagraph(doc, cmd.paragraphPath)
  const nextParagraph = cloneParagraph(
    paragraph,
    paragraph.children,
    applyPropsPatch(paragraph.props, { pStyle: cmd.styleId }),
  )
  const nextDocument = replaceParagraphOrThrow(doc, cmd.paragraphPath, nextParagraph)

  return {
    document: nextDocument,
    inverse: {
      kind: 'apply-para-format',
      paragraphPaths: [clonePath(cmd.paragraphPath)],
      format: { pStyle: paragraph.props?.pStyle },
    },
  }
}

function requireParagraph(doc: Document, paragraphPath: ReadonlyArray<number>): Paragraph {
  const paragraph = findParagraph(doc, paragraphPath)
  if (paragraph === null) {
    throw new Error('Paragraph not found')
  }

  return paragraph
}

function replaceParagraphOrThrow(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  newParagraph: Paragraph,
): Document {
  const nextDocument = updateDocumentAtParagraphPath(doc, paragraphPath, (blocks, blockIndex) => {
    const block = blocks[blockIndex]
    if (block === undefined || block.kind !== 'paragraph') {
      return null
    }

    return replaceArrayItem(blocks, blockIndex, newParagraph)
  })

  if (nextDocument === null) {
    throw new Error('Paragraph not found')
  }

  return nextDocument
}

function insertParagraphAfter(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  currentParagraph: Paragraph,
  nextParagraph: Paragraph,
): Document {
  const nextDocument = updateDocumentAtParagraphPath(doc, paragraphPath, (blocks, blockIndex) => {
    const block = blocks[blockIndex]
    if (block === undefined || block.kind !== 'paragraph') {
      return null
    }

    return freezeArray([
      ...blocks.slice(0, blockIndex),
      currentParagraph,
      nextParagraph,
      ...blocks.slice(blockIndex + 1),
    ])
  })

  if (nextDocument === null) {
    throw new Error('Paragraph not found')
  }

  return nextDocument
}

function mergeParagraphWithNext(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  mergedParagraph: Paragraph,
): Document {
  const nextDocument = updateDocumentAtParagraphPath(doc, paragraphPath, (blocks, blockIndex) => {
    const firstBlock = blocks[blockIndex]
    const secondBlock = blocks[blockIndex + 1]

    if (
      firstBlock === undefined ||
      firstBlock.kind !== 'paragraph' ||
      secondBlock === undefined ||
      secondBlock.kind !== 'paragraph'
    ) {
      return null
    }

    return freezeArray([
      ...blocks.slice(0, blockIndex),
      mergedParagraph,
      ...blocks.slice(blockIndex + 2),
    ])
  })

  if (nextDocument === null) {
    throw new Error('Paragraph not found')
  }

  return nextDocument
}

function findParagraphInBlocks(
  blocks: ReadonlyArray<Block>,
  blockPath: ReadonlyArray<number>,
): Paragraph | null {
  if (blockPath.length === 0) {
    return null
  }

  const [blockIndex, ...rest] = blockPath
  const block = blocks[blockIndex]
  if (block === undefined) {
    return null
  }

  if (rest.length === 0) {
    return block.kind === 'paragraph' ? block : null
  }

  if (block.kind !== 'table' || rest.length < 3) {
    return null
  }

  const [rowIndex, cellIndex, ...childPath] = rest
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return null
  }

  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  return findParagraphInBlocks(cell.blocks, childPath)
}

function updateDocumentAtParagraphPath(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  update: BlocksUpdater,
): Document | null {
  const resolvedPath = resolveParagraphPath(doc, paragraphPath)
  if (resolvedPath === null) {
    return null
  }

  const section = doc.sections[resolvedPath.sectionIndex]
  if (section === undefined) {
    return null
  }

  const nextBlocks = updateBlocksAtPath(section.blocks, resolvedPath.blockPath, update)
  if (nextBlocks === null) {
    return null
  }

  if (nextBlocks === section.blocks) {
    return doc
  }

  const nextSection = cloneSection(section, nextBlocks)
  const nextSections = replaceArrayItem(doc.sections, resolvedPath.sectionIndex, nextSection)

  return cloneDocument(doc, nextSections)
}

function updateBlocksAtPath(
  blocks: ReadonlyArray<Block>,
  blockPath: ReadonlyArray<number>,
  update: BlocksUpdater,
): ReadonlyArray<Block> | null {
  if (blockPath.length === 0) {
    return null
  }

  const [blockIndex, ...rest] = blockPath
  const block = blocks[blockIndex]
  if (block === undefined) {
    return null
  }

  if (rest.length === 0) {
    return update(blocks, blockIndex)
  }

  if (block.kind !== 'table' || rest.length < 3) {
    return null
  }

  const [rowIndex, cellIndex, ...childPath] = rest
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return null
  }

  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  const nextCellBlocks = updateBlocksAtPath(cell.blocks, childPath, update)
  if (nextCellBlocks === null) {
    return null
  }

  if (nextCellBlocks === cell.blocks) {
    return blocks
  }

  const nextCell = cloneTableCell(cell, nextCellBlocks)
  const nextCells = replaceArrayItem(row.cells, cellIndex, nextCell)
  const nextRow = cloneTableRow(row, nextCells)
  const nextRows = replaceArrayItem(block.rows, rowIndex, nextRow)
  const nextTable = cloneTable(block, nextRows)

  return replaceArrayItem(blocks, blockIndex, nextTable)
}

function resolveParagraphPath(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
): ResolvedParagraphPath | null {
  if (paragraphPath.length === 0) {
    return null
  }

  if (paragraphPath.length === 1) {
    return {
      sectionIndex: 0,
      blockPath: clonePath(paragraphPath),
    }
  }

  const [sectionIndex, ...blockPath] = paragraphPath
  if (doc.sections[sectionIndex] !== undefined) {
    return {
      sectionIndex,
      blockPath: freezeArray(blockPath),
    }
  }

  return {
    sectionIndex: 0,
    blockPath: clonePath(paragraphPath),
  }
}

function requireEditableRuns(paragraph: Paragraph): ReadonlyArray<EditableRun> {
  const editableRuns = getEditableRuns(paragraph)
  if (editableRuns === null) {
    throw new Error('This command currently supports paragraphs with direct text runs only')
  }

  return editableRuns
}

function getEditableRunsOrEmpty(paragraph: Paragraph): ReadonlyArray<EditableRun> {
  return getEditableRuns(paragraph) ?? []
}

function getEditableRuns(paragraph: Paragraph): ReadonlyArray<EditableRun> | null {
  const editableRuns: EditableRun[] = []

  for (const child of paragraph.children) {
    if (child.kind !== 'run') {
      return null
    }

    const text = getRunText(child)
    if (text === null) {
      return null
    }

    editableRuns.push({ run: child, text })
  }

  return editableRuns
}

function getRunText(run: Run): string | null {
  let text = ''

  for (const child of run.children) {
    if (child.kind !== 'text') {
      return null
    }

    text += child.value
  }

  return text
}

function resolveInsertTarget(
  editableRuns: ReadonlyArray<EditableRun>,
  at: Position,
): {
  readonly runIndex: number
  readonly charOffset: number
} {
  if (at.runIndex < 0) {
    throw new Error('InsertText position is outside the paragraph')
  }

  if (at.runIndex === editableRuns.length) {
    if (at.charOffset !== 0) {
      throw new Error('InsertText position is outside the paragraph')
    }

    const lastIndex = editableRuns.length - 1
    return {
      runIndex: lastIndex,
      charOffset: editableRuns[lastIndex].text.length,
    }
  }

  const run = editableRuns[at.runIndex]
  if (run === undefined || at.charOffset < 0 || at.charOffset > run.text.length) {
    throw new Error('InsertText position is outside the paragraph')
  }

  return {
    runIndex: at.runIndex,
    charOffset: at.charOffset,
  }
}

function splitRunsAtPosition(
  editableRuns: ReadonlyArray<EditableRun>,
  at: Position,
): {
  readonly beforeRuns: ReadonlyArray<Run>
  readonly afterRuns: ReadonlyArray<Run>
} {
  const runs = editableRuns.map((entry) => entry.run)

  if (editableRuns.length === 0) {
    if (at.runIndex !== 0 || at.charOffset !== 0) {
      throw new Error('InsertParagraphBreak position is outside the paragraph')
    }

    return {
      beforeRuns: freezeArray<Run>([]),
      afterRuns: freezeArray<Run>([]),
    }
  }

  if (at.runIndex === editableRuns.length) {
    if (at.charOffset !== 0) {
      throw new Error('InsertParagraphBreak position is outside the paragraph')
    }

    return {
      beforeRuns: freezeArray(runs.slice()),
      afterRuns: freezeArray<Run>([]),
    }
  }

  const targetRun = editableRuns[at.runIndex]
  if (targetRun === undefined || at.charOffset < 0 || at.charOffset > targetRun.text.length) {
    throw new Error('InsertParagraphBreak position is outside the paragraph')
  }

  if (at.charOffset === 0) {
    return {
      beforeRuns: freezeArray(runs.slice(0, at.runIndex)),
      afterRuns: freezeArray(runs.slice(at.runIndex)),
    }
  }

  if (at.charOffset === targetRun.text.length) {
    return {
      beforeRuns: freezeArray(runs.slice(0, at.runIndex + 1)),
      afterRuns: freezeArray(runs.slice(at.runIndex + 1)),
    }
  }

  const leftRun = createRunLike(
    targetRun.run,
    targetRun.text.slice(0, at.charOffset),
    targetRun.run.props,
  )
  const rightRun = createRunLike(
    targetRun.run,
    targetRun.text.slice(at.charOffset),
    targetRun.run.props,
  )

  return {
    beforeRuns: freezeArray([...runs.slice(0, at.runIndex), leftRun]),
    afterRuns: freezeArray([rightRun, ...runs.slice(at.runIndex + 1)]),
  }
}

function normalizeRange(doc: Document, range: Range): Range {
  if (comparePositions(doc, range.anchor, range.focus) <= 0) {
    return createRange(range.anchor, range.focus)
  }

  return createRange(range.focus, range.anchor)
}

function sameParagraphPath(
  doc: Document,
  left: ReadonlyArray<number>,
  right: ReadonlyArray<number>,
): boolean {
  const leftPath = normalizePathForComparison(doc, left)
  const rightPath = normalizePathForComparison(doc, right)

  return compareNumberPaths(leftPath, rightPath) === 0
}

function comparePositions(doc: Document, left: Position, right: Position): number {
  const pathCompare = compareNumberPaths(
    normalizePathForComparison(doc, left.paragraphPath),
    normalizePathForComparison(doc, right.paragraphPath),
  )

  if (pathCompare !== 0) {
    return pathCompare
  }

  if (left.runIndex !== right.runIndex) {
    return left.runIndex - right.runIndex
  }

  return left.charOffset - right.charOffset
}

function normalizePathForComparison(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
): ReadonlyArray<number> {
  const resolvedPath = resolveParagraphPath(doc, paragraphPath)
  if (resolvedPath === null) {
    return clonePath(paragraphPath)
  }

  return freezeArray([resolvedPath.sectionIndex, ...resolvedPath.blockPath])
}

function compareNumberPaths(
  left: ReadonlyArray<number>,
  right: ReadonlyArray<number>,
): number {
  const maxLength = Math.max(left.length, right.length)

  for (let index = 0; index < maxLength; index += 1) {
    const leftValue = left[index]
    const rightValue = right[index]

    if (leftValue === undefined) {
      return -1
    }

    if (rightValue === undefined) {
      return 1
    }

    if (leftValue !== rightValue) {
      return leftValue - rightValue
    }
  }

  return 0
}

function getRangeOffsets(paragraph: Paragraph, range: Range): RunRange {
  const startOffset = positionToOffset(paragraph, range.anchor)
  const endOffset = positionToOffset(paragraph, range.focus)

  return { startOffset, endOffset }
}

function positionToOffset(paragraph: Paragraph, position: Position): number {
  const editableRuns = requireEditableRuns(paragraph)

  if (editableRuns.length === 0) {
    if (position.runIndex === 0 && position.charOffset === 0) {
      return 0
    }

    throw new Error('Position is outside the paragraph')
  }

  let offset = 0
  for (let index = 0; index < editableRuns.length; index += 1) {
    const run = editableRuns[index]
    if (index === position.runIndex) {
      if (position.charOffset < 0 || position.charOffset > run.text.length) {
        throw new Error('Position is outside the run')
      }

      return offset + position.charOffset
    }

    offset += run.text.length
  }

  if (position.runIndex === editableRuns.length && position.charOffset === 0) {
    return offset
  }

  throw new Error('Position is outside the paragraph')
}

function offsetToPosition(
  paragraphPath: ReadonlyArray<number>,
  editableRuns: ReadonlyArray<EditableRun>,
  offset: number,
  bias: 'forward' | 'backward',
): Position {
  if (editableRuns.length === 0) {
    return createPosition(paragraphPath, 0, 0)
  }

  let currentOffset = 0
  for (let index = 0; index < editableRuns.length; index += 1) {
    const run = editableRuns[index]
    const nextOffset = currentOffset + run.text.length

    if (offset < nextOffset) {
      return createPosition(paragraphPath, index, offset - currentOffset)
    }

    if (offset === currentOffset && bias === 'forward') {
      return createPosition(paragraphPath, index, 0)
    }

    if (offset === nextOffset) {
      if (bias === 'forward' && index < editableRuns.length - 1) {
        return createPosition(paragraphPath, index + 1, 0)
      }

      return createPosition(paragraphPath, index, run.text.length)
    }

    currentOffset = nextOffset
  }

  const lastIndex = editableRuns.length - 1
  return createPosition(paragraphPath, lastIndex, editableRuns[lastIndex].text.length)
}

function buildFormattedRuns(
  editableRuns: ReadonlyArray<EditableRun>,
  offsets: RunRange,
  format: Partial<RunProps>,
): ReadonlyArray<Run> {
  const nextRuns: Run[] = []
  let currentOffset = 0

  for (const entry of editableRuns) {
    const runStart = currentOffset
    const runEnd = currentOffset + entry.text.length

    if (runEnd <= offsets.startOffset || runStart >= offsets.endOffset) {
      nextRuns.push(entry.run)
      currentOffset = runEnd
      continue
    }

    const localStart = Math.max(0, offsets.startOffset - runStart)
    const localEnd = Math.min(entry.text.length, offsets.endOffset - runStart)

    if (localStart > 0) {
      nextRuns.push(createRunLike(entry.run, entry.text.slice(0, localStart), entry.run.props))
    }

    nextRuns.push(
      createRunLike(
        entry.run,
        entry.text.slice(localStart, localEnd),
        applyPropsPatch(entry.run.props, format),
      ),
    )

    if (localEnd < entry.text.length) {
      nextRuns.push(createRunLike(entry.run, entry.text.slice(localEnd), entry.run.props))
    }

    currentOffset = runEnd
  }

  return freezeArray(nextRuns)
}

function collectInverseRunFormat(
  editableRuns: ReadonlyArray<EditableRun>,
  offsets: RunRange,
  format: Partial<RunProps>,
): Partial<RunProps> {
  const targetProps: Array<RunProps | undefined> = []
  let currentOffset = 0

  for (const entry of editableRuns) {
    const runStart = currentOffset
    const runEnd = currentOffset + entry.text.length

    if (runEnd > offsets.startOffset && runStart < offsets.endOffset) {
      targetProps.push(entry.run.props)
    }

    currentOffset = runEnd
  }

  return collectHomogeneousPatchValues(targetProps, format)
}

function collectInverseParagraphFormat(
  paragraphs: ReadonlyArray<Paragraph>,
  format: Partial<ParaProps>,
): Partial<ParaProps> {
  return collectHomogeneousPatchValues(
    paragraphs.map((paragraph) => paragraph.props),
    format,
  )
}

function collectHomogeneousPatchValues<T extends object>(
  sources: ReadonlyArray<T | undefined>,
  format: Partial<T>,
): Partial<T> {
  const inversePatch: Partial<T> = {}

  for (const key of getPatchKeys(format)) {
    const values = sources.map((source) => source?.[key])
    if (!allValuesMatch(values)) {
      throw new Error('Cannot invert a format command over heterogeneous existing values')
    }

    inversePatch[key] = values[0]
  }

  return inversePatch
}

function getPatchKeys<T extends object>(format: Partial<T>): ReadonlyArray<keyof T> {
  return Object.keys(format) as Array<keyof T>
}

function allValuesMatch(values: ReadonlyArray<unknown>): boolean {
  if (values.length <= 1) {
    return true
  }

  const first = values[0]
  return values.every((value) => isDeepEqual(value, first))
}

function applyPropsPatch<T extends object>(
  base: T | undefined,
  patch: Partial<T>,
): T | undefined {
  const nextEntries = new Map<string, unknown>()

  if (base !== undefined) {
    for (const [key, value] of Object.entries(base)) {
      nextEntries.set(key, value)
    }
  }

  for (const key of getPatchKeys(patch)) {
    const value = patch[key]
    if (value === undefined) {
      nextEntries.delete(String(key))
      continue
    }

    nextEntries.set(String(key), value)
  }

  if (nextEntries.size === 0) {
    return undefined
  }

  const result: Record<string, unknown> = {}
  for (const [key, value] of nextEntries.entries()) {
    result[key] = value
  }

  return Object.freeze(result) as T
}

function mergeAdjacentRuns(runs: ReadonlyArray<Run>): ReadonlyArray<Run> {
  const merged: Run[] = []

  for (const run of runs) {
    const text = getRunText(run)
    if (text === null) {
      throw new Error('This command currently supports text-only runs')
    }

    const previous = merged[merged.length - 1]
    if (previous === undefined || !sameRunProps(previous.props, run.props)) {
      merged.push(run)
      continue
    }

    const previousText = getRunText(previous)
    if (previousText === null) {
      throw new Error('This command currently supports text-only runs')
    }

    merged[merged.length - 1] = createRunLike(previous, previousText + text, previous.props)
  }

  return freezeArray(merged)
}

function sameRunProps(left: RunProps | undefined, right: RunProps | undefined): boolean {
  return isDeepEqual(left, right)
}

function isDeepEqual(left: unknown, right: unknown): boolean {
  if (left === undefined || right === undefined) {
    return left === right
  }

  return JSON.stringify(left) === JSON.stringify(right)
}

function isParagraphEndPosition(
  editableRuns: ReadonlyArray<EditableRun>,
  position: Position,
): boolean {
  if (editableRuns.length === 0) {
    return position.runIndex === 0 && position.charOffset === 0
  }

  const lastIndex = editableRuns.length - 1
  return (
    (position.runIndex === editableRuns.length && position.charOffset === 0) ||
    (position.runIndex === lastIndex && position.charOffset === editableRuns[lastIndex].text.length)
  )
}

function isParagraphStartPosition(position: Position): boolean {
  return position.runIndex === 0 && position.charOffset === 0
}

function isNextSiblingPath(
  left: ResolvedParagraphPath,
  right: ResolvedParagraphPath,
): boolean {
  if (left.sectionIndex !== right.sectionIndex || left.blockPath.length !== right.blockPath.length) {
    return false
  }

  if (left.blockPath.length === 0) {
    return false
  }

  for (let index = 0; index < left.blockPath.length - 1; index += 1) {
    if (left.blockPath[index] !== right.blockPath[index]) {
      return false
    }
  }

  return right.blockPath[left.blockPath.length - 1] === left.blockPath[left.blockPath.length - 1] + 1
}

function createParagraphEndPosition(
  paragraphPath: ReadonlyArray<number>,
  editableRuns: ReadonlyArray<EditableRun>,
): Position {
  if (editableRuns.length === 0) {
    return createPosition(paragraphPath, 0, 0)
  }

  const lastIndex = editableRuns.length - 1
  return createPosition(paragraphPath, lastIndex, editableRuns[lastIndex].text.length)
}

function createPosition(
  paragraphPath: ReadonlyArray<number>,
  runIndex: number,
  charOffset: number,
): Position {
  return {
    paragraphPath: clonePath(paragraphPath),
    runIndex,
    charOffset,
  }
}

function createRange(anchor: Position, focus: Position): Range {
  return {
    anchor: createPosition(anchor.paragraphPath, anchor.runIndex, anchor.charOffset),
    focus: createPosition(focus.paragraphPath, focus.runIndex, focus.charOffset),
  }
}

function createDeleteRangeCommand(anchor: Position, focus: Position): DeleteRangeCommand {
  return {
    kind: 'delete-range',
    range: createRange(anchor, focus),
  }
}

function createNoOpInsertText(at: Position): Command {
  return {
    kind: 'insert-text',
    at: createPosition(at.paragraphPath, at.runIndex, at.charOffset),
    text: '',
  }
}

function createNoOpRunFormat(range: Range): Command {
  return {
    kind: 'apply-run-format',
    range: createRange(range.anchor, range.focus),
    format: {},
  }
}

function incrementParagraphPath(paragraphPath: ReadonlyArray<number>): ReadonlyArray<number> {
  if (paragraphPath.length === 0) {
    throw new Error('Cannot increment an empty paragraph path')
  }

  const nextPath = paragraphPath.slice()
  nextPath[nextPath.length - 1] += 1
  return freezeArray(nextPath)
}

function clonePath(path: ReadonlyArray<number>): ReadonlyArray<number> {
  return freezeArray(path.slice())
}

function cloneParagraph(
  paragraph: Paragraph,
  children: ReadonlyArray<Run> | ReadonlyArray<ParagraphChild>,
  props?: ParaProps,
): Paragraph {
  const nextProps = arguments.length >= 3 ? props : paragraph.props

  return Object.freeze({
    kind: 'paragraph',
    ...(nextProps !== undefined ? { props: nextProps } : {}),
    children: freezeArray(children.slice()),
  })
}

function cloneSection(section: Section, blocks: ReadonlyArray<Block>): Section {
  return Object.freeze({
    kind: 'section',
    props: section.props,
    blocks,
  })
}

function cloneTable(table: Table, rows: ReadonlyArray<Table['rows'][number]>): Table {
  return Object.freeze({
    kind: 'table',
    ...(table.props !== undefined ? { props: table.props } : {}),
    rows,
  })
}

function cloneTableRow(
  row: TableRow,
  cells: ReadonlyArray<TableRow['cells'][number]>,
): TableRow {
  return Object.freeze({
    kind: 'table-row',
    ...(row.props !== undefined ? { props: row.props } : {}),
    cells,
  })
}

function cloneTableCell(cell: TableCell, blocks: ReadonlyArray<Block>): TableCell {
  return Object.freeze({
    kind: 'table-cell',
    ...(cell.props !== undefined ? { props: cell.props } : {}),
    blocks,
  })
}

function cloneDocument(doc: Document, sections: ReadonlyArray<Section>): Document {
  return Object.freeze({
    ...doc,
    sections,
  })
}

function createRunLike(run: Run, text: string, props: RunProps | undefined): Run {
  const currentText = getRunText(run)
  if (currentText === text && sameRunProps(run.props, props)) {
    return run
  }

  return createRunWithText(props, text)
}

function createRunWithText(props: RunProps | undefined, text: string): Run {
  return Object.freeze({
    kind: 'run',
    ...(props !== undefined ? { props } : {}),
    children: freezeArray([createTextNode(text)]),
  })
}

function createTextNode(value: string): TextNode {
  return Object.freeze({
    kind: 'text',
    value,
    ...(/^[\s]|[\s]$/.test(value) ? { preserveSpace: true } : {}),
  })
}

function replaceArrayItem<T>(
  items: ReadonlyArray<T>,
  index: number,
  value: T,
): ReadonlyArray<T> {
  if (items[index] === value) {
    return items
  }

  const nextItems = items.slice()
  nextItems[index] = value
  return freezeArray(nextItems)
}

function freezeArray<T>(items: ReadonlyArray<T>): ReadonlyArray<T> {
  return Object.freeze([...items])
}
