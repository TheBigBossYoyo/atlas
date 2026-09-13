import type { ParaProps, RunProps } from '../model'

export type Position = {
  readonly paragraphPath: ReadonlyArray<number>
  readonly runIndex: number
  readonly charOffset: number
}

export type Range = {
  readonly anchor: Position
  readonly focus: Position
}

export type InsertTextCommand = {
  readonly kind: 'insert-text'
  readonly at: Position
  readonly text: string
}

export type DeleteRangeCommand = {
  readonly kind: 'delete-range'
  readonly range: Range
}

export type InsertParagraphBreakCommand = {
  readonly kind: 'insert-paragraph-break'
  readonly at: Position
}

export type ApplyRunFormatCommand = {
  readonly kind: 'apply-run-format'
  readonly range: Range
  readonly format: Partial<RunProps>
}

export type ApplyParaFormatCommand = {
  readonly kind: 'apply-para-format'
  readonly paragraphPaths: ReadonlyArray<ReadonlyArray<number>>
  readonly format: Partial<ParaProps>
}

export type InsertTableCommand = {
  readonly kind: 'insert-table'
  readonly at: Position
  readonly rows: number
  readonly cols: number
}

export type InsertHyperlinkCommand = {
  readonly kind: 'insert-hyperlink'
  readonly range: Range
  readonly url: string
}

export type ApplyStyleCommand = {
  readonly kind: 'apply-style'
  readonly paragraphPath: ReadonlyArray<number>
  readonly styleId: string
}

export type InsertListCommand = {
  readonly kind: 'insert-list'
  readonly paragraphPaths: ReadonlyArray<ReadonlyArray<number>>
  readonly numId: number
  readonly level: number
}

export type ChangeListLevelCommand = {
  readonly kind: 'change-list-level'
  readonly paragraphPath: ReadonlyArray<number>
  readonly delta: 1 | -1
}

type RevisionTarget =
  | {
      readonly id: string
    }
  | {
      readonly paragraphPath: ReadonlyArray<number>
      readonly childIndex: number
    }

export type AcceptRevisionCommand = {
  readonly kind: 'accept-revision'
} & RevisionTarget

export type RejectRevisionCommand = {
  readonly kind: 'reject-revision'
} & RevisionTarget

export type AcceptAllRevisionsCommand = {
  readonly kind: 'accept-all-revisions'
}

export type RejectAllRevisionsCommand = {
  readonly kind: 'reject-all-revisions'
}

export type Command =
  | InsertTextCommand
  | DeleteRangeCommand
  | InsertParagraphBreakCommand
  | ApplyRunFormatCommand
  | ApplyParaFormatCommand
  | InsertTableCommand
  | InsertHyperlinkCommand
  | ApplyStyleCommand
  | InsertListCommand
  | ChangeListLevelCommand
  | AcceptRevisionCommand
  | RejectRevisionCommand
  | AcceptAllRevisionsCommand
  | RejectAllRevisionsCommand

export function acceptRevision(id: string): AcceptRevisionCommand {
  return {
    kind: 'accept-revision',
    id,
  }
}

export function rejectRevision(id: string): RejectRevisionCommand {
  return {
    kind: 'reject-revision',
    id,
  }
}

export function acceptAllRevisions(): AcceptAllRevisionsCommand {
  return {
    kind: 'accept-all-revisions',
  }
}

export function rejectAllRevisions(): RejectAllRevisionsCommand {
  return {
    kind: 'reject-all-revisions',
  }
}
