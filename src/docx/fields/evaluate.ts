/**
 * Atlas — per-field-type evaluators (DEFER-5 / DXS-20)
 *
 * Each evaluator computes the plain-text replacement for one field type's
 * cached display content from its parsed instruction + a
 * `FieldEvaluationContext`. `evaluateFieldText` is the single dispatcher
 * `updateFields.ts` calls; it returns `undefined` for a field type Atlas
 * doesn't evaluate (including `TOC`, which isn't a single-text-value field
 * at all — see `toc.ts`) or one that's evaluable in principle but missing
 * the context data it needs right now (e.g. PAGE before layout has run) —
 * either way, "leave the field's cached result exactly as it was" is the
 * correct, honest fallback per `Field`'s own doc comment.
 */
import type { Field, FieldType } from '../model'
import { resolvePageField } from '../layout/pageFields'

import { DEFAULT_DATE_PICTURE, DEFAULT_TIME_PICTURE, formatDatePicture } from './dateFormat'
import { parseFieldInstruction, type ParsedFieldInstruction } from './instruction'
import type { FieldEvaluationContext } from './types'

export function evaluateFieldText(
  field: Field,
  paragraphPath: ReadonlyArray<number>,
  context: FieldEvaluationContext,
): string | undefined {
  const instruction = parseFieldInstruction(field.instruction)

  switch (field.fieldType) {
    case 'DATE':
      return evaluateDateTimeField(instruction, 'DATE', context)
    case 'TIME':
      return evaluateDateTimeField(instruction, 'TIME', context)
    case 'AUTHOR':
      return context.author
    case 'TITLE':
      return context.title
    case 'REF':
      return evaluateRefField(instruction, context)
    case 'PAGEREF':
      return evaluatePageRefField(instruction, context)
    case 'SEQ':
      return evaluateSeqField(instruction, context)
    case 'NUMPAGES':
      return evaluateNumPagesField(field.instruction, context)
    case 'PAGE':
      return evaluatePageField(field.instruction, paragraphPath, context)
    case 'HYPERLINK':
      // A HYPERLINK field's display text is user-authored (Word never
      // regenerates it from the target URL either) — see
      // `parseHyperlinkField` for extracting its structural target/switches
      // instead, for a caller that needs those.
      return undefined
    case 'TOC':
      return undefined
    case 'unknown':
      return undefined
    default:
      // Every FieldType is handled above; this only guards a future model
      // change that adds a variant here without also updating this switch —
      // "leave it unchanged" is the same safe fallback as every other
      // not-evaluable case, not a crash.
      return undefined
  }
}

function evaluateDateTimeField(
  instruction: ParsedFieldInstruction,
  kind: 'DATE' | 'TIME',
  context: FieldEvaluationContext,
): string {
  const now = context.now ?? new Date()
  const picture = instruction.switches.get('@')
  const format =
    typeof picture === 'string'
      ? picture
      : kind === 'DATE'
        ? DEFAULT_DATE_PICTURE
        : DEFAULT_TIME_PICTURE
  return formatDatePicture(now, format)
}

/**
 * `REF`'s default (no switches, or `\h`) mode displays the bookmarked
 * text. Word's `\p` (relative position, "above"/"below") and `\n`
 * (paragraph number) modes need paragraph-numbering context this scope
 * doesn't model — an instruction using either falls back to plain
 * bookmark text rather than fabricating a wrong position/number, an
 * honest approximation rather than silent wrong output.
 */
function evaluateRefField(
  instruction: ParsedFieldInstruction,
  context: FieldEvaluationContext,
): string | undefined {
  const bookmarkName = instruction.arguments[0]
  return bookmarkName !== undefined ? context.bookmarkText.get(bookmarkName) : undefined
}

function evaluatePageRefField(
  instruction: ParsedFieldInstruction,
  context: FieldEvaluationContext,
): string | undefined {
  const bookmarkName = instruction.arguments[0]
  if (bookmarkName === undefined) {
    return undefined
  }
  const page = context.bookmarkPage?.get(bookmarkName)
  return page !== undefined ? String(page) : undefined
}

/**
 * `\c` repeats the sequence's current value without incrementing (0 if the
 * sequence has never been set); `\r n` resets it to `n` (and that becomes
 * the field's displayed value, matching Word); otherwise the sequence
 * increments by 1. `\s` (heading-level-scoped restart) isn't modeled —
 * Atlas has no per-heading-level reset scope for a sequence counter — so a
 * `\s` switch is accepted (doesn't break parsing) but has no effect beyond
 * the plain-increment default.
 */
function evaluateSeqField(
  instruction: ParsedFieldInstruction,
  context: FieldEvaluationContext,
): string | undefined {
  const sequenceName = instruction.arguments[0]
  if (sequenceName === undefined) {
    return undefined
  }

  const current = context.sequenceCounters.get(sequenceName) ?? 0
  const resetTo = instruction.switches.get('r')
  const resetValue = typeof resetTo === 'string' ? Number.parseInt(resetTo, 10) : undefined

  const next = resetValue !== undefined && Number.isFinite(resetValue)
    ? resetValue
    : instruction.switches.has('c')
      ? current
      : current + 1

  context.sequenceCounters.set(sequenceName, next)
  return String(next)
}

/**
 * Delegates the actual numeric formatting to `layout/pageFields.ts`'s
 * `resolvePageField` (owned by wave3/docx-pagination) rather than
 * duplicating its `\* ROMAN`/`\* roman`/`\* ALPHABETIC`/`\* alphabetic`
 * switch handling here — the two branches independently built a PAGE/
 * NUMPAGES resolver (this evaluator wired to the real Field/context model,
 * `resolvePageField` a pure instruction-string formatter that was never
 * wired to a parsed field), and unifying on `resolvePageField` for the
 * numeric-format piece means this evaluator gets that support for free
 * instead of leaving it as dead, unreferenced code.
 */
function evaluateNumPagesField(rawInstruction: string, context: FieldEvaluationContext): string | undefined {
  if (context.pageCount === undefined) {
    return undefined
  }
  return (
    resolvePageField(rawInstruction, { pageNumber: 1, totalPages: context.pageCount }) ??
    String(context.pageCount)
  )
}

function evaluatePageField(
  rawInstruction: string,
  paragraphPath: ReadonlyArray<number>,
  context: FieldEvaluationContext,
): string | undefined {
  const page = context.currentPageOf?.(paragraphPath)
  if (page === undefined) {
    return undefined
  }
  return (
    resolvePageField(rawInstruction, { pageNumber: page, totalPages: context.pageCount ?? page }) ??
    String(page)
  )
}

/** Structural (not text-value) info extracted from a HYPERLINK field's instruction, for a caller that wants the link target rather than a recalculated display string (see `evaluateFieldText`'s HYPERLINK case). */
export interface HyperlinkFieldInfo {
  readonly target: string
  /** `\l`: `target` is a bookmark name within this document, not an external URL/file. */
  readonly isLocalAnchor: boolean
  readonly tooltip?: string
  readonly targetFrame?: string
}

export function parseHyperlinkField(instruction: ParsedFieldInstruction): HyperlinkFieldInfo | undefined {
  const localAnchor = instruction.switches.get('l')
  const isLocalAnchor = localAnchor !== undefined
  const target = isLocalAnchor
    ? (typeof localAnchor === 'string' ? localAnchor : instruction.arguments[0])
    : instruction.arguments[0]

  if (target === undefined) {
    return undefined
  }

  const tooltip = instruction.switches.get('o')
  const targetFrame = instruction.switches.get('t')

  return {
    target,
    isLocalAnchor,
    ...(typeof tooltip === 'string' ? { tooltip } : {}),
    ...(typeof targetFrame === 'string' ? { targetFrame } : {}),
  }
}

export type { FieldType }
