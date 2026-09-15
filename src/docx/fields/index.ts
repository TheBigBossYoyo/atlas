/**
 * Atlas — DOCX field evaluation, recalculation, and TOC generation
 * (DEFER-5 / DXS-20)
 *
 * See `types.ts` (the context an evaluation needs), `instruction.ts`
 * (parsing `Field.instruction`), `dateFormat.ts` (the `\@` picture
 * formatter), `evaluate.ts` (per-field-type evaluators), `updateFields.ts`
 * (the "Update field(s)" command), `toc.ts` ("Update table of contents" —
 * a TOC field's cached result is a whole entry list, not a single value,
 * so it has its own generator + command), and `bookmarks.ts` (building the
 * `bookmarkText`/`bookmarkPage` context maps REF/PAGEREF evaluation needs
 * from the actual document).
 */
export type { FieldEvaluationContext } from './types'
export { collectBookmarkMaps, type BookmarkMaps } from './bookmarks'
export { parseFieldInstruction, type ParsedFieldInstruction } from './instruction'
export { DEFAULT_DATE_PICTURE, DEFAULT_TIME_PICTURE, formatDatePicture } from './dateFormat'
export { evaluateFieldText, parseHyperlinkField, type HyperlinkFieldInfo } from './evaluate'
export { updateFields, type UpdateFieldsResult } from './updateFields'
export {
  collectTocEntries,
  parseTocOptions,
  renderTocFieldResult,
  updateTableOfContents,
  type TocEntry,
  type TocOptions,
  type TocPageResolver,
  type UpdateTableOfContentsResult,
} from './toc'
export { paragraphPlainText } from './paragraphText'
export { parseCoreProps, type CoreProps } from './docProps'
