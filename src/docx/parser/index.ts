/**
 * Atlas — DOCX parser barrel (Wave A.1)
 *
 * Re-exports all public symbols from the A.1 foundation modules so
 * downstream code can import from a single entry point:
 *
 *   import { unzipDocx, parseRelationships, parseContentTypes } from '@/docx/parser'
 */

export { unzipDocx, DocxParseError } from './unzip'
export type { DocxArchive } from './unzip'

export { parseRelationships } from './relationships'
export type { Relationship } from './relationships'

export { parseContentTypes } from './contentTypes'
export type {
  ContentTypes,
  ContentTypeDefault,
  ContentTypeOverride,
} from './contentTypes'

export { parseDocument } from './document'

export { parseStyles } from './styles'
export type { StylesPart } from './styles'

export { parseNumbering } from './numbering'
export type { AbstractNum, NumberingPart, NumInstance } from './numbering'

export { resolveParaProps, resolveRunProps } from './cascade'

// Wave A.5 — part parsers

export { parseHeader } from './headers'

export { parseFooter } from './footers'

export { parseFootnotes } from './footnotes'

export { parseEndnotes } from './endnotes'

export { parseComments } from './comments'

export { parseCommentsExtended } from './commentsExtended'

export { parseTheme, resolveThemeFont } from './theme'
export type { Theme, FontScheme, FontSet } from './theme'
