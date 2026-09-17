/**
 * Atlas — Format type system (W1.1)
 *
 * Single source of truth for FormatId, LoadedFile, ViewerStats, and NavItem.
 * Every viewer, detector, registry entry, sidebar, and status bar must reference
 * these types so exhaustiveness is enforced by the compiler via `assertNever`.
 */

import type { ComponentType } from 'react'
import type { LucideIcon } from 'lucide-react'

/**
 * The 14 supported format identifiers.
 * `unknown` is the fallback when extension and magic-byte detection both fail.
 */
export type FormatId =
  | 'markdown'
  | 'docx'
  | 'xlsx'
  | 'pptx'
  | 'pdf'
  | 'csv'
  | 'tsv'
  | 'text'
  | 'code'
  | 'odt'
  | 'ods'
  | 'odp'
  | 'rtf'
  // wave-4 legacy-office — legacy binary (pre-XML) Word/PowerPoint formats.
  // Distinct from 'docx'/'pptx': read-only, text-only best-effort extraction
  // via `src/legacy/` (no OLE2/CFB parsing in common with the ZIP-based
  // OOXML formats) — see `viewers/LegacyDocViewer.tsx`/`LegacyPptViewer.tsx`.
  | 'doc'
  | 'ppt'
  | 'unknown'

/** All non-`unknown` FormatIds, useful for iteration / registry exhaustiveness. */
export const KNOWN_FORMAT_IDS = [
  'markdown',
  'docx',
  'xlsx',
  'pptx',
  'pdf',
  'csv',
  'tsv',
  'text',
  'code',
  'odt',
  'ods',
  'odp',
  'rtf',
  'doc',
  'ppt',
] as const satisfies ReadonlyArray<Exclude<FormatId, 'unknown'>>

/** All FormatIds including `unknown`. Order matches `FormatId` declaration. */
export const ALL_FORMAT_IDS = [...KNOWN_FORMAT_IDS, 'unknown'] as const satisfies ReadonlyArray<FormatId>

/**
 * A loaded file ready for a viewer. Discriminated on `kind`:
 *   - `text`   → UTF-8 string content (markdown, csv/tsv, text, code)
 *   - `binary` → ArrayBuffer content   (pdf, docx, xlsx, pptx, odt/ods/odp, rtf)
 *
 * `format` carries the resolved FormatId (post extension + magic-byte detection).
 * `path` is the absolute filesystem path; renderer code must never mutate it.
 */
export type LoadedFile =
  | {
      readonly kind: 'text'
      readonly content: string
      readonly path: string
      readonly format: FormatId
    }
  | {
      readonly kind: 'binary'
      readonly content: ArrayBuffer
      readonly path: string
      readonly format: FormatId
    }

/**
 * Per-format statistics surfaced in the StatusBar. Discriminated on `kind`
 * (NOT FormatId — multiple formats share the same stats shape, e.g. xlsx + ods
 * both report `kind: 'spreadsheet'`).
 */
export type ViewerStats =
  | { readonly kind: 'markdown'; readonly words: number; readonly headings: number }
  | { readonly kind: 'spreadsheet'; readonly sheet: string; readonly rows: number; readonly cols: number }
  | { readonly kind: 'pdf'; readonly page: number; readonly pageCount: number }
  | { readonly kind: 'slides'; readonly slide: number; readonly slideCount: number }
  | { readonly kind: 'code'; readonly language: string; readonly lines: number }
  | { readonly kind: 'text'; readonly lines: number; readonly chars: number }
  | { readonly kind: 'document'; readonly words: number; readonly pages: number }

/**
 * A single navigation item for the universal Sidebar (TOC for markdown, sheets
 * for spreadsheets, slide thumbnails for slides, page list for PDF, etc.).
 */
export type NavItem = {
  readonly id: string
  readonly label: string
  readonly icon?: LucideIcon
  readonly level?: number
  onSelect: () => void
}

/**
 * Standard viewer component contract. Every entry in the viewerRegistry resolves
 * to a `ComponentType<ViewerProps>`. Viewers may optionally report stats and
 * nav items via callbacks (W3.1 / W3.2 will wire these into Sidebar/StatusBar).
 */
export type ViewerProps = {
  readonly file: LoadedFile
  readonly onStats?: (stats: ViewerStats) => void
  readonly onNavItems?: (items: ReadonlyArray<NavItem>) => void
}

/** Convenience alias for the registry value type. */
export type ViewerComponent = ComponentType<ViewerProps>

/**
 * Exhaustiveness helper. Use in the `default` arm of a `switch (x.format)` or
 * `switch (x.kind)` to make TypeScript flag any unhandled FormatId / kind at
 * compile time and throw at runtime if the type system is bypassed.
 *
 * @example
 *   function pick(f: FormatId) {
 *     switch (f) {
 *       case 'markdown': return MD
 *       // ...all other cases...
 *       default: return assertNever(f)
 *     }
 *   }
 */
export function assertNever(x: never): never {
  throw new Error(`assertNever: unexpected value ${JSON.stringify(x)}`)
}
