/**
 * Atlas — canonical extension manifest (P2.2 / ELEC-05, ELEC-15, LOAD-03, LOAD-12, LOAD-22)
 *
 * Before this file existed, "which extensions does Atlas know about" was
 * hand-duplicated across four disagreeing sources: `detect.ts`'s
 * `EXTENSION_MAP`, `electron/main.cjs`'s `KNOWN_EXTENSIONS` set + its Open
 * dialog filter, and `electron-builder.yml`'s `fileAssociations` list. Each
 * only knew a different subset, so a Windows file association could exist
 * (`.mdown`, `.ini`) that the app itself couldn't recognize on open
 * (LOAD-03), and vice versa.
 *
 * This module is now the single source of truth:
 *   - `detect.ts` derives its extension→format map from `EXTENSION_TO_FORMAT`.
 *   - `viewers/extToLang.ts` derives its shiki-language map from
 *     `CODE_EXTENSION_TO_SHIKI_LANG` (T6 / DAT-14).
 *   - `scripts/generate-extension-manifest.mjs` reads `EXTENSION_MANIFEST`
 *     directly (Node 24's native TS type-stripping — see that script's
 *     header) and regenerates `electron/lib/extensionManifest.generated.cjs`
 *     (consumed by `electron/main.cjs` for its known-extensions set and Open
 *     dialog filter) and the `fileAssociations` block of
 *     `electron-builder.yml`. Run `npm run generate:extensions` after
 *     editing this file — `npm run build` also does this automatically via
 *     the `prebuild` script, and
 *     `src/formats/__tests__/extensionManifest.test.ts` fails CI if either
 *     generated artifact drifts from what's below.
 *
 * `associationRole` implements ELEC-20: only formats Atlas can actually save
 * (markdown and DOCX) are registered as `Editor`; every other format —
 * spreadsheets, slides, PDF, plain text/code, RTF, ODF — is view-only today
 * and registered as `Viewer`.
 */
import type { FormatId } from './types'

export type WindowsAssociationRole = 'Editor' | 'Viewer'

export interface ExtensionManifestEntry {
  /** Lowercase extension with no leading dot, e.g. `"docx"`. */
  readonly ext: string
  readonly format: FormatId
  /** Windows file-association display name (electron-builder `name`). */
  readonly associationName: string
  /** Windows file-association description (electron-builder `description`). */
  readonly associationDescription: string
  readonly associationRole: WindowsAssociationRole
  /** Language id for `format: 'code'` entries — the code editor's language label and symbol-pattern key (USR-18). */
  readonly shikiLang?: string
}

type RawEntry = readonly [
  ext: string,
  format: FormatId,
  associationName: string,
  associationDescription: string,
  associationRole: WindowsAssociationRole,
  shikiLang?: string,
]

// ---------------------------------------------------------------------------
// Raw table — one row per extension. Grouped by format for readability; the
// derived exports below are what every consumer actually imports.
// ---------------------------------------------------------------------------

const MARKDOWN_ROWS: ReadonlyArray<RawEntry> = [
  ['md', 'markdown', 'Markdown Document', 'Markdown Document', 'Editor'],
  ['markdown', 'markdown', 'Markdown Document', 'Markdown Document', 'Editor'],
  ['mdown', 'markdown', 'Markdown Document', 'Markdown Document', 'Editor'],
  ['mkd', 'markdown', 'Markdown Document', 'Markdown Document', 'Editor'],
  ['mdx', 'markdown', 'MDX Document', 'Markdown Document with JSX', 'Editor'],
]

const DOCX_ROWS: ReadonlyArray<RawEntry> = [
  ['docx', 'docx', 'Word Document', 'Word Document', 'Editor'],
  ['docm', 'docx', 'Word Macro-Enabled Document', 'Word Macro-Enabled Document', 'Editor'],
  ['dotx', 'docx', 'Word Template', 'Word Template', 'Editor'],
  ['dotm', 'docx', 'Word Macro-Enabled Template', 'Word Macro-Enabled Template', 'Editor'],
]

const XLSX_ROWS: ReadonlyArray<RawEntry> = [
  ['xlsx', 'xlsx', 'Excel Spreadsheet', 'Excel Spreadsheet', 'Viewer'],
  ['xlsm', 'xlsx', 'Excel Macro-Enabled Spreadsheet', 'Excel Macro-Enabled Spreadsheet', 'Viewer'],
  ['xlsb', 'xlsx', 'Excel Binary Spreadsheet', 'Excel Binary Spreadsheet', 'Viewer'],
  ['xltx', 'xlsx', 'Excel Template', 'Excel Template', 'Viewer'],
  ['xltm', 'xlsx', 'Excel Macro-Enabled Template', 'Excel Macro-Enabled Template', 'Viewer'],
  // Wave 3 — legacy binary BIFF8 workbook. Unlike `.doc`/`.ppt`, SheetJS's
  // `XLSX.read`/`XLSX.write` genuinely parse and re-serialize this OLE2/CFB
  // -container format (`bookType: 'xls'`), so it is routed to the same
  // spreadsheet viewer rather than `legacyOffice.ts`'s honest-unsupported
  // message, with full in-place Save supported (not Save-As-only) — see
  // `spreadsheetWrite.ts`. One real, documented limitation: this build's
  // BIFF8 writer never serializes a cell's formula text, only its cached
  // value, so a formula saved back to `.xls` round-trips as a plain value.
  ['xls', 'xlsx', 'Excel 97-2003 Workbook', 'Excel 97-2003 Workbook', 'Viewer'],
]

const PPTX_ROWS: ReadonlyArray<RawEntry> = [
  ['pptx', 'pptx', 'PowerPoint Presentation', 'PowerPoint Presentation', 'Viewer'],
  ['pptm', 'pptx', 'PowerPoint Macro-Enabled Presentation', 'PowerPoint Macro-Enabled Presentation', 'Viewer'],
  ['potx', 'pptx', 'PowerPoint Template', 'PowerPoint Template', 'Viewer'],
  ['potm', 'pptx', 'PowerPoint Macro-Enabled Template', 'PowerPoint Macro-Enabled Template', 'Viewer'],
  ['ppsx', 'pptx', 'PowerPoint Slide Show', 'PowerPoint Slide Show', 'Viewer'],
  ['ppsm', 'pptx', 'PowerPoint Macro-Enabled Slide Show', 'PowerPoint Macro-Enabled Slide Show', 'Viewer'],
]

const MISC_DOCUMENT_ROWS: ReadonlyArray<RawEntry> = [
  ['pdf', 'pdf', 'PDF Document', 'PDF Document', 'Viewer'],
  ['csv', 'csv', 'Comma-Separated Values', 'Comma-Separated Values File', 'Viewer'],
  ['tsv', 'tsv', 'Tab-Separated Values', 'Tab-Separated Values File', 'Viewer'],
  ['tab', 'tsv', 'Tab-Separated Values', 'Tab-Separated Values File', 'Viewer'],
  ['txt', 'text', 'Plain Text', 'Plain Text', 'Viewer'],
  ['log', 'text', 'Plain Text', 'Plain Text Log', 'Viewer'],
  ['odt', 'odt', 'OpenDocument Text', 'OpenDocument Text Document', 'Viewer'],
  ['ods', 'ods', 'OpenDocument Spreadsheet', 'OpenDocument Spreadsheet', 'Viewer'],
  ['odp', 'odp', 'OpenDocument Presentation', 'OpenDocument Presentation', 'Viewer'],
  ['rtf', 'rtf', 'Rich Text Document', 'Rich Text Document', 'Viewer'],
  // Wave 3 — "Flat ODS": a single flat-XML file (no ZIP container) holding
  // the same OpenDocument spreadsheet schema as `.ods`. SheetJS reads and
  // writes it directly (`bookType: 'fods'`), so it routes to the same
  // spreadsheet viewer with full in-place Save supported.
  ['fods', 'ods', 'Flat OpenDocument Spreadsheet', 'Flat OpenDocument Spreadsheet', 'Viewer'],
]

// Source-code extensions -> language id (historically shiki ids; kept as the
// code editor's language label — see viewers/CodeViewer.tsx). Some are aliases
// rather than canonical names (e.g. `bash` for `.zsh`, `bat` for `.cmd`).
const CODE_ROWS: ReadonlyArray<RawEntry> = [
  ['ts', 'code', 'Source Code', 'Source Code', 'Viewer', 'typescript'],
  ['tsx', 'code', 'Source Code', 'Source Code', 'Viewer', 'tsx'],
  ['js', 'code', 'Source Code', 'Source Code', 'Viewer', 'javascript'],
  ['jsx', 'code', 'Source Code', 'Source Code', 'Viewer', 'jsx'],
  ['mjs', 'code', 'Source Code', 'Source Code', 'Viewer', 'javascript'],
  ['cjs', 'code', 'Source Code', 'Source Code', 'Viewer', 'javascript'],
  ['py', 'code', 'Source Code', 'Source Code', 'Viewer', 'python'],
  ['pyw', 'code', 'Source Code', 'Source Code', 'Viewer', 'python'],
  ['rb', 'code', 'Source Code', 'Source Code', 'Viewer', 'ruby'],
  ['go', 'code', 'Source Code', 'Source Code', 'Viewer', 'go'],
  ['rs', 'code', 'Source Code', 'Source Code', 'Viewer', 'rust'],
  ['java', 'code', 'Source Code', 'Source Code', 'Viewer', 'java'],
  ['kt', 'code', 'Source Code', 'Source Code', 'Viewer', 'kotlin'],
  ['kts', 'code', 'Source Code', 'Source Code', 'Viewer', 'kotlin'],
  ['swift', 'code', 'Source Code', 'Source Code', 'Viewer', 'swift'],
  ['c', 'code', 'Source Code', 'Source Code', 'Viewer', 'c'],
  ['h', 'code', 'Source Code', 'Source Code', 'Viewer', 'c'],
  ['cpp', 'code', 'Source Code', 'Source Code', 'Viewer', 'cpp'],
  ['hpp', 'code', 'Source Code', 'Source Code', 'Viewer', 'cpp'],
  ['cc', 'code', 'Source Code', 'Source Code', 'Viewer', 'cpp'],
  ['cs', 'code', 'Source Code', 'Source Code', 'Viewer', 'csharp'],
  ['php', 'code', 'Source Code', 'Source Code', 'Viewer', 'php'],
  ['pl', 'code', 'Source Code', 'Source Code', 'Viewer', 'perl'],
  ['lua', 'code', 'Source Code', 'Source Code', 'Viewer', 'lua'],
  ['r', 'code', 'Source Code', 'Source Code', 'Viewer', 'r'],
  ['scala', 'code', 'Source Code', 'Source Code', 'Viewer', 'scala'],
  ['clj', 'code', 'Source Code', 'Source Code', 'Viewer', 'clojure'],
  ['ex', 'code', 'Source Code', 'Source Code', 'Viewer', 'elixir'],
  ['exs', 'code', 'Source Code', 'Source Code', 'Viewer', 'elixir'],
  ['erl', 'code', 'Source Code', 'Source Code', 'Viewer', 'erlang'],
  ['hs', 'code', 'Source Code', 'Source Code', 'Viewer', 'haskell'],
  ['ml', 'code', 'Source Code', 'Source Code', 'Viewer', 'ocaml'],
  ['dart', 'code', 'Source Code', 'Source Code', 'Viewer', 'dart'],
  ['vue', 'code', 'Source Code', 'Source Code', 'Viewer', 'vue'],
  ['svelte', 'code', 'Source Code', 'Source Code', 'Viewer', 'svelte'],
  ['sh', 'code', 'Source Code', 'Source Code', 'Viewer', 'bash'],
  ['bash', 'code', 'Source Code', 'Source Code', 'Viewer', 'bash'],
  ['zsh', 'code', 'Source Code', 'Source Code', 'Viewer', 'bash'],
  ['fish', 'code', 'Source Code', 'Source Code', 'Viewer', 'fish'],
  ['ps1', 'code', 'Source Code', 'Source Code', 'Viewer', 'powershell'],
  ['bat', 'code', 'Source Code', 'Source Code', 'Viewer', 'bat'],
  ['cmd', 'code', 'Source Code', 'Source Code', 'Viewer', 'bat'],
  ['json', 'code', 'Source Code', 'Source Code', 'Viewer', 'json'],
  ['jsonc', 'code', 'Source Code', 'Source Code', 'Viewer', 'jsonc'],
  ['yaml', 'code', 'Source Code', 'Source Code', 'Viewer', 'yaml'],
  ['yml', 'code', 'Source Code', 'Source Code', 'Viewer', 'yaml'],
  ['toml', 'code', 'Source Code', 'Source Code', 'Viewer', 'toml'],
  ['xml', 'code', 'Source Code', 'Source Code', 'Viewer', 'xml'],
  ['html', 'code', 'Source Code', 'Source Code', 'Viewer', 'html'],
  ['htm', 'code', 'Source Code', 'Source Code', 'Viewer', 'html'],
  ['css', 'code', 'Source Code', 'Source Code', 'Viewer', 'css'],
  ['scss', 'code', 'Source Code', 'Source Code', 'Viewer', 'scss'],
  ['sass', 'code', 'Source Code', 'Source Code', 'Viewer', 'sass'],
  ['less', 'code', 'Source Code', 'Source Code', 'Viewer', 'less'],
  ['sql', 'code', 'Source Code', 'Source Code', 'Viewer', 'sql'],
  ['graphql', 'code', 'Source Code', 'Source Code', 'Viewer', 'graphql'],
  ['proto', 'code', 'Source Code', 'Source Code', 'Viewer', 'proto'],
  ['dockerfile', 'code', 'Source Code', 'Source Code', 'Viewer', 'docker'],
  ['ini', 'code', 'Source Code', 'Source Code', 'Viewer', 'ini'],
  ['cfg', 'code', 'Source Code', 'Source Code', 'Viewer', 'ini'],
  ['conf', 'code', 'Source Code', 'Source Code', 'Viewer', 'ini'],
]

const RAW_ENTRIES: ReadonlyArray<RawEntry> = [
  ...MARKDOWN_ROWS,
  ...DOCX_ROWS,
  ...XLSX_ROWS,
  ...PPTX_ROWS,
  ...MISC_DOCUMENT_ROWS,
  ...CODE_ROWS,
]

/** The full, ordered manifest — one entry per known extension. */
export const EXTENSION_MANIFEST: ReadonlyArray<ExtensionManifestEntry> = RAW_ENTRIES.map(
  ([ext, format, associationName, associationDescription, associationRole, shikiLang]) => ({
    ext,
    format,
    associationName,
    associationDescription,
    associationRole,
    ...(shikiLang !== undefined ? { shikiLang } : {}),
  }),
)

// ---------------------------------------------------------------------------
// Derived exports — what consumers actually import.
// ---------------------------------------------------------------------------

/** `detect.ts`'s extension→format table. */
export const EXTENSION_TO_FORMAT: Readonly<Record<string, FormatId>> = Object.freeze(
  Object.fromEntries(EXTENSION_MANIFEST.map((entry) => [entry.ext, entry.format])),
)

/** `viewers/extToLang.ts`'s extension→shiki-language table (T6 / DAT-14). */
export const CODE_EXTENSION_TO_SHIKI_LANG: Readonly<Record<string, string>> = Object.freeze(
  Object.fromEntries(
    EXTENSION_MANIFEST.filter((entry): entry is ExtensionManifestEntry & { shikiLang: string } =>
      entry.shikiLang !== undefined,
    ).map((entry) => [entry.ext, entry.shikiLang]),
  ),
)

/** Every known extension, deduplicated, sorted — feeds `main.cjs`'s known-extensions set and Open dialog filter. */
export const ALL_MANIFEST_EXTENSIONS: ReadonlyArray<string> = Object.freeze(
  [...new Set(EXTENSION_MANIFEST.map((entry) => entry.ext))].sort(),
)
