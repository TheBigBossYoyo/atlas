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
  /**
   * The Windows ProgID for this extension (electron-builder's `name`, which it
   * writes verbatim as a key under `HKLM\SOFTWARE\Classes`).
   *
   * PROGID-1 — this field used to hold the human label ("Word Document"), which
   * is what its old name, `associationName`, invited. electron-builder does not
   * treat `name` as a label: it uses it as the ProgID, so an install created
   * about thirty UNQUALIFIED keys in a global hive — `Word Document`,
   * `Excel 97-2003 Workbook`, `Source Code` — each pointing at Atlas.exe. No
   * collision with Office, which uses `Word.Document.12`-style ids, but any
   * other app following the same electron-builder pattern would collide, and
   * Atlas's uninstaller would then delete keys it does not own. The visible
   * label was never this field's job; `associationDescription` is what Explorer
   * shows.
   *
   * Constraints enforced by `__tests__/extensionManifest.test.ts`: `Atlas.`
   * prefix, letters/digits/periods only, at most 39 characters, and one ProgID
   * per distinct description (so Explorer can still label a `.dotx` differently
   * from a `.docx`).
   */
  readonly associationProgId: string
  /** Windows file-association description — the label Explorer shows (electron-builder `description`). */
  readonly associationDescription: string
  readonly associationRole: WindowsAssociationRole
  /** Language id for `format: 'code'` entries — the code editor's language label and symbol-pattern key (USR-18). */
  readonly shikiLang?: string
}

type RawEntry = readonly [
  ext: string,
  format: FormatId,
  associationProgId: string,
  associationDescription: string,
  associationRole: WindowsAssociationRole,
  shikiLang?: string,
]

// ---------------------------------------------------------------------------
// Raw table — one row per extension. Grouped by format for readability; the
// derived exports below are what every consumer actually imports.
// ---------------------------------------------------------------------------

const MARKDOWN_ROWS: ReadonlyArray<RawEntry> = [
  ['md', 'markdown', 'Atlas.MarkdownDocument', 'Markdown Document', 'Editor'],
  ['markdown', 'markdown', 'Atlas.MarkdownDocument', 'Markdown Document', 'Editor'],
  ['mdown', 'markdown', 'Atlas.MarkdownDocument', 'Markdown Document', 'Editor'],
  ['mkd', 'markdown', 'Atlas.MarkdownDocument', 'Markdown Document', 'Editor'],
  ['mdx', 'markdown', 'Atlas.MDXDocument', 'Markdown Document with JSX', 'Editor'],
]

const DOCX_ROWS: ReadonlyArray<RawEntry> = [
  ['docx', 'docx', 'Atlas.WordDocument', 'Word Document', 'Editor'],
  ['docm', 'docx', 'Atlas.WordMacroDocument', 'Word Macro-Enabled Document', 'Editor'],
  ['dotx', 'docx', 'Atlas.WordTemplate', 'Word Template', 'Editor'],
  ['dotm', 'docx', 'Atlas.WordMacroTemplate', 'Word Macro-Enabled Template', 'Editor'],
]

const XLSX_ROWS: ReadonlyArray<RawEntry> = [
  ['xlsx', 'xlsx', 'Atlas.ExcelSpreadsheet', 'Excel Spreadsheet', 'Viewer'],
  ['xlsm', 'xlsx', 'Atlas.ExcelMacroSpreadsheet', 'Excel Macro-Enabled Spreadsheet', 'Viewer'],
  ['xlsb', 'xlsx', 'Atlas.ExcelBinarySpreadsheet', 'Excel Binary Spreadsheet', 'Viewer'],
  ['xltx', 'xlsx', 'Atlas.ExcelTemplate', 'Excel Template', 'Viewer'],
  ['xltm', 'xlsx', 'Atlas.ExcelMacroTemplate', 'Excel Macro-Enabled Template', 'Viewer'],
  // Wave 3 — legacy binary BIFF8 workbook. Unlike `.doc`/`.ppt`, SheetJS's
  // `XLSX.read`/`XLSX.write` genuinely parse and re-serialize this OLE2/CFB
  // -container format (`bookType: 'xls'`), so it is routed to the same
  // spreadsheet viewer rather than `legacyOffice.ts`'s honest-unsupported
  // message, with full in-place Save supported (not Save-As-only) — see
  // `spreadsheetWrite.ts`. One real, documented limitation: this build's
  // BIFF8 writer never serializes a cell's formula text, only its cached
  // value, so a formula saved back to `.xls` round-trips as a plain value.
  ['xls', 'xlsx', 'Atlas.Excel972003Workbook', 'Excel 97-2003 Workbook', 'Viewer'],
]

const PPTX_ROWS: ReadonlyArray<RawEntry> = [
  ['pptx', 'pptx', 'Atlas.PowerPointPresentation', 'PowerPoint Presentation', 'Viewer'],
  ['pptm', 'pptx', 'Atlas.PowerPointMacroPresentation', 'PowerPoint Macro-Enabled Presentation', 'Viewer'],
  ['potx', 'pptx', 'Atlas.PowerPointTemplate', 'PowerPoint Template', 'Viewer'],
  ['potm', 'pptx', 'Atlas.PowerPointMacroTemplate', 'PowerPoint Macro-Enabled Template', 'Viewer'],
  ['ppsx', 'pptx', 'Atlas.PowerPointSlideShow', 'PowerPoint Slide Show', 'Viewer'],
  ['ppsm', 'pptx', 'Atlas.PowerPointMacroSlideShow', 'PowerPoint Macro-Enabled Slide Show', 'Viewer'],
]

// wave-4 legacy-office — unlike `.xls` (genuine SheetJS BIFF8 read/write
// support, routed to the spreadsheet viewer above), `.doc`/`.ppt` get a
// hand-rolled, read-only, text-only best-effort reader (`src/legacy/`):
// OLE2/CFB container + FIB/piece-table (.doc) or PPT binary record tree
// (.ppt), with no writer at all. `associationRole: 'Viewer'` reflects that —
// no Save, Save As-only (through the existing modern DOCX/PPTX writers) if
// that ever gets wired up. Their `.dot`/`.pot`/`.pps` template/show siblings
// are deliberately left out of this wave's scope — they still fall through
// to `formats/legacyOffice.ts`'s honest "not supported" message via
// `UnknownViewer`, same as `.xlt` since wave 3.
const LEGACY_DOC_ROWS: ReadonlyArray<RawEntry> = [
  ['doc', 'doc', 'Atlas.Word972003Document', 'Word 97-2003 Document', 'Viewer'],
]

const LEGACY_PPT_ROWS: ReadonlyArray<RawEntry> = [
  ['ppt', 'ppt', 'Atlas.PowerPoint972003Presentation', 'PowerPoint 97-2003 Presentation', 'Viewer'],
]

const MISC_DOCUMENT_ROWS: ReadonlyArray<RawEntry> = [
  ['pdf', 'pdf', 'Atlas.PDFDocument', 'PDF Document', 'Viewer'],
  ['csv', 'csv', 'Atlas.CommaSeparatedValues', 'Comma-Separated Values File', 'Viewer'],
  ['tsv', 'tsv', 'Atlas.TabSeparatedValues', 'Tab-Separated Values File', 'Viewer'],
  ['tab', 'tsv', 'Atlas.TabSeparatedValues', 'Tab-Separated Values File', 'Viewer'],
  ['txt', 'text', 'Atlas.PlainText', 'Plain Text', 'Viewer'],
  ['log', 'text', 'Atlas.PlainTextLog', 'Plain Text Log', 'Viewer'],
  ['odt', 'odt', 'Atlas.OpenDocumentText', 'OpenDocument Text Document', 'Viewer'],
  ['ods', 'ods', 'Atlas.OpenDocumentSpreadsheet', 'OpenDocument Spreadsheet', 'Viewer'],
  ['odp', 'odp', 'Atlas.OpenDocumentPresentation', 'OpenDocument Presentation', 'Viewer'],
  ['rtf', 'rtf', 'Atlas.RichTextDocument', 'Rich Text Document', 'Viewer'],
  // Wave 3 — "Flat ODS": a single flat-XML file (no ZIP container) holding
  // the same OpenDocument spreadsheet schema as `.ods`. SheetJS reads and
  // writes it directly (`bookType: 'fods'`), so it routes to the same
  // spreadsheet viewer with full in-place Save supported.
  ['fods', 'ods', 'Atlas.FlatOpenDocumentSpreadsheet', 'Flat OpenDocument Spreadsheet', 'Viewer'],
]

// Source-code extensions -> language id (historically shiki ids; kept as the
// code editor's language label — see viewers/CodeViewer.tsx). Some are aliases
// rather than canonical names (e.g. `bash` for `.zsh`, `bat` for `.cmd`).
const CODE_ROWS: ReadonlyArray<RawEntry> = [
  ['ts', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'typescript'],
  ['tsx', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'tsx'],
  ['js', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'javascript'],
  ['jsx', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'jsx'],
  ['mjs', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'javascript'],
  ['cjs', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'javascript'],
  ['py', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'python'],
  ['pyw', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'python'],
  ['rb', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'ruby'],
  ['go', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'go'],
  ['rs', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'rust'],
  ['java', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'java'],
  ['kt', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'kotlin'],
  ['kts', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'kotlin'],
  ['swift', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'swift'],
  ['c', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'c'],
  ['h', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'c'],
  ['cpp', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'cpp'],
  ['hpp', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'cpp'],
  ['cc', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'cpp'],
  ['cs', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'csharp'],
  ['php', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'php'],
  ['pl', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'perl'],
  ['lua', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'lua'],
  ['r', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'r'],
  ['scala', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'scala'],
  ['clj', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'clojure'],
  ['ex', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'elixir'],
  ['exs', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'elixir'],
  ['erl', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'erlang'],
  ['hs', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'haskell'],
  ['ml', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'ocaml'],
  ['dart', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'dart'],
  ['vue', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'vue'],
  ['svelte', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'svelte'],
  ['sh', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'bash'],
  ['bash', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'bash'],
  ['zsh', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'bash'],
  ['fish', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'fish'],
  ['ps1', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'powershell'],
  ['bat', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'bat'],
  ['cmd', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'bat'],
  ['json', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'json'],
  ['jsonc', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'jsonc'],
  ['yaml', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'yaml'],
  ['yml', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'yaml'],
  ['toml', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'toml'],
  ['xml', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'xml'],
  ['html', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'html'],
  ['htm', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'html'],
  ['css', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'css'],
  ['scss', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'scss'],
  ['sass', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'sass'],
  ['less', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'less'],
  ['sql', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'sql'],
  ['graphql', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'graphql'],
  ['proto', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'proto'],
  ['dockerfile', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'docker'],
  ['ini', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'ini'],
  ['cfg', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'ini'],
  ['conf', 'code', 'Atlas.SourceCode', 'Source Code', 'Viewer', 'ini'],
]

const RAW_ENTRIES: ReadonlyArray<RawEntry> = [
  ...MARKDOWN_ROWS,
  ...DOCX_ROWS,
  ...XLSX_ROWS,
  ...PPTX_ROWS,
  ...LEGACY_DOC_ROWS,
  ...LEGACY_PPT_ROWS,
  ...MISC_DOCUMENT_ROWS,
  ...CODE_ROWS,
]

/** The full, ordered manifest — one entry per known extension. */
export const EXTENSION_MANIFEST: ReadonlyArray<ExtensionManifestEntry> = RAW_ENTRIES.map(
  ([ext, format, associationProgId, associationDescription, associationRole, shikiLang]) => ({
    ext,
    format,
    associationProgId,
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
