import { t } from '../i18n'

/**
 * P2.11 / LOAD-11 — legacy (pre-XML) Microsoft Office binary format
 * detection. `.doc`/`.ppt` (and their `.dot`/`.pot`/`.pps` template/show
 * siblings) use the OLE2 Compound File Binary format, which Atlas cannot
 * parse — it is a completely different container from the ZIP-based OOXML
 * formats (`.docx`/`.xlsx`/`.pptx`) the app does support.
 *
 * Rather than falling through to a bare "unknown format" empty state, a file
 * with the CFB magic header gets a specific, actionable message pointing at
 * the fix (re-save as the modern XML format). `UnknownViewer` is the only
 * consumer.
 *
 * Wave 3 exception: `.xls`/`.xlt` are ALSO CFB-container files, but SheetJS's
 * `XLSX.read` has genuine native BIFF8 support and parses them correctly —
 * so `extensionManifest.ts` now routes `.xls` straight to the spreadsheet
 * viewer, which never reaches this module at all (only a genuinely-unknown
 * extension gets this far). `.xlt` is left mapped below and still shown this
 * message, since it was never in this wave's explicit scope.
 *
 * Wave 4 exception: `.doc`/`.ppt` are likewise routed straight past this
 * module now — `extensionManifest.ts` sends them to `LegacyDocViewer`/
 * `LegacyPptViewer`, a hand-rolled read-only, text-only reader (`src/legacy/`)
 * for these two specifically. Their `.dot`/`.pot`/`.pps` template/show
 * siblings were not in that wave's scope either, so they (like `.xlt`) still
 * reach `UnknownViewer` and this module's honest "not supported" message.
 */

/** The 8-byte OLE2/CFB signature shared by every legacy Office binary format. */
const CFB_MAGIC: ReadonlyArray<number> = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

export function isLegacyOfficeMagic(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength < CFB_MAGIC.length) {
    return false
  }

  const bytes = new Uint8Array(buffer, 0, CFB_MAGIC.length)
  return CFB_MAGIC.every((expected, index) => bytes[index] === expected)
}

export type LegacyOfficeKind = 'doc' | 'xls' | 'ppt' | 'unknown'

const LEGACY_EXTENSION_KIND: Readonly<Record<string, LegacyOfficeKind>> = {
  doc: 'doc',
  dot: 'doc',
  xls: 'xls',
  xlt: 'xls',
  ppt: 'ppt',
  pot: 'ppt',
  pps: 'ppt',
}

/**
 * The CFB container itself doesn't cheaply reveal which application wrote
 * it without walking its internal directory sectors — but these files
 * almost always still carry their original `.doc`/`.xls`/`.ppt`-family
 * extension, which is a reliable-enough signal for a friendly message.
 */
export function guessLegacyOfficeKind(path: string): LegacyOfficeKind {
  const fileName = path.split(/[\\/]/).pop() ?? path
  const dotIndex = fileName.lastIndexOf('.')

  if (dotIndex <= 0 || dotIndex === fileName.length - 1) {
    return 'unknown'
  }

  const extension = fileName.slice(dotIndex + 1).toLowerCase()
  return LEGACY_EXTENSION_KIND[extension] ?? 'unknown'
}

/** The `legacyOffice.label.*` i18n key for each kind — resolved at call time
 * (not module load) so the message stays in the active locale even if it
 * changes after this module is first imported. Also reused verbatim by
 * `LegacyDocViewer.tsx`/`LegacyPptViewer.tsx` for their own format-label
 * prop, so the "doc"/"ppt" wording never drifts between the two call sites. */
export function legacyOfficeLabelKey(kind: LegacyOfficeKind): string {
  return `legacyOffice.label.${kind}`
}

function legacyOfficeModernExtension(kind: LegacyOfficeKind): string {
  switch (kind) {
    case 'doc':
      return '.docx'
    case 'xls':
      return '.xlsx'
    case 'ppt':
      return '.pptx'
    case 'unknown':
      return t('legacyOffice.modernExtensionUnknown')
  }
}

export function legacyOfficeMessage(kind: LegacyOfficeKind): string {
  return t('legacyOffice.message', {
    label: t(legacyOfficeLabelKey(kind)),
    modernExtension: legacyOfficeModernExtension(kind),
  })
}
