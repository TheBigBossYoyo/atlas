/**
 * Atlas — DOCX embedded-font loading (DEFER-4 / DXP-13)
 *
 * Orchestrates `./fontTable.ts` (parses `word/fontTable.xml`'s font/embed
 * entries) and `./deobfuscate.ts` (recovers real TTF/OTF bytes from an
 * obfuscated embedded font part) against a loaded DOCX archive's raw files,
 * producing ready-to-register font data per Word font name + variant.
 *
 * Deliberately a pure function over `DocxArchive['files']` (no DOM/FontFace
 * access here) so it stays unit-testable without a browser — the viewer
 * layer (`DocxViewer.tsx`) is responsible for actually registering the
 * returned bytes via `FontFace` and feeding them through `parseTtf`/
 * `buildMetrics` for canvas-measurement fallback metrics, exactly as it
 * already does for the bundled substitute fonts.
 */

import { parseRelationships } from '../parser/relationships'
import type { FontVariant } from './families'
import { xorObfuscatedFontHeader } from './deobfuscate'
import { parseFontTable, type EmbeddedFontRef, type FontTableEntry } from './fontTable'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** De-obfuscated font-file bytes for whichever of the four variants this document embeds. */
export type EmbeddedFontFaces = Readonly<Partial<Record<FontVariant, Uint8Array>>>

export interface EmbeddedFontFamily {
  /** The Word font name this document's runs reference (e.g. `w:rFonts`'s `w:ascii`). */
  readonly name: string
  readonly faces: EmbeddedFontFaces
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const FONT_TABLE_PART_PATH = 'word/fontTable.xml'
const FONT_TABLE_RELS_PATH = 'word/_rels/fontTable.xml.rels'

const VARIANT_BY_EMBED_KEY: ReadonlyArray<{
  readonly key: keyof Pick<FontTableEntry, 'embedRegular' | 'embedBold' | 'embedItalic' | 'embedBoldItalic'>
  readonly variant: FontVariant
}> = [
  { key: 'embedRegular', variant: 'regular' },
  { key: 'embedBold', variant: 'bold' },
  { key: 'embedItalic', variant: 'italic' },
  { key: 'embedBoldItalic', variant: 'boldItalic' },
]

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Reads and de-obfuscates every embedded font face this DOCX archive
 * carries. Returns an empty array when the document has no `fontTable.xml`,
 * no embedded fonts within it, or no resolvable relationships part — all
 * treated as "this document embeds no fonts" rather than an error, since
 * embedding is an opt-in Word feature most documents don't use.
 *
 * A face whose relationship target or font-part bytes can't be resolved is
 * skipped individually (its family may still surface other, resolvable
 * faces) rather than failing the whole document load.
 *
 * @param files - The raw archive files map from `unzipDocx` (`DocxArchive.files`
 *   / `DocxBundle.rawArchive`), keyed by full zip entry path.
 */
export function loadEmbeddedFonts(
  files: ReadonlyMap<string, Uint8Array> | undefined,
): ReadonlyArray<EmbeddedFontFamily> {
  if (files === undefined) {
    return []
  }

  const fontTableXml = decodeUtf8(files.get(FONT_TABLE_PART_PATH));
  if (fontTableXml === undefined) {
    return []
  }

  const entries = parseFontTable(fontTableXml)
  if (entries.length === 0) {
    return []
  }

  const relsXml = decodeUtf8(files.get(FONT_TABLE_RELS_PATH))
  const relationships = relsXml !== undefined ? parseRelationships(relsXml) : []
  const targetByRelId = new Map(relationships.map((rel) => [rel.id, rel.target]))

  const families: EmbeddedFontFamily[] = []
  for (const entry of entries) {
    const faces = resolveFaces(entry, targetByRelId, files)
    if (Object.keys(faces).length > 0) {
      families.push({ name: entry.name, faces })
    }
  }

  return families
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function resolveFaces(
  entry: FontTableEntry,
  targetByRelId: ReadonlyMap<string, string>,
  files: ReadonlyMap<string, Uint8Array>,
): EmbeddedFontFaces {
  const faces: Partial<Record<FontVariant, Uint8Array>> = {}

  for (const { key, variant } of VARIANT_BY_EMBED_KEY) {
    const ref = entry[key]
    if (ref === undefined) {
      continue
    }

    const face = resolveFace(ref, targetByRelId, files)
    if (face !== undefined) {
      faces[variant] = face
    }
  }

  return faces
}

function resolveFace(
  ref: EmbeddedFontRef,
  targetByRelId: ReadonlyMap<string, string>,
  files: ReadonlyMap<string, Uint8Array>,
): Uint8Array | undefined {
  const target = targetByRelId.get(ref.relId)
  if (target === undefined) {
    return undefined
  }

  const partPath = resolvePartPath(FONT_TABLE_PART_PATH, target)
  const rawBytes = files.get(partPath)
  if (rawBytes === undefined) {
    return undefined
  }

  if (ref.fontKey === undefined) {
    // No obfuscation key recorded: Word always writes one alongside an
    // embed element it obfuscated, so this is an unusual/malformed part —
    // use the bytes as-is rather than guessing at a key.
    return rawBytes
  }

  try {
    return xorObfuscatedFontHeader(rawBytes, ref.fontKey)
  } catch {
    // Malformed w:fontKey: skip this face rather than fail the whole load.
    return undefined
  }
}

/**
 * Resolves a relationship `Target` against the part that declared it, the
 * same way `docx/index.ts` resolves header/footer relationship targets:
 * OPC-relative targets are resolved against the declaring part's own
 * directory (here `word/`, since `fontTable.xml` lives directly under
 * `word/`), with `../` segments walked normally.
 */
function resolvePartPath(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1)
  }

  const sourceDir = sourcePartPath.split('/').slice(0, -1)
  const segments = [...sourceDir, ...target.split('/')]

  const resolved: string[] = []
  for (const segment of segments) {
    if (segment === '' || segment === '.') {
      continue
    }
    if (segment === '..') {
      resolved.pop()
      continue
    }
    resolved.push(segment)
  }

  return resolved.join('/')
}

function decodeUtf8(bytes: Uint8Array | undefined): string | undefined {
  if (bytes === undefined) {
    return undefined
  }
  return new TextDecoder().decode(bytes)
}
