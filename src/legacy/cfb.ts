/**
 * Compound File Binary (OLE2/CFB) container reader (wave-4 legacy-office).
 *
 * `.doc`/`.ppt` (and `.xls`) all share the same OLE2 "structured storage"
 * container format underneath their format-specific streams. Rather than
 * re-implementing OLE2's FAT/MiniFAT sector-chain walking ourselves, this
 * wraps the SheetJS `xlsx` package's bundled `CFB` codec — the exact same
 * reader `XLSX.read` already relies on internally for genuine `.xls`
 * support (see `formats/legacyOffice.ts`'s header comment), so it's already
 * exercised against real-world CFB files at scale.
 *
 * `XLSX.CFB`'s public type is `any` (see `node_modules/xlsx/types/index.d.ts`
 * — it's commented out entirely upstream), and every file this module reads
 * is untrusted input. So nothing from `XLSX.CFB.read`'s result is trusted
 * structurally either: every field is validated here before being handed
 * back as a typed `CfbContainer`, and a container/entry shaped differently
 * than expected degrades to a safe default (empty content, 'unknown' type)
 * rather than throwing deep inside an unrelated caller.
 */
import * as XLSX from 'xlsx'

import { LegacyFormatError } from './errors'

export type CfbEntryType = 'unknown' | 'storage' | 'stream' | 'lockbytes' | 'property' | 'root'

export interface CfbEntry {
  /** Full path within the container, e.g. `"Root Entry/WordDocument"`. */
  readonly path: string
  /** Final path segment, e.g. `"WordDocument"`. */
  readonly name: string
  readonly type: CfbEntryType
  /** `null` for storages/root; a stream's raw bytes otherwise. */
  readonly content: Uint8Array | null
}

export interface CfbContainer {
  readonly entries: ReadonlyArray<CfbEntry>
}

// [MS-CFB] 2.6.1 — index into this array is the on-disk `objectType` byte.
const ENTRY_TYPES: ReadonlyArray<CfbEntryType> = ['unknown', 'storage', 'stream', 'lockbytes', 'property', 'root']

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function toEntryType(raw: unknown): CfbEntryType {
  return typeof raw === 'number' && raw >= 0 && raw < ENTRY_TYPES.length ? ENTRY_TYPES[raw] : 'unknown'
}

function toContent(raw: unknown): Uint8Array | null {
  if (raw == null) {
    return null
  }

  // `raw instanceof Uint8Array` is realm-sensitive: under Vitest's jsdom
  // environment, the bundled CFB codec's Node `Buffer` result can carry a
  // *different* `Uint8Array` prototype than the one this module's own
  // `instanceof` check closes over, making the check silently fail (Buffer
  // content coming back as `null` for every entry) despite `raw` genuinely
  // being a byte buffer at runtime. `ArrayBuffer.isView` uses an internal
  // slot check instead of prototype identity, so it correctly recognizes a
  // typed array/Buffer from *any* realm — the same reason it's the standard
  // cross-realm-safe alternative to `instanceof TypedArray`.
  if (ArrayBuffer.isView(raw)) {
    return raw instanceof Uint8Array ? raw : new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength)
  }

  return null
}

/**
 * Parses raw bytes as an OLE2/CFB container. Throws `LegacyFormatError` for
 * anything that isn't a well-formed CFB document — including a structurally
 * plausible but empty/mismatched directory, which the underlying codec
 * itself would happily hand back rather than reject.
 */
export function readCfb(bytes: Uint8Array): CfbContainer {
  let raw: unknown
  try {
    raw = XLSX.CFB.read(bytes, { type: 'array' })
  } catch (err) {
    throw new LegacyFormatError(
      `This file isn't a valid Compound File Binary (OLE2) document: ${err instanceof Error ? err.message : String(err)}`,
    )
  }

  if (!isRecord(raw) || !Array.isArray(raw.FullPaths) || !Array.isArray(raw.FileIndex)) {
    throw new LegacyFormatError("This file isn't a valid Compound File Binary (OLE2) document.")
  }

  const { FullPaths, FileIndex } = raw
  if (FullPaths.length !== FileIndex.length || FullPaths.length === 0) {
    throw new LegacyFormatError('Corrupt Compound File Binary document: directory entry count mismatch.')
  }

  const entries: CfbEntry[] = FullPaths.map((rawPath: unknown, index: number) => {
    const rawEntry: unknown = FileIndex[index]
    const path = typeof rawPath === 'string' ? rawPath : ''

    return {
      path,
      name: isRecord(rawEntry) && typeof rawEntry.name === 'string' ? rawEntry.name : '',
      type: isRecord(rawEntry) ? toEntryType(rawEntry.type) : 'unknown',
      content: isRecord(rawEntry) ? toContent(rawEntry.content) : null,
    }
  })

  return { entries }
}

/**
 * Looks up a stream/storage by name — case-insensitively, matching either a
 * full path (`"Root Entry/WordDocument"`) or just its final path segment
 * (`"WordDocument"`). Every stream this package reads (`WordDocument`,
 * `0Table`/`1Table`, `PowerPoint Document`) sits directly under the root, so
 * matching on the last segment alone is unambiguous in practice; this is a
 * deliberately simplified subset of `XLSX.CFB.find`'s own lookup (which also
 * handles CFB's control-character name-escaping fallback — never observed on
 * these particular streams, so not worth carrying the extra complexity for).
 */
export function findEntry(cfb: CfbContainer, name: string): CfbEntry | null {
  const upperName = name.toUpperCase()

  for (const entry of cfb.entries) {
    if (entry.path.toUpperCase() === upperName) {
      return entry
    }
  }

  for (const entry of cfb.entries) {
    if (entry.name.toUpperCase() === upperName) {
      return entry
    }
  }

  return null
}
