/**
 * W1.1 — Type system tests.
 *
 * These are PRIMARILY compile-time tests: `npx tsc --noEmit` must fail if a
 * `FormatId` case is omitted from a switch covered by `assertNever`. The vitest
 * runtime assertions cover the few values we actually need at runtime
 * (`assertNever` throwing, `ALL_FORMAT_IDS` length / dedup).
 */

import { describe, it, expect } from 'vitest'
import {
  ALL_FORMAT_IDS,
  KNOWN_FORMAT_IDS,
  assertNever,
  type FormatId,
  type LoadedFile,
  type ViewerStats,
} from '../types'

describe('FormatId constants', () => {
  it('KNOWN_FORMAT_IDS has exactly 15 entries (every format except unknown)', () => {
    expect(KNOWN_FORMAT_IDS).toHaveLength(15)
  })

  it('ALL_FORMAT_IDS has exactly 16 entries and contains unknown', () => {
    expect(ALL_FORMAT_IDS).toHaveLength(16)
    expect(ALL_FORMAT_IDS).toContain('unknown')
  })

  it('ALL_FORMAT_IDS has no duplicates', () => {
    expect(new Set(ALL_FORMAT_IDS).size).toBe(ALL_FORMAT_IDS.length)
  })
})

describe('assertNever', () => {
  it('throws when called (only reachable if a switch is non-exhaustive at runtime)', () => {
    expect(() => assertNever('boom' as never)).toThrow(/assertNever/)
  })
})

describe('compile-time exhaustiveness (proves assertNever guards FormatId)', () => {
  // If a new FormatId is added without updating this switch, tsc fails.
  function describeFormat(f: FormatId): string {
    switch (f) {
      case 'markdown':
        return 'md'
      case 'docx':
        return 'docx'
      case 'xlsx':
        return 'xlsx'
      case 'pptx':
        return 'pptx'
      case 'pdf':
        return 'pdf'
      case 'csv':
        return 'csv'
      case 'tsv':
        return 'tsv'
      case 'text':
        return 'text'
      case 'code':
        return 'code'
      case 'odt':
        return 'odt'
      case 'ods':
        return 'ods'
      case 'odp':
        return 'odp'
      case 'rtf':
        return 'rtf'
      case 'doc':
        return 'doc'
      case 'ppt':
        return 'ppt'
      case 'unknown':
        return 'unknown'
      default:
        return assertNever(f)
    }
  }

  it('every FormatId in ALL_FORMAT_IDS is handled by the switch', () => {
    for (const f of ALL_FORMAT_IDS) {
      expect(describeFormat(f)).toBeDefined()
    }
  })
})

describe('LoadedFile discriminated union', () => {
  it('text variant carries string content', () => {
    const f: LoadedFile = { kind: 'text', content: '# hi', path: 'C:\\a.md', format: 'markdown' }
    expect(f.kind).toBe('text')
    if (f.kind === 'text') {
      expect(f.content).toBe('# hi')
    }
  })

  it('binary variant carries ArrayBuffer content', () => {
    const buf = new Uint8Array([0x25, 0x50, 0x44, 0x46]).buffer
    const f: LoadedFile = { kind: 'binary', content: buf, path: 'C:\\a.pdf', format: 'pdf' }
    expect(f.kind).toBe('binary')
    if (f.kind === 'binary') {
      expect(f.content.byteLength).toBe(4)
    }
  })
})

describe('ViewerStats discriminated union', () => {
  it('every stats kind is constructible and narrows correctly', () => {
    const cases: ViewerStats[] = [
      { kind: 'markdown', words: 100, headings: 5 },
      { kind: 'spreadsheet', sheet: 'Sheet1', rows: 10, cols: 3 },
      { kind: 'pdf', page: 1, pageCount: 12 },
      { kind: 'slides', slide: 1, slideCount: 8 },
      { kind: 'code', language: 'typescript', lines: 42 },
      { kind: 'text', lines: 10, chars: 100 },
      { kind: 'document', words: 500, pages: 3 },
    ]
    expect(cases).toHaveLength(7)
    for (const s of cases) {
      expect(s.kind).toBeDefined()
    }
  })
})
