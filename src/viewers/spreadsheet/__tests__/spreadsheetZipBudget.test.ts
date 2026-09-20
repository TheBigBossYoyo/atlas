/**
 * USR-17 (security review) — the spreadsheet passthrough readers refuse a
 * zip whose declared uncompressed size is implausibly large, before
 * inflating anything.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { checkWorkbookZipBudgetSync, loadWorkbookZip, SpreadsheetZipBombError } from '../spreadsheetZipBudget'

async function buildZip(entryBytes: number): Promise<ArrayBuffer> {
  const zip = new JSZip()
  // Highly compressible content (all zeros) — a small compressed size next
  // to a large declared/actual uncompressed size, the shape a real zip bomb
  // takes.
  zip.file('big.xml', new Uint8Array(entryBytes))
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
}

/** Locates the (single) central directory file header's signature (`PK\x01\x02`) in a real, freshly-generated zip buffer. */
function findCentralDirectoryHeaderOffset(bytes: Uint8Array): number {
  for (let i = 0; i < bytes.length - 4; i++) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x01 && bytes[i + 3] === 0x02) {
      return i
    }
  }
  throw new Error('central directory header not found in test fixture')
}

/**
 * Rewrites a real zip's central-directory-declared uncompressed size to
 * `declaredSize`, WITHOUT touching the actual (small, real) compressed data
 * — exactly the "small compressed archive, huge declared uncompressed size"
 * shape a crafted zip bomb takes, and exactly what `checkWorkbookZipBudgetSync`
 * must catch by reading the declared size alone, before anything is inflated.
 */
function withDeclaredUncompressedSize(buffer: ArrayBuffer, declaredSize: number): ArrayBuffer {
  const bytes = new Uint8Array(buffer.slice(0))
  const headerOffset = findCentralDirectoryHeaderOffset(bytes)
  new DataView(bytes.buffer).setUint32(headerOffset + 24, declaredSize, true)
  return bytes.buffer
}

describe('loadWorkbookZip', () => {
  it('loads a normal-sized file exactly like plain JSZip.loadAsync', async () => {
    const buffer = await buildZip(1024)
    const zip = await loadWorkbookZip(buffer)
    expect(await zip.file('big.xml')!.async('uint8array')).toHaveLength(1024)
  })

  it('rejects a single entry whose declared size is over the per-entry budget', async () => {
    const buffer = await buildZip(10_000)
    await expect(loadWorkbookZip(buffer, 1_000, 1_000_000)).rejects.toBeInstanceOf(SpreadsheetZipBombError)
  })

  it('rejects a combined declared size over the total budget, even under the per-entry limit', async () => {
    const zip = new JSZip()
    zip.file('a.xml', new Uint8Array(600))
    zip.file('b.xml', new Uint8Array(600))
    const buffer = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
    await expect(loadWorkbookZip(buffer, 1_000, 1_000)).rejects.toBeInstanceOf(SpreadsheetZipBombError)
  })

  it('does not reject at the default (large) budgets', async () => {
    const buffer = await buildZip(10_000)
    await expect(loadWorkbookZip(buffer)).resolves.toBeInstanceOf(JSZip)
  })
})

describe('checkWorkbookZipBudgetSync — SHEET-1 (parseWorkbookBuffer\'s synchronous guard)', () => {
  it('does not throw for a normal, honestly-declared zip', async () => {
    const buffer = await buildZip(1024)
    expect(() => checkWorkbookZipBudgetSync(buffer)).not.toThrow()
  })

  it('rejects a crafted central directory whose declared size is over the per-entry budget, even though the real bytes are tiny', async () => {
    const buffer = await buildZip(10)
    const crafted = withDeclaredUncompressedSize(buffer, 10_000)
    expect(() => checkWorkbookZipBudgetSync(crafted, 1_000, 1_000_000)).toThrow(SpreadsheetZipBombError)
  })

  it('rejects a crafted combined declared size over the total budget, even under the per-entry limit', async () => {
    const zip = new JSZip()
    zip.file('a.xml', new Uint8Array(10))
    const buffer = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
    const crafted = withDeclaredUncompressedSize(buffer, 900)
    expect(() => checkWorkbookZipBudgetSync(crafted, 1_000, 500)).toThrow(SpreadsheetZipBombError)
  })

  it('does not reject a crafted-but-still-under-budget declared size', async () => {
    const buffer = await buildZip(10)
    const crafted = withDeclaredUncompressedSize(buffer, 900)
    expect(() => checkWorkbookZipBudgetSync(crafted, 1_000, 1_000_000)).not.toThrow()
  })

  it('does not reject at the default (large) budgets for a real, small workbook-shaped zip', async () => {
    const buffer = await buildZip(10_000)
    expect(() => checkWorkbookZipBudgetSync(buffer)).not.toThrow()
  })
})
