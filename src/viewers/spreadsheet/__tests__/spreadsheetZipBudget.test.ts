/**
 * USR-17 (security review) — the spreadsheet passthrough readers refuse a
 * zip whose declared uncompressed size is implausibly large, before
 * inflating anything.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { loadWorkbookZip, SpreadsheetZipBombError } from '../spreadsheetZipBudget'

async function buildZip(entryBytes: number): Promise<ArrayBuffer> {
  const zip = new JSZip()
  // Highly compressible content (all zeros) — a small compressed size next
  // to a large declared/actual uncompressed size, the shape a real zip bomb
  // takes.
  zip.file('big.xml', new Uint8Array(entryBytes))
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
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
