/**
 * T2/DAT-07 — CSV/TSV parsing helper (Papa Parse worker:true above a
 * threshold, resolved as a Promise either way).
 */
import Papa from 'papaparse'
import { describe, expect, it, vi } from 'vitest'

import { parseCsv } from '../csvParse'
import { CSV_WORKER_CHAR_THRESHOLD } from '../sizeThresholds'

describe('parseCsv', () => {
  it('parses a small CSV synchronously (no worker) and resolves with its rows', async () => {
    const result = await parseCsv('name,value\nAtlas,2\nPhase,1\n')

    expect(result.data).toEqual([
      ['name', 'value'],
      ['Atlas', '2'],
      ['Phase', '1'],
    ])
  })

  // T2/DAT-07's actual browser-facing guarantee — the main thread never runs
  // Papa's parse loop for a large file — comes entirely from this delegation
  // decision; jsdom has no real Worker to parse *inside*, but the *decision*
  // to hand off is plain, synchronous branching this test can verify
  // directly by spying on the options `Papa.parse` is actually called with.
  it('enables Papa Parse worker mode once content crosses the size threshold', async () => {
    const spy = vi.spyOn(Papa, 'parse')
    const bigContent = 'a,b\n'.repeat(Math.ceil(CSV_WORKER_CHAR_THRESHOLD / 4) + 1)
    expect(bigContent.length).toBeGreaterThanOrEqual(CSV_WORKER_CHAR_THRESHOLD)

    await parseCsv(bigContent)

    expect(spy).toHaveBeenCalledTimes(1)
    expect(spy.mock.calls[0][1]).toMatchObject({ worker: true })
    spy.mockRestore()
  })

  it('does not request worker mode for content under the size threshold', async () => {
    const spy = vi.spyOn(Papa, 'parse')

    await parseCsv('a,b\n1,2\n')

    expect(spy).toHaveBeenCalledTimes(1)
    expect(spy.mock.calls[0][1]).not.toMatchObject({ worker: true })
    spy.mockRestore()
  })

  it('honors a custom delimiter (TSV)', async () => {
    const result = await parseCsv('name\tvalue\nAtlas\t2\n', { delimiter: '\t' })

    expect(result.data).toEqual([
      ['name', 'value'],
      ['Atlas', '2'],
    ])
  })

  it('skips empty lines', async () => {
    const result = await parseCsv('a,b\n\nc,d\n')

    expect(result.data).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
  })
})
