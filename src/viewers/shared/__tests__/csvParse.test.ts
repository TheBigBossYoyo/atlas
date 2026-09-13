/**
 * T2/DAT-07 — CSV/TSV parsing helper (Papa Parse worker:true above a
 * threshold, resolved as a Promise either way).
 */
import { describe, expect, it } from 'vitest'

import { parseCsv } from '../csvParse'

describe('parseCsv', () => {
  it('parses a small CSV synchronously (no worker) and resolves with its rows', async () => {
    const result = await parseCsv('name,value\nAtlas,2\nPhase,1\n')

    expect(result.data).toEqual([
      ['name', 'value'],
      ['Atlas', '2'],
      ['Phase', '1'],
    ])
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
