import { describe, expect, it } from 'vitest'

import { viewerRegistry } from '../registry'
import { ALL_FORMAT_IDS } from '../types'

describe('viewerRegistry', () => {
  it('registers every format id', () => {
    const registryKeys = Object.keys(viewerRegistry).sort()

    expect(registryKeys).toEqual([...ALL_FORMAT_IDS].sort())
    expect(registryKeys).toMatchInlineSnapshot(`
      [
        "code",
        "csv",
        "doc",
        "docx",
        "markdown",
        "odp",
        "ods",
        "odt",
        "pdf",
        "ppt",
        "pptx",
        "rtf",
        "text",
        "tsv",
        "unknown",
        "xlsx",
      ]
    `)
  })

  it('exposes loader functions for every format', () => {
    for (const loader of Object.values(viewerRegistry)) {
      expect(typeof loader).toBe('function')
    }
  })

  it('loads the unknown viewer placeholder', async () => {
    const unknownViewer = await viewerRegistry.unknown()

    expect(typeof unknownViewer === 'function' || typeof unknownViewer === 'object').toBe(true)
  })
})
