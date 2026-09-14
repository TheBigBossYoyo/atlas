/**
 * Integration test for D17/DXE-11's settings.xml round-trip: the DOCX
 * editor's Track Changes toggle should reflect a real document's
 * `word/settings.xml` on load, and a toggle should persist through
 * `saveDocx` instead of the passthrough copy of the original part silently
 * winning.
 */
import { describe, expect, it } from 'vitest'

import { loadDocx, saveDocx, type DocxBundle } from '..'
import { loadRawPackage, readCorpusFixture } from './corpusRoundtripHelpers'

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return Uint8Array.from(bytes).buffer
}

describe('DOCX settings.xml Track Changes round-trip', () => {
  it('loads a fixture with no trackChanges element as trackChanges: false', async () => {
    const buffer = await readCorpusFixture('plain-paragraphs-styles')
    const bundle = await loadDocx(buffer)

    expect(bundle.settings?.trackChanges ?? false).toBe(false)
  })

  it('does not touch word/settings.xml when the toggle is never flipped', async () => {
    const buffer = await readCorpusFixture('plain-paragraphs-styles')
    const bundle = await loadDocx(buffer)
    const originalPkg = await loadRawPackage(buffer)

    const saved = await saveDocx(bundle)
    const savedPkg = await loadRawPackage(toArrayBuffer(saved))

    expect(savedPkg.get('word/settings.xml')?.toString('utf-8')).toEqual(
      originalPkg.get('word/settings.xml')?.toString('utf-8'),
    )
  })

  it('persists an explicit trackChanges: true toggle into the saved word/settings.xml', async () => {
    const buffer = await readCorpusFixture('plain-paragraphs-styles')
    const bundle = await loadDocx(buffer)

    const toggledBundle: DocxBundle = { ...bundle, settings: { trackChanges: true } }
    const saved = await saveDocx(toggledBundle)
    const savedPkg = await loadRawPackage(toArrayBuffer(saved))

    expect(savedPkg.get('word/settings.xml')?.toString('utf-8')).toContain('<w:trackChanges/>')

    const reloaded = await loadDocx(toArrayBuffer(saved))
    expect(reloaded.settings?.trackChanges).toBe(true)
  })

  it('turns trackChanges back off and removes the element on save', async () => {
    const buffer = await readCorpusFixture('plain-paragraphs-styles')
    const bundle = await loadDocx(buffer)

    const onBundle: DocxBundle = { ...bundle, settings: { trackChanges: true } }
    const onBytes = await saveDocx(onBundle)
    const onReloaded = await loadDocx(toArrayBuffer(onBytes))
    expect(onReloaded.settings?.trackChanges).toBe(true)

    const offBundle: DocxBundle = { ...onReloaded, settings: { trackChanges: false } }
    const offBytes = await saveDocx(offBundle)
    const offPkg = await loadRawPackage(toArrayBuffer(offBytes))

    expect(offPkg.get('word/settings.xml')?.toString('utf-8')).not.toContain('trackChanges')

    const offReloaded = await loadDocx(toArrayBuffer(offBytes))
    expect(offReloaded.settings?.trackChanges).toBe(false)
  })
})
