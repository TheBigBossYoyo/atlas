/**
 * T8 (DAT-19) — ODT fixture-corpus snapshot guard.
 *
 * odf-kit is a pre-1.0 dependency (now pinned to an exact version —
 * package.json) with, until this task, no local fixtures exercising its
 * output. Each fixture in `__fixtures__/odt-corpus/` (built by
 * `scripts/generate-odt-corpus.mjs`) isolates one real-world ODF feature;
 * snapshotting its rendered HTML here means a future odf-kit version bump
 * that silently changes output shows up as a snapshot diff to review,
 * instead of shipping unnoticed.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { readOdt } from 'odf-kit/reader'
import { describe, expect, it } from 'vitest'

import { parseLengthToPx } from '../shared/pageEstimate'

const CORPUS_DIR = path.resolve(process.cwd(), 'src/viewers/__fixtures__/odt-corpus')

function readFixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(path.join(CORPUS_DIR, `${name}.odt`)))
}

const FIXTURE_IDS = ['tables-and-lists', 'embedded-image', 'tracked-changes'] as const

describe('ODT fixture corpus (T8/DAT-19)', () => {
  it.each(FIXTURE_IDS)('renders %s to stable HTML', (id) => {
    const doc = readOdt(readFixtureBytes(id), { trackedChanges: 'changes' })
    const html = doc.toHtml({ fragment: true, trackedChanges: 'changes' })
    expect(html).toMatchSnapshot()
  })

  it('tables-and-lists renders a real <table>, <ul>, and <ol>', () => {
    const doc = readOdt(readFixtureBytes('tables-and-lists'), {})
    const html = doc.toHtml({ fragment: true })
    expect(html).toContain('<table>')
    expect(html).toContain('<ul>')
    expect(html).toContain('<ol>')
  })

  it('embedded-image renders an <img> with a resolved data: src', () => {
    const doc = readOdt(readFixtureBytes('embedded-image'), {})
    const html = doc.toHtml({ fragment: true })
    expect(html).toMatch(/<img src="data:image\/png;base64,[^"]+"/)
  })

  it('tracked-changes renders visible <ins>/<del> markup in "changes" mode', () => {
    const doc = readOdt(readFixtureBytes('tracked-changes'), { trackedChanges: 'changes' })
    const html = doc.toHtml({ fragment: true, trackedChanges: 'changes' })
    expect(html).toContain('<ins')
    expect(html).toContain('<del')
    expect(html).toContain('This paragraph was inserted by a reviewer.')
    expect(html).toContain('This paragraph was deleted.')
  })

  it('tracked-changes renders as fully accepted in "final" (default) mode', () => {
    const doc = readOdt(readFixtureBytes('tracked-changes'), {})
    const html = doc.toHtml({ fragment: true })
    expect(html).not.toContain('<ins')
    expect(html).not.toContain('<del')
    expect(html).toContain('This paragraph was inserted by a reviewer.')
    expect(html).not.toContain('This paragraph was deleted.')
  })

  it('every corpus fixture declares a US-Letter page layout (matches DEFAULT_PAGE_HEIGHT_PX)', () => {
    for (const id of FIXTURE_IDS) {
      const doc = readOdt(readFixtureBytes(id), {})
      const heightPx = parseLengthToPx(doc.pageLayout?.height)
      expect(heightPx).toBeCloseTo(11 * 96, 0)
    }
  })
})
