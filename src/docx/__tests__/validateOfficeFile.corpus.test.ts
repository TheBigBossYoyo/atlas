// Atlas — wires `scripts/lib/officeValidator.mjs` (the spec-level OPC/OOXML
// validation harness) into CI against every DOCX corpus fixture, saved
// through Atlas's REAL `loadDocx`/`saveDocx`/editor-command save path — the
// same fixtures and edit `roundtrip.corpus.test.ts` exercises, but checked
// against the actual OOXML package rules (content-types coverage, dangling
// r:id, XML-1.0-illegal characters, namespace prefixes declared, w:pPr/
// w:tblGrid/w:tblPr element order) rather than `roundtrip.corpus.test.ts`'s
// narrower well-formedness + relationship-target checks.
//
// This is deliberately a SEPARATE suite (not folded into
// `roundtrip.corpus.test.ts`) so the two can be read independently: this
// one asks "would Office refuse this file", that one asks "did the save
// preserve content".
import { describe, expect, it } from 'vitest'

import { applyCommand } from '../editor'
import type { Position } from '../editor'
import { loadDocx, saveDocx } from '../index'
import { listCorpusFixtureIds, readCorpusFixture } from './corpusRoundtripHelpers'

import { validateOfficeFile } from '../../../scripts/lib/officeValidator.mjs'

const FIRST_PARAGRAPH_START: Position = { paragraphPath: [0], runIndex: 0, charOffset: 0 }

function errorsOnly(issues: ReadonlyArray<{ readonly severity: string }>) {
  return issues.filter((i) => i.severity === 'error')
}

// Real top-level await (module scope) — resolved once before any `describe`
// callback runs, so the callbacks themselves stay synchronous.
const FIXTURE_IDS = await listCorpusFixtureIds()

describe('DOCX corpus saves pass spec-level OPC/OOXML validation', () => {
  it('found at least one corpus fixture', () => {
    expect(FIXTURE_IDS.length).toBeGreaterThan(0)
  })

  it.for(FIXTURE_IDS)('%s — no-op save', async (fixtureId) => {
    const buffer = await readCorpusFixture(fixtureId)
    const bundle = await loadDocx(buffer)
    const saved = await saveDocx(bundle)

    const result = validateOfficeFile(Buffer.from(saved))
    expect(result.format).toEqual({ family: 'opc', kind: 'docx' })
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it.for(FIXTURE_IDS)('%s — trivial-edit save', async (fixtureId) => {
    const buffer = await readCorpusFixture(fixtureId)
    const bundle = await loadDocx(buffer)
    const edited = applyCommand(bundle.document, { kind: 'insert-text', at: FIRST_PARAGRAPH_START, text: 'EDITED: ' })
    const saved = await saveDocx({ ...bundle, document: edited.document })

    const result = validateOfficeFile(Buffer.from(saved))
    expect(result.format).toEqual({ family: 'opc', kind: 'docx' })
    expect(errorsOnly(result.issues)).toEqual([])
  })
})
