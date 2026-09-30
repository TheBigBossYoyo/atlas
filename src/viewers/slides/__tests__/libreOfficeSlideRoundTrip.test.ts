/**
 * OFFICE-VERIFY-3 — does a real office suite read what Atlas writes for SLIDES?
 *
 * Third of the round-trip harnesses, after DOCX (which found `FID-DEFAULTS-1`) and
 * spreadsheets. Until now the slide write path had no outside opinion at all.
 *
 * WHAT IS ASSERTED: the slide text LibreOffice reads, in order, from every slide.
 * Not layout, not theme, not shape geometry — text equality says nothing about
 * those. What it catches is text that vanished, duplicated, changed slide or came
 * back in a different order, judged by something with no stake in Atlas's model.
 *
 * WHY FLAT ODP. Impress's HTML export writes a set of files per deck, and its PDF
 * export would need a PDF text extractor in the middle; `.fodp` is one XML file
 * whose `text:p` elements are exactly the slide text. Chosen for that reason, not
 * because it resembles anything Atlas writes.
 *
 * The slide save path is pleasingly small — `loadOfficePackage` then
 * `writeOfficePackage`, with no passthrough/rebuild split of the kind
 * `useSpreadsheetEditor` has — so a no-edit save is the whole of what a user's
 * Ctrl+S does to an untouched deck, and that is what is exercised.
 *
 * Skipped, not failed, when `soffice` is missing — CI has none. A green suite does
 * not imply this ran. `scoop install extras/libreoffice`, per-user, no elevation.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { CONVERT, convert, findSoffice, flatOdfText } from '../../../__tests__/helpers/libreOffice'
import { loadOfficePackage, readPart, withParts, writeOfficePackage } from '../../../office/officePackage'

const SOFFICE = findSoffice()
const FIXTURE_DIR = path.resolve(process.cwd(), 'tests/e2e/fixtures')

const FIXTURES = ['sample.pptx', 'sample-multislide.pptx', 'sample.odp'] as const

function readFixture(name: string): ArrayBuffer {
  // Rebuilt through the realm's own `Uint8Array`: JSZip rejects a Node `Buffer`'s
  // backing `ArrayBuffer` under jsdom (see `docx/__tests__/corpusRoundtripHelpers.ts`).
  return Uint8Array.from(fs.readFileSync(path.join(FIXTURE_DIR, name))).buffer
}

function textOf(file: string, workDir: string, label: string): string {
  const converted = convert(SOFFICE!, file, CONVERT.impressFlatOdp, workDir)
  expect(converted, `LibreOffice could not read ${label}`).not.toBeNull()
  return flatOdfText(fs.readFileSync(converted!, 'utf8'))
}

describe.skipIf(SOFFICE === null)('LibreOffice reads the slides Atlas writes (OFFICE-VERIFY-3)', () => {
  it.each(FIXTURES)('%s round-trips its slide text through a no-edit save', async (name) => {
    const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-slide-'))
    try {
      const original = path.join(workDir, name)
      fs.writeFileSync(original, Buffer.from(readFixture(name)))

      const pkg = await loadOfficePackage(readFixture(name))
      const resaved = path.join(workDir, `resaved-${name}`)
      fs.writeFileSync(resaved, Buffer.from(await writeOfficePackage(pkg)))

      const before = textOf(original, workDir, `the UNTOUCHED ${name}`)
      // Two empty extractions compare equal, which would pass while proving nothing.
      expect(before, `no slide text extracted from ${name}`).not.toBe('')

      expect(textOf(resaved, workDir, `the file Atlas saved for ${name}`), 'slide text changed on save').toBe(before)
    } finally {
      fs.rmSync(workDir, { recursive: true, force: true })
    }
  }, 600_000)

  /**
   * Proves the comparison above can actually FAIL.
   *
   * A round-trip check that passes only means something if it would notice a change,
   * and there are several ways for one to quietly become a no-op: an extractor that
   * returns the same string for any input, a filter that writes nothing, a converted
   * file nobody reads. Each of those looks like a green run. So one slide's text is
   * deliberately changed and the check has to catch it.
   */
  it('detects changed slide text', async () => {
    const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-slide-neg-'))
    try {
      const original = path.join(workDir, 'sample.pptx')
      fs.writeFileSync(original, Buffer.from(readFixture('sample.pptx')))

      const pkg = await loadOfficePackage(readFixture('sample.pptx'))
      const slidePath = 'ppt/slides/slide1.xml'
      const slideXml = readPart(pkg, slidePath)
      expect(slideXml, `${slidePath} is missing, so this proves nothing`).not.toBeNull()
      expect(slideXml, 'the fixture text this test rewrites has changed').toContain('Atlas PPTX fixture')

      const mutated = withParts(pkg, {
        [slidePath]: slideXml!.replace('Atlas PPTX fixture', 'CHANGED-BY-THE-SELF-CHECK'),
      })
      const changed = path.join(workDir, 'changed.pptx')
      fs.writeFileSync(changed, Buffer.from(await writeOfficePackage(mutated)))

      const before = textOf(original, workDir, 'the untouched fixture')
      const after = textOf(changed, workDir, 'the deliberately changed deck')

      expect(after, 'the comparison is blind to changed slide text').not.toBe(before)
      expect(after).toContain('CHANGED-BY-THE-SELF-CHECK')
    } finally {
      fs.rmSync(workDir, { recursive: true, force: true })
    }
  }, 600_000)
})
