/**
 * OFFICE-VERIFY-1 — does a real office suite actually read what Atlas writes?
 *
 * WHY THIS EXISTS. Nothing Atlas had ever written had been opened in Microsoft Office
 * or LibreOffice. The whole DOCX write path rested on
 * `scripts/validate-office-file.mjs`, and the handoff notes are blunt about that
 * substitute: it passed two genuinely invalid files clean before attribute datatype
 * checks were added. Every round-trip guarantee in this repo was self-asserted.
 *
 * WHAT THIS ASSERTS, and why it is stronger than "it opens". For each corpus fixture,
 * LibreOffice converts the ORIGINAL file and the same file after Atlas has parsed and
 * re-saved it, and the text of the two must match. That catches content Atlas dropped,
 * duplicated or mangled, judged by a consumer with no stake in Atlas's own model. "It
 * converted without an error" alone would prove little: LibreOffice repairs quietly,
 * so a successful conversion means openable, not valid.
 *
 * It has earned its keep: `FID-DEFAULTS-1` was found here and nowhere else. Collapsing
 * an empty `<w:rPrDefault/>`/`<w:pPrDefault/>` pair to `<w:docDefaults/>` is identical
 * by Atlas's own model, and made LibreOffice compute different document defaults.
 *
 * WHAT IT STILL DOES NOT PROVE. LibreOffice is not Word — more forgiving in places,
 * stricter in others, and it reads none of the `w15:`/`w16:` extensions this codebase
 * spends real effort preserving. Text equality also says nothing about formatting: a
 * lost `w:noProof` or a flattened image z-order is invisible here, which is why
 * `fidelity/__tests__/lossySaveWarnings.test.ts` exists alongside it. A floor under
 * the write path, not a certificate.
 *
 * The LibreOffice plumbing, and the several things about it that had to be learned the
 * hard way (one process per file, our own profile, insignificant newlines, the
 * `txt:Text` filter's runaway on tables), lives in `__tests__/helpers/libreOffice.ts`,
 * shared with the spreadsheet and slide harnesses.
 *
 * Skipped, not failed, when `soffice` is missing — CI has none, and a developer
 * without it must not be blocked. So a green suite does NOT imply this ran: check for
 * the skip. `scoop install extras/libreoffice`, per-user, no elevation.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { CONVERT, convert, findSoffice, htmlToText } from '../../__tests__/helpers/libreOffice'
import { loadDocx, saveDocx } from '../index'
import { listCorpusFixtureIds, readCorpusFixture } from './corpusRoundtripHelpers'

const SOFFICE = findSoffice()

/** `null` when LibreOffice could not read the file at all. */
function extractText(file: string, workDir: string): string | null {
  const converted = convert(SOFFICE!, file, CONVERT.writerHtml, workDir)
  return converted === null ? null : htmlToText(fs.readFileSync(converted, 'utf8'))
}

describe.skipIf(SOFFICE === null)('LibreOffice reads what Atlas writes (OFFICE-VERIFY-1)', () => {
  it('extracts the same text from an Atlas-saved .docx as from the original, for every corpus fixture', async () => {
    const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-'))
    const fixtureIds = await listCorpusFixtureIds()
    expect(fixtureIds.length).toBeGreaterThan(0)

    const failures: string[] = []
    try {
      for (const fixtureId of fixtureIds) {
        const original = path.join(workDir, `${fixtureId}.original.docx`)
        const resaved = path.join(workDir, `${fixtureId}.resaved.docx`)
        fs.writeFileSync(original, Buffer.from(await readCorpusFixture(fixtureId)))
        // No edits — the plain "open and save" path, which is the one a user takes
        // most often and the one every round-trip guarantee here is about.
        const bundle = await loadDocx(await readCorpusFixture(fixtureId))
        fs.writeFileSync(resaved, Buffer.from(await saveDocx(bundle)))

        const before = extractText(original, workDir)
        if (before === null) {
          // A statement about the fixture or about LibreOffice, not about Atlas.
          failures.push(`${fixtureId}: LibreOffice could not read the UNTOUCHED fixture — not an Atlas failure`)
          continue
        }
        // Two empty extractions compare equal, which would pass while proving nothing.
        if (before === '') {
          failures.push(`${fixtureId}: extracted NO text from the untouched fixture, so the comparison is vacuous`)
          continue
        }

        const after = extractText(resaved, workDir)
        if (after === null) {
          failures.push(`${fixtureId}: LibreOffice read the original but could NOT read the file Atlas saved`)
          continue
        }
        if (before !== after) {
          failures.push(
            `${fixtureId}: text differs after an Atlas save\n` +
              `  before: ${JSON.stringify(before.slice(0, 400))}\n` +
              `  after:  ${JSON.stringify(after.slice(0, 400))}`,
          )
        }
      }
    } finally {
      // Several hundred MB per run even when nothing goes wrong.
      fs.rmSync(workDir, { recursive: true, force: true })
    }

    expect(failures, `${failures.length} of ${fixtureIds.length} fixtures failed:\n${failures.join('\n')}`).toEqual([])
  }, 1_800_000)

  /**
   * Proves the comparison above can actually FAIL.
   *
   * It passes only because Atlas preserves the text, and there are several ways for a
   * check like this to quietly become a no-op instead: an extractor that returns the
   * same string for any input, a filter that writes nothing, a converted file nobody
   * reads. Each looks like a green run. So a document is saved with one paragraph's
   * text genuinely changed, and the comparison has to notice.
   */
  it('detects changed body text', async () => {
    const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-neg-'))
    try {
      const fixtureId = 'plain-paragraphs-styles'
      const original = path.join(workDir, 'original.docx')
      fs.writeFileSync(original, Buffer.from(await readCorpusFixture(fixtureId)))

      const bundle = await loadDocx(await readCorpusFixture(fixtureId))
      const section = bundle.document.sections[0]
      const paragraph = section?.blocks.find((block) => block.kind === 'paragraph')
      expect(paragraph, 'the fixture has no paragraph to change, so this proves nothing').toBeDefined()

      // One run's text, replaced in place. Everything else about the document is
      // untouched, so a difference can only come from this.
      const changedBundle = {
        ...bundle,
        document: {
          ...bundle.document,
          sections: [
            {
              ...section!,
              blocks: section!.blocks.map((block) =>
                block === paragraph && block.kind === 'paragraph'
                  ? {
                      ...block,
                      children: [
                        {
                          kind: 'run' as const,
                          children: [{ kind: 'text' as const, value: 'CHANGED-BY-THE-SELF-CHECK' }],
                        },
                      ],
                    }
                  : block,
              ),
            },
            ...bundle.document.sections.slice(1),
          ],
        },
      }

      const changed = path.join(workDir, 'changed.docx')
      fs.writeFileSync(changed, Buffer.from(await saveDocx(changedBundle)))

      const before = extractText(original, workDir)
      const after = extractText(changed, workDir)
      expect(before, 'LibreOffice could not read the untouched fixture').not.toBeNull()
      expect(after, 'LibreOffice could not read the deliberately changed document').not.toBeNull()
      expect(after, 'the comparison is blind to changed body text').not.toBe(before)
      expect(after).toContain('CHANGED-BY-THE-SELF-CHECK')
    } finally {
      fs.rmSync(workDir, { recursive: true, force: true })
    }
  }, 600_000)
})
