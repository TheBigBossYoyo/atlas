/**
 * OFFICE-VERIFY-1 — does a real office suite actually read what Atlas writes?
 *
 * WHY THIS EXISTS. Until now, nothing Atlas has ever written had been opened in
 * Microsoft Office or LibreOffice. The whole DOCX write path rested on
 * `scripts/validate-office-file.mjs`, and the handoff notes are blunt about that
 * substitute: it passed two genuinely invalid files clean before attribute
 * datatype checks were added, so a clean validator run is necessary, not
 * sufficient. Every round-trip guarantee in this repo was self-asserted.
 *
 * WHAT THIS ASSERTS, and why it is stronger than "it opens". For each corpus
 * fixture, LibreOffice extracts the text of the ORIGINAL file and of the same
 * file after Atlas has parsed and re-saved it, and the two must match. That
 * catches content Atlas dropped, duplicated or mangled, judged by a consumer with
 * no stake in Atlas's own model. "It converted without an error" alone would prove
 * little: LibreOffice repairs quietly, so a successful conversion means openable,
 * not valid.
 *
 * WHAT IT STILL DOES NOT PROVE. LibreOffice is not Word — more forgiving in
 * places, stricter in others, and it reads none of the `w15:`/`w16:` extensions
 * this codebase spends real effort preserving. Text equality also says nothing
 * about formatting: a lost `w:noProof` or a flattened image z-order is invisible
 * here. A floor under the write path, not a certificate. Word itself is untested.
 *
 * ONE FILE AT A TIME, WITH A TIMEOUT AND A CLEANUP — and the reason is specific.
 * `table-fixed-grid-merged-cells.docx`, an ORIGINAL fixture of 3.8 KB containing
 * one three-row table, sends LibreOffice 26.8.0 into a runaway loop when converted
 * to plain text: 620 MB of temp output in 120 seconds, 4.6 GB when left longer.
 * Nothing to do with Atlas — it does the same to the untouched fixture. A single
 * batch invocation for all files therefore cannot work (one bad file takes the
 * whole run down and fills the disk), so each conversion is its own process with
 * its own timeout, and a timeout deletes the partial output instead of leaving
 * gigabytes behind. `LIBREOFFICE_RUNAWAY_FIXTURES` lists what is known to do this.
 *
 * HOW IT IS GATED. Skipped, not failed, when `soffice` is missing — CI has no
 * LibreOffice and a developer without it must not be blocked. So a green suite
 * does NOT imply this ran: check for the skip. Install with
 * `scoop install extras/libreoffice` (per-user, no elevation).
 */
import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { loadDocx, saveDocx } from '../index'
import { listCorpusFixtureIds, readCorpusFixture } from './corpusRoundtripHelpers'

/**
 * Fixtures LibreOffice 26.8.0 cannot convert, verified against the UNTOUCHED
 * original with a fresh profile — so this is LibreOffice's behaviour on the
 * fixture, not a consequence of anything Atlas did.
 *
 * Both entries are the corpus's only two TABLE fixtures, which has a consequence
 * worth stating plainly rather than hiding behind a green test: **tables are not
 * externally verified at all by this harness.** Everything it proves about the
 * write path, it proves for documents without tables.
 *
 * Cause unknown. The obvious suspect was the `<w:tblW w:type="auto" w:w="100"/>`
 * both share — a width that OOXML itself defines as meaningless when the type is
 * `auto` (see `fidelity/__tests__/lossySaveWarnings.test.ts`'s note 6). Tested by
 * rebuilding the fixture with `w:w="0"`, with `type="dxa" w:w="9000"` and with
 * `type="pct"`: none of them converted either, so that hypothesis is wrong and is
 * recorded here so nobody spends the same hour on it again.
 *
 * Excluded rather than silently skipped: the assertion message names them on every
 * run, so the list cannot quietly grow into "we stopped checking". Re-test on a
 * LibreOffice upgrade; if one starts working, delete it from here.
 */
const LIBREOFFICE_RUNAWAY_FIXTURES: ReadonlySet<string> = new Set([
  // 3.8 KB, one three-row table, one gridSpan and two vMerge. Converting it to
  // plain text wrote 620 MB of temp output in 120 s, and 4.6 GB when left longer.
  'table-fixed-grid-merged-cells',
  // Same shape, plus a table style. Times out identically with a fresh profile.
  'table-styled-banded',
])

/** Generous for a document of this size, short enough that a runaway is caught early. */
const CONVERT_TIMEOUT_MS = 90_000

/** Where `scoop install extras/libreoffice` puts it, then the usual system paths. */
function findSoffice(): string | null {
  const candidates = [
    path.join(os.homedir(), 'scoop', 'apps', 'libreoffice', 'current', 'LibreOffice', 'program', 'soffice.exe'),
    'C:/Program Files/LibreOffice/program/soffice.exe',
    'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
    '/usr/bin/soffice',
    '/usr/local/bin/soffice',
  ]
  return candidates.find((candidate) => fs.existsSync(candidate)) ?? null
}

const SOFFICE = findSoffice()

/** Leaves nothing running or half-written after a timeout — see the header. */
function cleanUpAfterTimeout(outDir: string): void {
  try {
    execFileSync('taskkill', ['/F', '/IM', 'soffice.bin', '/T'], { stdio: 'ignore' })
  } catch {
    // Not running, or not Windows. Either is fine.
  }
  for (const name of fs.readdirSync(outDir)) {
    // The runaway writes its output to a `.tmp` beside the real target.
    if (name.endsWith('.tmp')) fs.rmSync(path.join(outDir, name), { force: true })
  }
}

/**
 * The document's text as LibreOffice reads it, or `null` when it could not read
 * it at all — no output file, or the conversion timed out.
 */
function extractText(file: string, workDir: string): string | null {
  const outDir = path.join(workDir, 'txt')
  fs.mkdirSync(outDir, { recursive: true })
  // A URL: the only form `-env:UserInstallation` accepts. Without a profile of
  // our own, a headless run shares one with any LibreOffice window the developer
  // has open, and then exits having done nothing.
  const profileUrl = `file:///${path.join(workDir, 'profile').replace(/\\/g, '/')}`

  try {
    execFileSync(
      SOFFICE!,
      [
        '--headless',
        '--norestore',
        '--nolockcheck',
        `-env:UserInstallation=${profileUrl}`,
        '--convert-to',
        'txt:Text',
        '--outdir',
        outDir,
        file,
      ],
      { stdio: 'pipe', timeout: CONVERT_TIMEOUT_MS },
    )
  } catch {
    cleanUpAfterTimeout(outDir)
    return null
  }

  const produced = path.join(outDir, `${path.basename(file, path.extname(file))}.txt`)
  // An empty document produces an empty file, not no file, so existence is the
  // right test for "could LibreOffice read this".
  return fs.existsSync(produced) ? fs.readFileSync(produced, 'utf8') : null
}

/**
 * Collapses whitespace, and nothing else. Normalising punctuation or case would
 * let a real corruption through.
 */
function normalise(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/[ \t]+/g, ' ').replace(/\n{2,}/g, '\n').trim()
}

describe.skipIf(SOFFICE === null)('LibreOffice reads what Atlas writes (OFFICE-VERIFY-1)', () => {
  it('extracts the same text from an Atlas-saved .docx as from the original, for every corpus fixture', async () => {
    const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-'))
    const allFixtureIds = await listCorpusFixtureIds()
    expect(allFixtureIds.length).toBeGreaterThan(0)

    const fixtureIds = allFixtureIds.filter((id) => !LIBREOFFICE_RUNAWAY_FIXTURES.has(id))
    const excluded = allFixtureIds.filter((id) => LIBREOFFICE_RUNAWAY_FIXTURES.has(id))
    // Reported, not silent: an exclusion list that nobody sees is how "we stopped
    // checking" happens.
    expect(fixtureIds.length, `every fixture is excluded; excluded: ${excluded.join(', ')}`).toBeGreaterThan(0)

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
          // Reported so it gets added to the exclusion list deliberately.
          failures.push(
            `${fixtureId}: LibreOffice could not read the UNTOUCHED fixture — not an Atlas failure; ` +
              `add it to LIBREOFFICE_RUNAWAY_FIXTURES with a note if that is expected`,
          )
          continue
        }

        const after = extractText(resaved, workDir)
        if (after === null) {
          failures.push(`${fixtureId}: LibreOffice read the original but could NOT read the file Atlas saved`)
          continue
        }
        if (normalise(before) !== normalise(after)) {
          failures.push(
            `${fixtureId}: text differs after an Atlas save\n` +
              `  before: ${JSON.stringify(normalise(before).slice(0, 400))}\n` +
              `  after:  ${JSON.stringify(normalise(after).slice(0, 400))}`,
          )
        }
      }
    } finally {
      // Several hundred MB per run even when nothing goes wrong.
      fs.rmSync(workDir, { recursive: true, force: true })
    }

    expect(
      failures,
      `${failures.length} of ${fixtureIds.length} fixtures failed ` +
        `(${excluded.length} excluded: ${excluded.join(', ') || 'none'}):\n${failures.join('\n')}`,
    ).toEqual([])
  }, 1_800_000)
})
