/**
 * OFFICE-VERIFY-1 — does a real office suite actually read what Atlas writes?
 *
 * WHY THIS EXISTS. Until now, nothing Atlas had ever written had been opened in
 * Microsoft Office or LibreOffice. The whole DOCX write path rested on
 * `scripts/validate-office-file.mjs`, and the handoff notes are blunt about that
 * substitute: it passed two genuinely invalid files clean before attribute
 * datatype checks were added. Every round-trip guarantee in this repo was
 * self-asserted.
 *
 * WHAT THIS ASSERTS, and why it is stronger than "it opens". For each corpus
 * fixture, LibreOffice converts the ORIGINAL file and the same file after Atlas
 * has parsed and re-saved it, and the text of the two must match. That catches
 * content Atlas dropped, duplicated or mangled, judged by a consumer with no stake
 * in Atlas's own model. "It converted without an error" alone would prove little:
 * LibreOffice repairs quietly, so a successful conversion means openable, not
 * valid.
 *
 * WHY HTML AND NOT `txt:Text`, WHICH IS THE OBVIOUS CHOICE. Because in LibreOffice
 * 26.8.0 the plain-text filter hangs on any document containing a table. Bisected
 * to be certain, since it first looked like a problem with two of our fixtures:
 *   - stripping `tblPr`, `tblGrid`, `tblLook`, `tblBorders` and the `tblStyle`
 *     reference changed nothing; removing the `<w:tbl>` fixed it; injecting that
 *     same table into a fixture that converted fine broke it;
 *   - a hand-built, minimal, schema-clean `.docx` with one 1x1 table and one
 *     paragraph hangs, while the same file without the table converts;
 *   - that same minimal table file converts to `odt` and to `html` without
 *     trouble.
 * So it is one output filter, not our markup and not LibreOffice as a whole. It is
 * a real runaway, not slowness: it wrote 620 MB of temp output in 120 s and 4.6 GB
 * when left longer. Using the HTML filter keeps every fixture in the comparison,
 * tables included — which is the whole point, since tables are exactly where a
 * write path is most likely to lose something.
 *
 * WHAT IT STILL DOES NOT PROVE. LibreOffice is not Word — more forgiving in
 * places, stricter in others, and it reads none of the `w15:`/`w16:` extensions
 * this codebase spends real effort preserving. Text equality also says nothing
 * about formatting: a lost `w:noProof` or a flattened image z-order is invisible
 * here, which is why `fidelity/__tests__/lossySaveWarnings.test.ts` exists
 * alongside it. A floor under the write path, not a certificate.
 *
 * ONE PROCESS PER FILE, with a timeout and a cleanup, so a future filter bug in
 * one document cannot take the run down or fill the disk.
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

/** LibreOffice's Writer-to-HTML filter. See the header for why not `txt:Text`. */
const CONVERT_TO = 'html:HTML (StarWriter)'

/** Generous for documents this size; short enough to catch a runaway early. */
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

/** Leaves nothing running or half-written behind after a timeout. */
function cleanUpAfterTimeout(outDir: string): void {
  try {
    execFileSync('taskkill', ['/F', '/IM', 'soffice.bin', '/T'], { stdio: 'ignore' })
  } catch {
    // Not running, or not Windows. Either is fine.
  }
  for (const name of fs.readdirSync(outDir)) {
    if (name.endsWith('.tmp')) fs.rmSync(path.join(outDir, name), { force: true })
  }
}

/**
 * The document's visible text, as LibreOffice reads it. `null` when LibreOffice
 * could not read the file at all.
 *
 * Both sides of every comparison go through this same extraction, so a quirk in
 * it cancels out. The one failure mode that would NOT cancel out is extracting
 * nothing from both, which would make any pair compare equal — the caller guards
 * that by requiring non-empty text.
 */
function extractText(file: string, workDir: string): string | null {
  const outDir = path.join(workDir, 'html')
  fs.mkdirSync(outDir, { recursive: true })
  // A URL: the only form `-env:UserInstallation` accepts. Without a profile of our
  // own, a headless run shares one with any LibreOffice window the developer has
  // open, and then exits having done nothing.
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
        CONVERT_TO,
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

  const produced = path.join(outDir, `${path.basename(file, path.extname(file))}.html`)
  if (!fs.existsSync(produced)) return null
  return htmlToText(fs.readFileSync(produced, 'utf8'))
}

/** Stands in for a block boundary while all real whitespace is collapsed around it. */
const BLOCK_BREAK = '\u0000'

/**
 * The text content of LibreOffice's HTML output.
 *
 * `<head>` is dropped wholesale: it carries a `<title>`, a generator `<meta>` and
 * a `<style>` block, none of which is document content. The generator meta would
 * otherwise be compared as content, and it genuinely differs — the fixture says
 * "Atlas Corpus Generator" and Atlas's own save says "Atlas", with a current
 * timestamp.
 *
 * NEWLINES IN THE HTML SOURCE ARE NOT CONTENT. LibreOffice wraps its output, so
 * `<span>שלום עולם</span>` can come back as `שלום\nעולם` — insignificant
 * whitespace that renders as a single space. An earlier version of this preserved
 * those newlines and reported `rtl-text` as a real difference for exactly that
 * reason. So block boundaries are marked FIRST with a sentinel, then every run of
 * real whitespace is collapsed, and only then does the sentinel become a newline.
 * Getting that order wrong is the difference between "the text moved" and "the
 * HTML happened to wrap elsewhere".
 */
function htmlToText(html: string): string {
  const bodyStart = html.search(/<body\b[^>]*>/i)
  const body = bodyStart === -1 ? html : html.slice(html.indexOf('>', bodyStart) + 1)
  const text = body
    .replace(/<\/body>[\s\S]*$/i, '')
    .replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, '')
    // Block boundaries must survive, or "A" and "B" in adjacent cells would
    // concatenate into "AB" and a genuinely moved cell boundary would compare
    // equal.
    .replace(/<(?:\/p|\/h[1-6]|br|\/td|\/th|\/tr|\/li|\/div)\b[^>]*>/gi, BLOCK_BREAK)
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
    // Last, so an escaped entity in the source is not double-decoded.
    .replace(/&amp;/g, '&')

  return text
    .replace(/\s+/g, ' ')
    .split(BLOCK_BREAK)
    .map((line) => line.trim())
    .filter((line) => line !== '')
    .join('\n')
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
          failures.push(`${fixtureId}: LibreOffice could not read the UNTOUCHED fixture — not an Atlas failure`)
          continue
        }
        // Without this, a fixture both of whose sides extract to nothing would
        // compare equal and pass while proving nothing at all.
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
})
