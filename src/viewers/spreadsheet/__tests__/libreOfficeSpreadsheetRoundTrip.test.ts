/**
 * OFFICE-VERIFY-2 — does a real office suite read what Atlas writes for SPREADSHEETS?
 *
 * The DOCX equivalent (`src/docx/__tests__/libreOfficeRoundTrip.test.ts`) found a
 * genuine fidelity bug nothing in Atlas's own model could see. Until now the
 * spreadsheet write path had no outside opinion at all.
 *
 * BOTH SAVE BRANCHES, because only one of them is the safe one. `useSpreadsheetEditor`'s
 * `writeToDisk` first tries `writeWorkbookThroughOriginal`, which patches the zip the
 * user opened and so preserves everything Atlas does not model; it falls back to
 * `writeWorkbookBytesWithTables`, which rebuilds the workbook from scratch and is
 * documented as dropping styles, charts and filters (see `docs/KNOWN_LIMITATIONS.md`).
 * Verifying only the passthrough would be verifying the branch that by design changes
 * least. The fallback is exercised here for every fixture, including the `.xlsx` ones
 * where real saves would not use it, precisely because it is the lossy one — and
 * because SHEET-4 exists: the passthrough returns `null` on several genuine internal
 * failures, which silently routes a real save down this path.
 *
 * WHAT IS ASSERTED: the CELL TEXT LibreOffice reads, every sheet of it. Not formatting
 * — the fallback writer is known to lose some, and text equality says nothing about it
 * either way. What this catches is content that moved, vanished, duplicated or landed
 * in the wrong sheet, judged by something with no stake in Atlas's model.
 *
 * The Calc HTML filter is used rather than CSV because CSV exports only the FIRST
 * sheet, which would have made `sample-multisheet.xlsx` look fine no matter what
 * happened to its second sheet.
 *
 * Skipped, not failed, when `soffice` is missing — CI has none. A green suite does not
 * imply this ran. `scoop install extras/libreoffice`, per-user, no elevation.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { CONVERT, convert, findSoffice, htmlToText } from '../../../__tests__/helpers/libreOffice'
import { attachFrozenPanes, attachSheetSources, attachTables, parseWorkbookBuffer } from '../../shared/spreadsheetGrid'
import { readFrozenPanes } from '../spreadsheetPanes'
import { readSheetPartPaths, readSheetTables } from '../spreadsheetTables'
import { createDocument, setCellValue } from '../spreadsheetDocument'
import { writeWorkbookBytesWithTables } from '../spreadsheetWrite'
import { writeWorkbookThroughOriginal } from '../xlsxPassthrough'

const SOFFICE = findSoffice()
const FIXTURE_DIR = path.resolve(process.cwd(), 'tests/e2e/fixtures')

type Fixture = {
  readonly file: string
  readonly bookType: 'xlsx' | 'ods'
  /** Whether a real save of this file would take the passthrough branch. */
  readonly passthroughEligible: boolean
}

const FIXTURES: ReadonlyArray<Fixture> = [
  { file: 'sample.xlsx', bookType: 'xlsx', passthroughEligible: true },
  { file: 'sample-multisheet.xlsx', bookType: 'xlsx', passthroughEligible: true },
  // `.ods` has no passthrough writer at all (`PASSTHROUGH_BOOK_TYPES` is OOXML only),
  // so every real save of one goes through the rebuild branch.
  { file: 'sample.ods', bookType: 'ods', passthroughEligible: false },
]

/**
 * Loads a fixture exactly as `useSpreadsheetWorkbook` does, so what is saved below
 * comes from the same document the app would have. Reproducing the four-step
 * `attach*` chain matters: `attachSheetSources` is what records the `sourcePath` the
 * passthrough writer keys off, so skipping it would silently test the fallback twice.
 */
async function loadDocument(buffer: ArrayBuffer): ReturnType<typeof createDocument> extends never ? never : Promise<ReturnType<typeof createDocument>> {
  const sheets = parseWorkbookBuffer(buffer)
  const [paneMap, tableMap, partPaths] = await Promise.all([
    readFrozenPanes(buffer),
    readSheetTables(buffer),
    readSheetPartPaths(buffer),
  ])
  return createDocument(attachSheetSources(attachTables(attachFrozenPanes(sheets, paneMap), tableMap), partPaths))
}

function readFixture(name: string): ArrayBuffer {
  // Rebuilt through the realm's own `Uint8Array`, for the reason
  // `docx/__tests__/corpusRoundtripHelpers.ts` documents: JSZip rejects a Node
  // `Buffer`'s backing `ArrayBuffer` under jsdom.
  return Uint8Array.from(fs.readFileSync(path.join(FIXTURE_DIR, name))).buffer
}

describe.skipIf(SOFFICE === null)('LibreOffice reads the spreadsheets Atlas writes (OFFICE-VERIFY-2)', () => {
  it.each(FIXTURES)('$file round-trips its cell text through both save paths', async (fixture) => {
    const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-sheet-'))
    try {
      const buffer = readFixture(fixture.file)
      const original = path.join(workDir, fixture.file)
      fs.writeFileSync(original, Buffer.from(readFixture(fixture.file)))

      const document = await loadDocument(buffer)

      const originalOut = convert(SOFFICE!, original, CONVERT.calcHtml, workDir)
      expect(originalOut, `LibreOffice could not read the UNTOUCHED fixture ${fixture.file}`).not.toBeNull()
      const before = htmlToText(fs.readFileSync(originalOut!, 'utf8'))
      // A vacuous comparison would otherwise pass: two empty extractions are equal.
      expect(before, `no cell text extracted from ${fixture.file}`).not.toBe('')

      // --- the rebuild branch, which every `.ods` save and every failed passthrough uses
      const rebuilt = path.join(workDir, `rebuilt-${fixture.file}`)
      fs.writeFileSync(rebuilt, Buffer.from(await writeWorkbookBytesWithTables(document, fixture.bookType)))
      const rebuiltOut = convert(SOFFICE!, rebuilt, CONVERT.calcHtml, workDir)
      expect(rebuiltOut, `LibreOffice could not read the REBUILT ${fixture.file}`).not.toBeNull()
      expect(htmlToText(fs.readFileSync(rebuiltOut!, 'utf8')), 'rebuild branch changed the cell text').toBe(before)

      // --- the passthrough branch, which a real `.xlsx` save uses
      const throughOriginal = await writeWorkbookThroughOriginal(buffer, document)
      if (fixture.passthroughEligible) {
        expect(throughOriginal, `the passthrough writer refused an eligible ${fixture.file}`).not.toBeNull()
        const patched = path.join(workDir, `patched-${fixture.file}`)
        fs.writeFileSync(patched, Buffer.from(throughOriginal!))
        const patchedOut = convert(SOFFICE!, patched, CONVERT.calcHtml, workDir)
        expect(patchedOut, `LibreOffice could not read the PASSTHROUGH ${fixture.file}`).not.toBeNull()
        expect(htmlToText(fs.readFileSync(patchedOut!, 'utf8')), 'passthrough branch changed the cell text').toBe(before)
      } else {
        // Recorded as an expectation rather than skipped: if `.ods` ever gains a
        // passthrough writer, this line is what says the branch above needs enabling.
        expect(throughOriginal, `${fixture.file} is not passthrough-eligible`).toBeNull()
      }
    } finally {
      fs.rmSync(workDir, { recursive: true, force: true })
    }
  }, 600_000)

  /**
   * Proves the comparison above can actually FAIL.
   *
   * A round-trip check that passes is only meaningful if it would notice a change,
   * and there are several ways for one of these to quietly become a no-op: an
   * extractor that returns the same string for every input, a filter that silently
   * writes nothing, a convert step whose output nobody reads. Every one of those
   * shows up as a green run. So one cell is deliberately changed, and the check has
   * to catch it — if this ever starts passing by "detecting" nothing, the harness is
   * broken rather than the code being perfect.
   *
   * Both branches are checked, because they could fail to notice independently.
   */
  it('detects a single changed cell through both save paths', async () => {
    const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-lo-sheet-neg-'))
    try {
      const buffer = readFixture('sample.xlsx')
      const original = path.join(workDir, 'sample.xlsx')
      fs.writeFileSync(original, Buffer.from(readFixture('sample.xlsx')))

      const document = await loadDocument(buffer)
      const mutated = setCellValue(document, 0, 0, 0, 'CHANGED-BY-THE-SELF-CHECK')
      expect(mutated, 'setCellValue did not change anything, so this proves nothing').not.toBe(document)

      const originalOut = convert(SOFFICE!, original, CONVERT.calcHtml, workDir)
      const before = htmlToText(fs.readFileSync(originalOut!, 'utf8'))

      const rebuilt = path.join(workDir, 'rebuilt.xlsx')
      fs.writeFileSync(rebuilt, Buffer.from(await writeWorkbookBytesWithTables(mutated, 'xlsx')))
      const rebuiltOut = convert(SOFFICE!, rebuilt, CONVERT.calcHtml, workDir)
      expect(rebuiltOut).not.toBeNull()
      const rebuiltText = htmlToText(fs.readFileSync(rebuiltOut!, 'utf8'))
      expect(rebuiltText, 'the rebuild branch comparison is blind to a changed cell').not.toBe(before)
      expect(rebuiltText).toContain('CHANGED-BY-THE-SELF-CHECK')

      const throughOriginal = await writeWorkbookThroughOriginal(buffer, mutated)
      expect(throughOriginal).not.toBeNull()
      const patched = path.join(workDir, 'patched.xlsx')
      fs.writeFileSync(patched, Buffer.from(throughOriginal!))
      const patchedOut = convert(SOFFICE!, patched, CONVERT.calcHtml, workDir)
      expect(patchedOut).not.toBeNull()
      const patchedText = htmlToText(fs.readFileSync(patchedOut!, 'utf8'))
      expect(patchedText, 'the passthrough branch comparison is blind to a changed cell').not.toBe(before)
      expect(patchedText).toContain('CHANGED-BY-THE-SELF-CHECK')
    } finally {
      fs.rmSync(workDir, { recursive: true, force: true })
    }
  }, 600_000)
})
