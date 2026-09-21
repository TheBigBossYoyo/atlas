/**
 * X1 — end-to-end coverage for the real per-format PDF export, exercising
 * the ONE part of the pipeline vitest's jsdom environment cannot: the actual
 * main-process `export:printToPdf` IPC (a real hidden `BrowserWindow` with
 * `javascript: false`, `webContents.printToPDF`, a real temp-file
 * write/cleanup — see `electron/lib/printToPdf.cjs`). Every unit test in
 * `src/utils/export/__tests__/` mocks `window.electronAPI.printToPdf`
 * entirely; this proves the real thing produces a valid, correctly-paginated
 * PDF.
 *
 * The native save dialog (`dialog.showSaveDialog`) is stubbed via
 * `electronApp.evaluate` to resolve to a fixed temp path instead of showing
 * a real OS dialog Playwright can't drive — this only replaces the DIALOG,
 * not any export/print/write logic downstream of it.
 */
import path from 'node:path'
import fs from 'node:fs/promises'
import os from 'node:os'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs'

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

test.beforeAll(() => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
})

async function countPdfPages(filePath: string): Promise<number> {
  const buffer = await fs.readFile(filePath)
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise
  return doc.numPages
}

/**
 * TEST-3 (phase4/b4-corpus-tests) — the DOCX/PPTX cases below used to assert
 * only `countPdfPages(...) === 2`, which a bug that duplicated page 1 onto
 * page 2 (or rendered page 2 blank) would still satisfy. Returns each page's
 * own extracted text (index 0 = page 1) so callers can assert real,
 * per-page content — the same technique the XLSX case in this file already
 * uses for its own two-sheet assertion.
 */
async function getPdfPageTexts(filePath: string): Promise<ReadonlyArray<string>> {
  const buffer = await fs.readFile(filePath)
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise
  return Promise.all(
    Array.from({ length: doc.numPages }, (_, i) =>
      doc
        .getPage(i + 1)
        .then(p => p.getTextContent())
        .then(content => content.items.map(item => ('str' in item ? item.str : '')).join(' ')),
    ),
  )
}

/**
 * pdf.js's `getTextContent` splits each printed page's text into one item
 * per positioned text run, which Chromium's print pipeline can break in the
 * middle of a word (observed: printing "fixture" produced separate "fi" and
 * "xture" items — a ligature/kerning artifact of the print rasterizer, not
 * of Atlas's own DOCX/PPTX renderer). `getPdfPageTexts` joins items with a
 * space, which is right for word boundaries but wrong for a mid-word split
 * like that. Stripping ALL whitespace before a `toContain`/`not.toContain`
 * check sidesteps both directions of the ambiguity — real word boundaries
 * and spurious mid-word ones alike — while still telling two genuinely
 * different phrases (e.g. this fixture's two distinct page/slide texts)
 * apart.
 */
function collapseWhitespace(text: string): string {
  return text.replace(/\s+/g, '')
}

/** Waits for a nonzero-size file to appear at `filePath` — `atomicWriteFile`
 * renames a temp file into place, so any observed nonzero size is already
 * the finished, final file, never a partial write. */
async function waitForFile(filePath: string): Promise<void> {
  await expect
    .poll(
      async () => {
        try {
          return (await fs.stat(filePath)).size
        } catch {
          return 0
        }
      },
      // Rendering a document to PDF through Chromium's print pipeline is slow
      // on a loaded CI runner — this timed out once at 30 s there.
      { timeout: 60_000, intervals: [250, 500, 1000] },
    )
    .toBeGreaterThan(0)
}

/**
 * CI investigation (2026-09-21, run 35552076678) — `waitForFile` failed with
 * its own 60 s predicate timeout (not the outer per-test clock), meaning the
 * exported PDF genuinely never reached non-zero size. `export:printToPdf`'s
 * only failure signal was `logMainEvent`, which writes to a rotating file on
 * disk (`electron/lib/crashLog.cjs`) that nothing in this harness ever reads
 * — so a real main-process error (e.g. `PrintToPdfError` from
 * `printToPdf.cjs`'s own 30 s `did-finish-load` timeout, well inside
 * `waitForFile`'s 60 s budget) would leave no trace in the CI log at all,
 * indistinguishable from a genuinely slow render. Piping the Electron
 * process's own stdout/stderr into the test's output turns that silence into
 * a visible `[main stderr]` line the next time this happens, without
 * changing any product behaviour. Paired with the `console.error` added
 * alongside `logMainEvent` in the `export:printToPdf` handler (main.cjs).
 */
function logMainProcessOutput(electronApp: ElectronApplication): void {
  electronApp.process().stdout?.on('data', d => console.log(`[main stdout] ${d.toString()}`))
  electronApp.process().stderr?.on('data', d => console.log(`[main stderr] ${d.toString()}`))
}

async function stubSaveDialog(electronApp: ElectronApplication, targetPath: string): Promise<void> {
  await electronApp.evaluate(({ dialog }, targetPathArg) => {
    dialog.showSaveDialog = (() =>
      Promise.resolve({ canceled: false, filePath: targetPathArg })) as typeof dialog.showSaveDialog
  }, targetPath)
}

async function waitForViewer(page: Page, format: string): Promise<void> {
  // DOCX in particular can take a while to become ready (font preloading +
  // pagination of the whole document before the first page ever mounts).
  await expect(page.locator(`[data-viewer="${format}"]`)).toHaveCount(1, { timeout: 30_000 })
}

async function exportToPdf(page: Page): Promise<void> {
  await page.getByRole('button', { name: 'Export' }).click()
  await page.getByRole('button', { name: 'Export to PDF' }).click()
}

test.describe('Export to PDF captures the FULL document (X1 / SLD-01 / UX-02 / RUN-07 / SHELL-13)', () => {
  test('a 2-page DOCX (explicit page break) exports as a real 2-page PDF', async () => {
    // CI flake investigation — this test's own steps already budget more
    // time than the global `timeout: 60_000` (playwright.config.ts) can
    // ever honor: `waitForViewer` alone allows up to 30 s (DOCX in
    // particular — font preloading + full-document pagination before the
    // first page mounts), then `waitForFile` allows up to another 60 s for
    // Chromium's print pipeline on a loaded CI runner, on top of the
    // Electron launch and the export click themselves. Those two budgets
    // were each raised independently (see their own comments) without ever
    // touching the *outer* per-test timeout that both of them share — so in
    // the one CI run that actually needed close to the advertised 30 s for
    // the viewer, `waitForFile` never got anywhere near its own advertised
    // 60 s before the outer timeout cut the whole test off first (observed:
    // run 35523638613, "Test timeout of 60000ms exceeded" inside
    // `waitForFile`, with no export error logged — the write simply never
    // got the time its own budget promised). Locally this whole test
    // completes in ~7 s; 120 s gives the documented per-step budgets room to
    // actually apply, without turning a genuine failure into a silent pass.
    test.setTimeout(120_000)

    const outDir = await fs.mkdtemp(path.join(os.tmpdir(), 'atlas-export-e2e-'))
    const outPath = path.join(outDir, 'out.pdf')

    const electronApp = await electron.launch({
      args: ['.', path.join(fixtureDir, 'sample-multipage.docx')],
      cwd: projectRoot,
      env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
    })
    logMainProcessOutput(electronApp)

    try {
      await stubSaveDialog(electronApp, outPath)
      const page = await electronApp.firstWindow()
      await waitForViewer(page, 'docx')
      await exportToPdf(page)
      await waitForFile(outPath)

      await expect.poll(() => countPdfPages(outPath), { timeout: 10_000 }).toBe(2)

      // Real per-page content (TEST-3): each page must carry its OWN fixture
      // text (see tests/e2e/fixtures/generate.mjs's `makeMultiPageDocxBuffer`
      // — page one's run text, then an explicit PageBreak, then page two's
      // paragraph), not just any two pages. This is what actually catches a
      // bug that duplicated page one onto page two, or rendered page two
      // blank — both would still satisfy `numPages === 2` above.
      const [page1Text, page2Text] = (await getPdfPageTexts(outPath)).map(collapseWhitespace)
      expect(page1Text).toContain(collapseWhitespace('Atlas multipage DOCX fixture'))
      expect(page1Text).toContain(collapseWhitespace('page one'))
      expect(page1Text).not.toContain(collapseWhitespace('Page two content'))

      expect(page2Text).toContain(collapseWhitespace('Page two content'))
      expect(page2Text).not.toContain(collapseWhitespace('Atlas multipage DOCX fixture'))
    } finally {
      await electronApp.close()
      await fs.rm(outDir, { recursive: true, force: true })
    }
  })

  test('a 2-slide PPTX exports as a 2-page PDF, one page per slide', async () => {
    // Same nested-budget fix as the DOCX test above — see its comment.
    test.setTimeout(120_000)

    const outDir = await fs.mkdtemp(path.join(os.tmpdir(), 'atlas-export-e2e-'))
    const outPath = path.join(outDir, 'out.pdf')

    const electronApp = await electron.launch({
      args: ['.', path.join(fixtureDir, 'sample-multislide.pptx')],
      cwd: projectRoot,
      env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
    })
    logMainProcessOutput(electronApp)

    try {
      await stubSaveDialog(electronApp, outPath)
      const page = await electronApp.firstWindow()
      await waitForViewer(page, 'pptx')
      await exportToPdf(page)
      await waitForFile(outPath)

      await expect.poll(() => countPdfPages(outPath), { timeout: 10_000 }).toBe(2)

      // Real per-page content (TEST-3): each page must carry its OWN slide's
      // text (see tests/e2e/fixtures/generate.mjs's
      // `makeMultiSlidePptxBuffer` — slide1.xml vs slide2.xml), not just any
      // two pages. This is what actually catches the SLD-01/UX-02 class of
      // bug this whole test exists for: the old exporter rasterized
      // whichever single slide the virtualized deck viewport happened to
      // have mounted, which could produce 2 pages that are really the same
      // slide twice.
      const [page1Text, page2Text] = (await getPdfPageTexts(outPath)).map(collapseWhitespace)
      expect(page1Text).toContain(collapseWhitespace('Atlas multislide fixture'))
      expect(page1Text).toContain(collapseWhitespace('slide one'))
      expect(page1Text).not.toContain(collapseWhitespace('Slide two content'))

      expect(page2Text).toContain(collapseWhitespace('Slide two content'))
      expect(page2Text).not.toContain(collapseWhitespace('Atlas multislide fixture'))
    } finally {
      await electronApp.close()
      await fs.rm(outDir, { recursive: true, force: true })
    }
  })

  test('a 2-sheet XLSX workbook exports with content from both sheets present', async () => {
    // Same nested-budget fix as the DOCX test above — see its comment.
    test.setTimeout(120_000)

    const outDir = await fs.mkdtemp(path.join(os.tmpdir(), 'atlas-export-e2e-'))
    const outPath = path.join(outDir, 'out.pdf')

    const electronApp = await electron.launch({
      args: ['.', path.join(fixtureDir, 'sample-multisheet.xlsx')],
      cwd: projectRoot,
      env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
    })
    logMainProcessOutput(electronApp)

    try {
      await stubSaveDialog(electronApp, outPath)
      const page = await electronApp.firstWindow()
      await waitForViewer(page, 'xlsx')
      await exportToPdf(page)
      await waitForFile(outPath)

      // One page-group per sheet — two sheets means at least two pages
      // (never asserted as exactly 2: a sheet's own content could in
      // principle overflow one page, which would only ever ADD pages).
      const pageCount = await countPdfPages(outPath)
      expect(pageCount).toBeGreaterThanOrEqual(2)

      const texts = await getPdfPageTexts(outPath)
      const allText = texts.join('\n')
      expect(allText).toContain('Berlin')
      expect(allText).toContain('Atlas')
    } finally {
      await electronApp.close()
      await fs.rm(outDir, { recursive: true, force: true })
    }
  })
})
