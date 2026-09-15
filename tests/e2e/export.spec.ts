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
// @ts-expect-error -- pdfjs-dist ships no types for the legacy Node entry point.
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
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer), disableWorker: true }).promise
  return doc.numPages as number
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
      { timeout: 30_000, intervals: [250, 500, 1000] },
    )
    .toBeGreaterThan(0)
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
    const outDir = await fs.mkdtemp(path.join(os.tmpdir(), 'atlas-export-e2e-'))
    const outPath = path.join(outDir, 'out.pdf')

    const electronApp = await electron.launch({
      args: ['.', path.join(fixtureDir, 'sample-multipage.docx')],
      cwd: projectRoot,
      env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
    })

    try {
      await stubSaveDialog(electronApp, outPath)
      const page = await electronApp.firstWindow()
      await waitForViewer(page, 'docx')
      // eslint-disable-next-line no-console
      console.log('DEBUG docx-page count:', await page.locator('.docx-page').count())
      // eslint-disable-next-line no-console
      console.log('DEBUG docx-page texts:', JSON.stringify(await page.locator('.docx-page').allTextContents()))
      await exportToPdf(page)
      await waitForFile(outPath)

      await expect.poll(() => countPdfPages(outPath), { timeout: 10_000 }).toBe(2)
    } finally {
      await electronApp.close()
      await fs.rm(outDir, { recursive: true, force: true })
    }
  })

  test('a 2-slide PPTX exports as a 2-page PDF, one page per slide', async () => {
    const outDir = await fs.mkdtemp(path.join(os.tmpdir(), 'atlas-export-e2e-'))
    const outPath = path.join(outDir, 'out.pdf')

    const electronApp = await electron.launch({
      args: ['.', path.join(fixtureDir, 'sample-multislide.pptx')],
      cwd: projectRoot,
      env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
    })

    try {
      await stubSaveDialog(electronApp, outPath)
      const page = await electronApp.firstWindow()
      await waitForViewer(page, 'pptx')
      await exportToPdf(page)
      await waitForFile(outPath)

      await expect.poll(() => countPdfPages(outPath), { timeout: 10_000 }).toBe(2)
    } finally {
      await electronApp.close()
      await fs.rm(outDir, { recursive: true, force: true })
    }
  })

  test('a 2-sheet XLSX workbook exports with content from both sheets present', async () => {
    const outDir = await fs.mkdtemp(path.join(os.tmpdir(), 'atlas-export-e2e-'))
    const outPath = path.join(outDir, 'out.pdf')

    const electronApp = await electron.launch({
      args: ['.', path.join(fixtureDir, 'sample-multisheet.xlsx')],
      cwd: projectRoot,
      env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
    })

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

      const buffer = await fs.readFile(outPath)
      const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer), disableWorker: true }).promise
      const texts = await Promise.all(
        Array.from({ length: doc.numPages }, (_, i) =>
          doc
            .getPage(i + 1)
            .then((p: { getTextContent: () => Promise<{ items: Array<{ str?: string }> }> }) => p.getTextContent())
            .then((content: { items: Array<{ str?: string }> }) => content.items.map(item => item.str ?? '').join(' ')),
        ),
      )
      const allText = texts.join('\n')
      expect(allText).toContain('Berlin')
      expect(allText).toContain('Atlas')
    } finally {
      await electronApp.close()
      await fs.rm(outDir, { recursive: true, force: true })
    }
  })
})
