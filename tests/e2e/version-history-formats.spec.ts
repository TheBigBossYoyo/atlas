/**
 * VERSIONS-2 — version history for the binary formats, driven through the real app.
 *
 * VERSIONS-1 gave every format a version on every SAVE, recorded in main. What it
 * could not do was record unsaved work for anything but markdown: the other
 * formats have to be re-serialised from their in-memory model, and nothing in
 * the shell knew how to ask them. The `registerCapture` half of the viewer
 * contract is that missing piece, and this is the test that it is actually
 * plumbed in — the unit suite can prove the contract works, but not that each
 * viewer is wired to it.
 *
 * Each case edits a document, captures a version WITHOUT saving, and then checks
 * both halves of the claim: a version exists, and the file on disk is untouched.
 * The second half matters as much as the first — a capture that quietly saved
 * would pass a test that only counted versions.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

const projectRoot = process.cwd()
const fixtures = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

/** Copies a fixture somewhere disposable, so a stray write cannot damage the repo's copy. */
function scratchCopy(fixtureName: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-versions-fmt-'))
  const file = path.join(dir, fixtureName)
  fs.copyFileSync(path.join(fixtures, fixtureName), file)
  return file
}

async function launch(file: string, readySelector: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const page = await app.firstWindow()
  await page.waitForSelector(readySelector, { timeout: 30_000 })
  await page.waitForTimeout(800)
  return { app, page }
}

/**
 * Opens the history panel and captures a version from what is on screen.
 *
 * This is the manual half of the same capability the two-minute timer uses — the
 * timer itself is not testable here (nobody can wait out two minutes per format
 * in CI, and faking the clock inside a packaged Electron app reaches past the
 * thing being tested), so this drives the identical `captureDocumentBytes` path
 * through the button instead. `useAutoCapture.test.tsx` covers the timer.
 */
async function captureAVersion(page: Page): Promise<Page> {
  await page.getByRole('button', { name: /Version history/i }).click()
  const panel = page.locator('.version-history')
  await expect(panel).toBeVisible({ timeout: 10_000 })

  // The action is only offered when the active viewer registered a capture, so
  // its presence is itself part of the assertion.
  const captureButton = panel.getByRole('button', { name: /Save a version now/i })
  await expect(captureButton).toBeVisible()
  await captureButton.click()
  await page.waitForTimeout(2500)
  return page
}

async function versionCount(page: Page): Promise<number> {
  return page.locator('.version-history__item').count()
}

// Grid geometry, copied from `spreadsheet-editor.spec.ts` — the canvas grid has
// no DOM cells to address.
const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const COLUMN_WIDTH = 120

test('a DOCX can be captured without being saved', async () => {
  const file = scratchCopy('sample.docx')
  const before = fs.readFileSync(file)
  const { app, page } = await launch(file, '[data-paragraph-path]')
  try {
    await page.locator('[data-paragraph-path="0"] .docx-run').first().click()
    await page.keyboard.type('Captured. ', { delay: 20 })
    await page.waitForTimeout(600)

    await captureAVersion(page)

    expect(await versionCount(page), 'the edited document should have produced a version').toBeGreaterThanOrEqual(1)
    // Nothing was saved: a capture that wrote to disk would be a far worse bug
    // than having no capture at all.
    expect(fs.readFileSync(file).equals(before), 'the file on disk must be untouched').toBe(true)
  } finally {
    kill(app)
  }
})

test('an XLSX can be captured without being saved', async () => {
  const file = scratchCopy('sample-multisheet.xlsx')
  const before = fs.readFileSync(file)
  const { app, page } = await launch(file, '.spreadsheet-viewer__grid canvas')
  try {
    // One cell edit through the grid's own overlay editor, the same route
    // `spreadsheet-editor.spec.ts` uses — the first key opens it, the rest is
    // typed once it has focus.
    const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
    await page.mouse.click(box.x + 48 + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT / 2)
    await page.keyboard.type('C')
    await expect(page.locator('#portal textarea')).toBeFocused()
    await page.keyboard.type('aptured', { delay: 20 })
    await page.keyboard.press('Enter')
    await expect(page.locator('#portal textarea')).toHaveCount(0)
    await page.waitForTimeout(800)

    await captureAVersion(page)

    expect(await versionCount(page)).toBeGreaterThanOrEqual(1)
    expect(fs.readFileSync(file).equals(before), 'the file on disk must be untouched').toBe(true)
  } finally {
    kill(app)
  }
})

test('a PPTX can be captured without being saved', async () => {
  const file = scratchCopy('sample-multislide.pptx')
  const before = fs.readFileSync(file)
  const { app, page } = await launch(file, '.slide-deck__main .slide-edit__layer')
  try {
    // The deck's package is captured whether or not an edit landed; the version
    // being recorded at all is what this proves, and the unsaved-file check is
    // what makes it meaningful.
    await captureAVersion(page)

    expect(await versionCount(page)).toBeGreaterThanOrEqual(1)
    expect(fs.readFileSync(file).equals(before), 'the file on disk must be untouched').toBe(true)
  } finally {
    kill(app)
  }
})

test('a source file keeps its CRLF line endings through a capture and restore', async () => {
  // The fidelity rule behind `encodeTextBytes`: a version is restored by writing
  // its bytes verbatim, so a capture that produced plain UTF-8 LF would convert
  // a CRLF file the moment it was restored. Nothing in the version list would
  // show it, which is exactly why it is worth a real end-to-end check.
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-versions-crlf-'))
  const file = path.join(dir, 'script.ts')
  fs.writeFileSync(file, 'const first = 1\r\nconst second = 2\r\n', 'utf8')

  const { app, page } = await launch(file, '.cm-content')
  try {
    await page.locator('.cm-content').click()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('const third = 3', { delay: 20 })
    await page.waitForTimeout(600)

    await captureAVersion(page)
    expect(await versionCount(page)).toBeGreaterThanOrEqual(1)

    // Save, so the captured version can be restored (restoring is refused over
    // unsaved changes), then put the capture back and read the bytes.
    await page.keyboard.press('Escape')
    await page.keyboard.press('Control+s')
    await page.waitForTimeout(2000)

    await page.getByRole('button', { name: /Version history/i }).click()
    const panel = page.locator('.version-history')
    await expect(panel).toBeVisible({ timeout: 10_000 })
    const items = panel.locator('.version-history__item')
    const count = await items.count()
    await items.nth(count - 1).getByRole('button', { name: /Restore/i }).click()
    await page.waitForTimeout(2500)

    const onDisk = fs.readFileSync(file, 'utf8')
    expect(onDisk, `file after restore: ${JSON.stringify(onDisk)}`).toContain('\r\n')
    expect(onDisk).not.toMatch(/[^\r]\n/)
  } finally {
    kill(app)
  }
})
