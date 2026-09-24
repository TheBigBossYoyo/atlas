import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync, execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'

/**
 * SHELL-5 — "New Slide" (and any other structural slide edit that needs a
 * layout) used to be a silent no-op, with no error anywhere, on a .pptx that
 * has no `ppt/slideLayouts/` part at all. Such files are non-conforming
 * (real PowerPoint/LibreOffice/python-pptx always write one), but the repo's
 * own `sample-multislide.pptx` fixture was itself exactly this shape until
 * FIXTURE-1/SHELL-3's generator fix — meaning the bug was reachable from
 * Atlas's own "canonical" sample file. Root cause:
 * `pptxSlideOps.ts`'s `layoutForNewSlide` returns `null` when
 * `pkg.parts.keys()` has no `ppt/slideLayouts/slideLayoutN.xml` entries, and
 * `addSlide` treated that as "do nothing, return the same package/index".
 *
 * Fix: `addSlide` now synthesizes a minimal, spec-clean fallback
 * slideMaster/slideLayout/theme the first time this happens (see
 * `ensureFallbackLayout`), so the action actually works instead of failing
 * silently.
 */

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')
const main = '.slide-deck__main'

test.beforeAll(() => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
})

async function createLayoutlessDeck(): Promise<string> {
  // sample-noid.pptx is the repo's deliberately non-conforming pptx fixture
  // (see generate.mjs's makeNoIdPptxBuffer): no cNvPr ids AND no
  // slideLayouts/slideMaster/theme parts at all — exactly SHELL-5's repro
  // shape, and no longer a coincidence like the other fixtures were before
  // FIXTURE-1.
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-shell5-')), 'deck.pptx')
  fs.copyFileSync(path.join(fixtureDir, 'sample-noid.pptx'), file)
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  await page.waitForSelector(`${main} .slide-edit__layer`, { timeout: 20_000 })
  await page.waitForTimeout(500)
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

test('SHELL-5: "New Slide" works on a .pptx with no slideLayout at all, instead of silently doing nothing', async () => {
  const file = await createLayoutlessDeck()
  const { app, page } = await launch(file)
  try {
    const status = page.getByText(/^Slide \d+ \/ \d+$/).last()
    await expect(status).toHaveText('Slide 1 / 1')

    await page.getByRole('button', { name: 'New slide' }).click()

    // Before the fix, this stayed at "Slide 1 / 1" forever — a silent no-op.
    // (Which slide ends up active isn't SHELL-5's concern; a real second
    // slide existing at all is.)
    await expect(status).toHaveText(/^Slide \d+ \/ 2$/)

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(300)

    const zip = await JSZip.loadAsync(fs.readFileSync(file))
    const names = Object.keys(zip.files)
    expect(names.some((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name))).toBe(true)
    expect(names.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))).toHaveLength(2)
  } finally {
    kill(app)
  }

  const validation = execSync(`node scripts/validate-office-file.mjs "${file}"`, { cwd: projectRoot, encoding: 'utf8' })
  expect(validation).not.toContain('error')

  // Reopen: the new slide (and its synthesized layout chain) survive a real
  // relaunch of the app, not just the in-memory session.
  const reopened = await launch(file)
  try {
    await expect(reopened.page.getByText(/^Slide \d+ \/ \d+$/).last()).toHaveText('Slide 1 / 2')
  } finally {
    kill(reopened.app)
  }
})
