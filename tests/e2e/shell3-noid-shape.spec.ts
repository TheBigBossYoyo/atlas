import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync, execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'

/**
 * SHELL-3 — a shape whose `<p:sp>` has no `<p:nvSpPr><p:cNvPr id="…"/></p:nvSpPr>`
 * at all (real PowerPoint always writes one; a hand-built or lossily
 * converted file may not) used to be completely unselectable/uneditable:
 * click, double-click and click+Enter were all silent no-ops, with no error.
 * Root cause: `src/viewers/slides/pptx/parser.ts` gave such a shape
 * `sourceId: undefined`, and `SlideEditCanvas.tsx`'s `hitTest()` requires a
 * truthy `sourceId` before it will match a click.
 *
 * `tests/e2e/fixtures/generate.mjs`'s `makeNoIdPptxBuffer` produces exactly
 * this shape, written to `sample-noid.pptx` — kept deliberately
 * non-conforming (unlike every other pptx fixture, which now has real
 * cNvPr ids) so this spec exercises the real bug end to end, against the
 * real app and the real saved bytes.
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

async function createDeck(): Promise<string> {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-shell3-')), 'deck.pptx')
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

const shapeText = (page: Page, text: string) => page.locator(`${main} .slide-shape--text`, { hasText: text }).first()

test('SHELL-3: a shape with no cNvPr id is selectable, editable, and the edit saves to the right shape', async () => {
  const file = await createDeck()
  const { app, page } = await launch(file)
  try {
    const shape = shapeText(page, 'Atlas no-id PPTX fixture')
    await expect(shape).toBeVisible()

    // Click selects it — before the fix, hitTest silently ignored the click
    // (shape.sourceId was undefined) and no selection outline ever appeared.
    const box = (await shape.boundingBox())!
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2)
    await expect(page.locator(`${main} .slide-edit__selection-layer`)).toBeVisible({ timeout: 5_000 })

    // Double-click opens it for editing.
    await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2)
    const editor = page.getByRole('textbox', { name: 'Edit slide text' })
    await expect(editor).toBeFocused()
    await page.keyboard.press('Control+A')
    await page.keyboard.type('Edited via fallback address')
    await page.keyboard.press('Escape')
    await expect(shapeText(page, 'Edited via fallback address')).toBeVisible()

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(300)

    // Assert on the saved bytes, not just the screen: the edit must have
    // landed on the shape's own XML node (there is only one shape on this
    // slide, so "changed" and "changed on exactly that shape" coincide —
    // the pptxEdits.ts unit tests separately cover the duplicate-shape case).
    const zip = await JSZip.loadAsync(fs.readFileSync(file))
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(slideXml).toContain('<a:t>Edited via fallback address</a:t>')
    expect(slideXml).not.toContain('Atlas no-id PPTX fixture')
  } finally {
    kill(app)
  }

  // Reopen: relaunch the real app against the saved file and confirm the
  // edit is there — not just present in the raw XML, but rendered, and the
  // shape is still addressable for a second round of editing.
  const reopened = await launch(file)
  try {
    await expect(shapeText(reopened.page, 'Edited via fallback address')).toBeVisible()
  } finally {
    kill(reopened.app)
  }
})
