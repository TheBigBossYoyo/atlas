import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { Document, Packer, Paragraph, TextRun } from 'docx'

/**
 * DOCX-16 — the toolbar's Insert Table used to render the new table at zero
 * width (unclickable), and typing right after inserting it silently
 * discarded every keystroke: the caret never actually reached the model's
 * first-cell position because nothing under a `<table>` is addressable in
 * the rendered DOM yet (a separate, sibling-owned rendering gap — see
 * `commands.ts`'s `applyInsertTable` doc comment), so the app's DOM-
 * selection sync silently failed and left focus stranded outside the
 * editor. Both driven with the real Electron app, not inferred from the
 * model layer alone — the model-level bug (zero-width `buildEmptyTable`
 * geometry) has its own unit coverage in `commands.test.ts`.
 */

const projectRoot = process.cwd()

async function createFixture(): Promise<string> {
  const doc = new Document({
    sections: [
      {
        children: [new Paragraph({ children: [new TextRun('Insert table test')] })],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-insert-table-'))
  const file = path.join(dir, 'insert-table.docx')
  fs.writeFileSync(file, await Packer.toBuffer(doc))
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const page = await app.firstWindow()
  await page.waitForSelector('[data-paragraph-path]', { timeout: 20_000 })
  await page.waitForTimeout(800)
  return { app, page }
}

function kill(app: ElectronApplication): void {
  // The unsaved-changes close guard would block app.close(); tear the whole
  // process tree down instead so the next test can take the single-instance lock.
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

const surfaceText = (page: Page): Promise<string> =>
  page.evaluate(() => document.querySelector('.docx-viewer__surface')?.textContent ?? '')

test('DOCX-16: an inserted table renders at a real width, is clickable, and typed text is never silently discarded', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    const line = page.locator('[data-paragraph-path="0"]').first()
    const box = await line.boundingBox()
    if (box === null) throw new Error('paragraph not rendered')
    await page.mouse.click(box.x + 20, box.y + box.height / 2)
    await page.keyboard.press('End')

    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page.getByRole('button', { name: 'Insert Table', exact: true }).click()
    await page.getByRole('button', { name: '1 by 1 table' }).click()
    await page.waitForTimeout(300)

    // The table must render at a real, non-zero, clickable width — not the
    // 0px box `buildEmptyTable`'s missing tcW/tblW used to produce.
    const table = page.locator('.docx-page__table').first()
    await expect(table).toBeVisible()
    const tableBox = await table.boundingBox()
    expect(tableBox).not.toBeNull()
    expect(tableBox!.width).toBeGreaterThan(0)
    // Playwright's actionability check refuses to click a zero-area element
    // without `force: true` — a real (non-forced) click proves it is
    // genuinely hittable, not just non-zero in the accessibility tree.
    await expect(table).toBeInViewport()
    await table.click({ trial: true })

    // Type immediately, exactly as the toolbar leaves the user positioned —
    // no extra manual click to "rescue" focus. This is the live work-loss
    // scenario: before the fix, every one of these keystrokes vanished.
    await page.keyboard.type('AfterTable', { delay: 30 })
    await expect.poll(() => surfaceText(page)).toContain('AfterTable')

    // And it landed somewhere real and inspectable, not merely "somewhere
    // in the DOM": the addressable paragraph the insert-table command
    // leaves right after the table.
    const afterTableText = await page.evaluate(() =>
      Array.from(document.querySelectorAll('[data-paragraph-path="2"]')).map((el) => el.textContent ?? '').join(''),
    )
    expect(afterTableText).toContain('AfterTable')

    // Save, close, reopen: the typed text must survive on disk.
    const before = fs.statSync(fixture).mtimeMs
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.statSync(fixture).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(400)

    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(fixture))
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml.replace(/<[^>]+>/g, '')).toContain('AfterTable')
  } finally {
    kill(app)
  }
})
