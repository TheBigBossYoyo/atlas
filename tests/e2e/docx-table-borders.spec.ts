import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { Document, Packer, Paragraph, TextRun } from 'docx'

/**
 * F2 — Table Properties' "Show borders" checkbox silently discarded on save:
 * the saved .docx contained no <w:tblBorders> at all after ticking it and
 * clicking Apply.
 *
 * Root cause: nothing to do with the command/model/writer layer (which had
 * its own passing unit coverage) — the dialog's "Width (in)" number input
 * had `step={0.1}`. `twipsToInches` rounds a table's width to the nearest
 * 0.01in, and Word's own default single-column table width (9000 twips)
 * rounds to 6.25in, which isn't on the 0.1-step grid from min 0.1 (the two
 * nearest valid values are 6.2 and 6.3). That failed the input's native
 * HTML5 step-mismatch validation, which silently blocks the browser from
 * ever firing the `<form>`'s `submit` event on the "Apply" button click —
 * no exception, no visible error in an automated/hidden window, and because
 * one invalid control blocks the WHOLE form, every field (borders,
 * alignment, width) was discarded, not just borders. Confirmed with the
 * real Electron app: `form.checkValidity()` was false and the width input's
 * own `validationMessage` named the exact conflict, while zero `submit`
 * events reached even a capturing `window` listener.
 *
 * Fixed in TablePropertiesDialog.tsx by widening the width input to
 * `step="any"`. Driven through the real UI here because this only showed up
 * in the real app: the command/writer unit suite exercises `apply-table-
 * props` directly and never goes through this HTML form at all.
 */

const projectRoot = process.cwd()

async function createFixture(): Promise<string> {
  const doc = new Document({
    sections: [
      {
        children: [new Paragraph({ children: [new TextRun('Table borders test')] })],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-table-borders-'))
  const file = path.join(dir, 'borders.docx')
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

async function insertTableAndSelect(page: Page): Promise<void> {
  const line = page.locator('[data-paragraph-path="0"]').first()
  const box = await line.boundingBox()
  if (box === null) throw new Error('paragraph not rendered')
  await page.mouse.click(box.x + 20, box.y + box.height / 2)
  await page.keyboard.press('End')

  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  await page.getByRole('button', { name: 'Insert Table', exact: true }).click()
  await page.getByRole('button', { name: '1 by 1 table' }).click()
  await page.waitForTimeout(300)

  const table = page.locator('.docx-page__table').first()
  await expect(table).toBeVisible()
  const tbox = await table.boundingBox()
  if (tbox === null) throw new Error('table not rendered')
  await page.mouse.click(tbox.x + tbox.width / 2, tbox.y + tbox.height / 2)
  await page.waitForTimeout(200)
}

// After reopening a saved fixture the table already has content, so this
// variant just clicks into the existing table instead of inserting one.
async function selectExistingTable(page: Page): Promise<void> {
  const table = page.locator('.docx-page__table').first()
  await expect(table).toBeVisible()
  const cell = table.locator('td').first()
  await cell.click()
  await page.waitForTimeout(300)
}

async function openTablePropertiesDialog(page: Page): Promise<void> {
  // TableEditPopover (and its "Edit Table" button) lives on the Insert tab,
  // not Home — the app's default active tab on a fresh window. Switching
  // tabs is a no-op if Insert is already active.
  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  await page.getByRole('button', { name: 'Edit Table', exact: true }).click()
  await page.waitForTimeout(150)
  await page.getByRole('button', { name: 'Table Properties…', exact: true }).click()
  await page.waitForTimeout(150)
}

async function saveAndReadDocumentXml(page: Page, fixture: string): Promise<string> {
  const before = fs.statSync(fixture).mtimeMs
  await page.keyboard.press('Control+s')
  await expect.poll(() => fs.statSync(fixture).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
  await page.waitForTimeout(400)

  const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(fixture))
  return zip.file('word/document.xml')!.async('string')
}

test('F2: ticking "Show borders" and applying writes w:tblBorders, and it survives reopening', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    await insertTableAndSelect(page)
    await openTablePropertiesDialog(page)

    const bordersCheckbox = page.getByRole('checkbox', { name: 'Show borders' })
    await expect(bordersCheckbox).toBeVisible()
    if (!(await bordersCheckbox.isChecked())) {
      await bordersCheckbox.check()
    }
    await expect(bordersCheckbox).toBeChecked()

    await page.getByRole('button', { name: 'Apply', exact: true }).click()
    await page.waitForTimeout(300)

    const documentXml = await saveAndReadDocumentXml(page, fixture)
    expect(documentXml).toContain('w:tblBorders')
    const borderMatch = documentXml.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/)
    expect(borderMatch).not.toBeNull()
    expect(borderMatch![0]).toContain('w:val="single"')

    // "Still there after reopening": close the app entirely and read the
    // .docx back from disk fresh (not from any in-memory state this process
    // still holds), then drive the real UI on that reloaded file and confirm
    // the dialog seeds "Show borders" as checked from the file's own content.
    kill(app)
    const reopened = await launch(fixture)
    try {
      await selectExistingTable(reopened.page)
      await openTablePropertiesDialog(reopened.page)
      const reopenedCheckbox = reopened.page.getByRole('checkbox', { name: 'Show borders' })
      await expect(reopenedCheckbox).toBeChecked()
    } finally {
      kill(reopened.app)
    }
  } finally {
    try {
      kill(app)
    } catch {
      // already killed above in the happy path
    }
  }
})

test('F2: unticking "Show borders" and applying removes/nils the borders', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    await insertTableAndSelect(page)
    await openTablePropertiesDialog(page)

    // Ensure a deterministic starting point: turn borders on and Apply
    // first, so unticking afterward is a real observable transition rather
    // than possibly a no-op against an already-off default.
    const bordersCheckbox = page.getByRole('checkbox', { name: 'Show borders' })
    await expect(bordersCheckbox).toBeVisible()
    if (!(await bordersCheckbox.isChecked())) {
      await bordersCheckbox.check()
    }
    await page.getByRole('button', { name: 'Apply', exact: true }).click()
    await page.waitForTimeout(200)

    await openTablePropertiesDialog(page)
    const reopenedCheckbox = page.getByRole('checkbox', { name: 'Show borders' })
    await expect(reopenedCheckbox).toBeChecked()
    await reopenedCheckbox.uncheck()
    await expect(reopenedCheckbox).not.toBeChecked()

    await page.getByRole('button', { name: 'Apply', exact: true }).click()
    await page.waitForTimeout(300)

    const documentXml = await saveAndReadDocumentXml(page, fixture)
    const borderMatch = documentXml.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/)
    expect(borderMatch).not.toBeNull()
    expect(borderMatch![0]).toContain('w:val="none"')
  } finally {
    kill(app)
  }
})
