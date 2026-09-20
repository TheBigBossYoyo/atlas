import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { Document, Packer, Paragraph, TextRun } from 'docx'

/**
 * F1 — Electron does not implement `window.prompt` (it throws "prompt() is
 * not supported"), so Insert Hyperlink (toolbar + Ctrl+K), Add Comment, and
 * Reply to Comment previously did nothing at all in the packaged app: no
 * dialog ever appeared, nothing was inserted, and the error landed silently
 * in the console. These drive the real app end to end through the new
 * DocxPromptDialog and, for the hyperlink case, verify the result by
 * unzipping the saved `.docx` after a real save/close/reopen — not by
 * reading the in-memory document model, which unit tests already cover.
 */

const projectRoot = process.cwd()

async function createFixture(): Promise<string> {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun('Atlas prompt dialog fixture.')] }),
          new Paragraph({ children: [new TextRun('A second paragraph to select text in.')] }),
        ],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-prompt-'))
  const file = path.join(dir, 'prompt-dialogs.docx')
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

async function doubleClickWord(page: Page, paragraph: number, word: number): Promise<void> {
  await page.locator(`[data-paragraph-path="${paragraph}"] .docx-run`).nth(word).dblclick()
  await page.waitForTimeout(200)
}

test('Insert Hyperlink (Ctrl+K): no window.prompt, a real dialog appears, applies on Enter, and survives save/close/reopen', async () => {
  const file = await createFixture()
  const { app, page } = await launch(file)
  let reopened: { app: ElectronApplication; page: Page } | null = null
  try {
    // Sanity check: Electron's `window.prompt` really does throw here, same
    // as production — this is what made all three features silent no-ops
    // before this fix, and must stay true (nothing in this batch should
    // ever call it).
    const promptThrows = await page.evaluate(() => {
      try {
        window.prompt('x')
        return false
      } catch {
        return true
      }
    })
    expect(promptThrows).toBe(true)

    await doubleClickWord(page, 0, 1)
    await page.keyboard.press('Control+k')

    const dialog = page.getByRole('dialog', { name: 'Insert hyperlink' })
    await expect(dialog).toBeVisible()
    const urlInput = page.getByLabel('URL', { exact: true })
    await expect(urlInput).toHaveValue('https://')
    await expect(urlInput).toBeFocused()

    await urlInput.fill('https://example.com/atlas-f1')
    await page.keyboard.press('Enter')
    await expect(dialog).not.toBeVisible()

    // Applied into the live document — a real <a class="docx-hyperlink">.
    await expect(page.locator('a.docx-hyperlink')).toBeVisible()

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(400)
  } finally {
    kill(app)
  }

  try {
    // Close → reopen: read back the actual saved bytes, not the in-memory
    // document model the app already had.
    reopened = await launch(file)
    await expect(reopened.page.locator('a.docx-hyperlink')).toBeVisible({ timeout: 20_000 })

    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml).toMatch(/<w:hyperlink[^>]*r:id="[^"]+"/)

    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(relsXml).toContain('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink')
    expect(relsXml).toContain('Target="https://example.com/atlas-f1"')
    expect(relsXml).toContain('TargetMode="External"')

    // The relationship id in the rels part must be the same one the
    // document part's <w:hyperlink> actually references.
    const relIdMatch = /<w:hyperlink[^>]*r:id="([^"]+)"/.exec(documentXml)
    expect(relIdMatch).not.toBeNull()
    expect(relsXml).toContain(`Id="${relIdMatch![1]}"`)
  } finally {
    if (reopened) kill(reopened.app)
  }
})

test('Insert Hyperlink: Cancel and Escape both leave the document untouched', async () => {
  const { app, page } = await launch(await createFixture())
  try {
    await doubleClickWord(page, 0, 1)
    await page.keyboard.press('Control+k')

    const dialog = page.getByRole('dialog', { name: 'Insert hyperlink' })
    await expect(dialog).toBeVisible()
    await page.getByLabel('URL', { exact: true }).fill('https://should-not-be-inserted.example')
    await page.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(dialog).not.toBeVisible()
    await expect(page.locator('a.docx-hyperlink')).toHaveCount(0)

    await page.keyboard.press('Control+k')
    await expect(dialog).toBeVisible()
    await page.getByLabel('URL', { exact: true }).fill('https://should-not-be-inserted.example')
    await page.keyboard.press('Escape')
    await expect(dialog).not.toBeVisible()
    await expect(page.locator('a.docx-hyperlink')).toHaveCount(0)
  } finally {
    kill(app)
  }
})

test('Add Comment and Reply: real dialogs apply on confirm, both persist on save', async () => {
  const file = await createFixture()
  const { app, page } = await launch(file)
  try {
    await doubleClickWord(page, 0, 1)
    await page.getByRole('tab', { name: 'Insert' }).click()
    await page.getByRole('button', { name: 'Comment', exact: true }).click()

    const commentDialog = page.getByRole('dialog', { name: 'Add comment' })
    await expect(commentDialog).toBeVisible()
    const commentInput = commentDialog.getByLabel('Comment', { exact: true })
    await expect(commentInput).toBeFocused()
    await commentInput.fill('Please review this line.')
    await page.keyboard.press('Enter')
    await expect(commentDialog).not.toBeVisible()
    await expect(page.getByText('Please review this line.')).toBeVisible()

    await page.getByRole('button', { name: 'Reply', exact: true }).click()
    const replyDialog = page.getByRole('dialog', { name: 'Reply to comment' })
    await expect(replyDialog).toBeVisible()
    const replyInput = replyDialog.getByLabel('Reply', { exact: true })
    await expect(replyInput).toBeFocused()
    await replyInput.fill('Looks good.')
    await page.keyboard.press('Enter')
    await expect(replyDialog).not.toBeVisible()
    await expect(page.getByText('Looks good.')).toBeVisible()

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(400)

    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
    const commentsPart = Object.keys(zip.files).find((name) => name === 'word/comments.xml')
    expect(commentsPart).toBeDefined()
    const commentsXml = await zip.file(commentsPart!)!.async('string')
    expect(commentsXml).toContain('Please review this line.')
    expect(commentsXml).toContain('Looks good.')
  } finally {
    kill(app)
  }
})
