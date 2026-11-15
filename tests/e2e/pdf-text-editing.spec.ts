/**
 * PDF-EDIT-2 — editing a PDF's own text in the real app, on a file a real
 * producer wrote (LibreOffice: subset TrueType fonts, a justified paragraph
 * with a bold span), then reading the saved file back with pdf.js outside
 * the app. The edit must be real: the old words gone from the file, the new
 * ones written as text in the document's own font, the paragraph still
 * justified, and new text addable on empty space.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'
import { pathToFileURL } from 'node:url'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist/legacy/build/pdf.mjs'

const projectRoot = process.cwd()
GlobalWorkerOptions.workerSrc = pathToFileURL(
  path.join(projectRoot, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs'),
).href
const FIXTURE = path.join(projectRoot, 'src', 'viewers', 'pdf', 'edit', '__fixtures__', 'libreoffice-paragraphs.pdf')

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page; errors: string[] }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.waitForSelector('.pdf-viewer__page-wrapper[data-page="1"] .pdf-viewer__text-layer span', { timeout: 20_000 })
  return { app, page, errors }
}

async function savedText(file: string): Promise<string> {
  const task = getDocument({ data: new Uint8Array(fs.readFileSync(file)) })
  try {
    const page = await (await task.promise).getPage(1)
    const content = await page.getTextContent()
    return content.items.map((item) => ('str' in item ? item.str : '')).join('')
  } finally {
    await task.destroy()
  }
}

/** Selects `needle` inside the open paragraph editor. */
async function selectInEditor(page: Page, needle: string): Promise<void> {
  await page.getByRole('textbox', { name: 'Paragraph text' }).evaluate((node, target) => {
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT)
    for (let text = walker.nextNode(); text; text = walker.nextNode()) {
      const at = (text.textContent ?? '').indexOf(target)
      if (at < 0) continue
      const range = document.createRange()
      range.setStart(text, at)
      range.setEnd(text, at + target.length)
      const selection = window.getSelection()!
      selection.removeAllRanges()
      selection.addRange(range)
      return
    }
    throw new Error(`"${target}" is not in the editor`)
  }, needle)
}

test('a paragraph is rewritten in place: old words gone, new words typed, a word made bold, new text added', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-pdf-text-')), 'report.pdf')
  fs.copyFileSync(FIXTURE, file)
  const { app, page, errors } = await launch(file)
  try {
    await page.getByRole('button', { name: 'Edit', exact: true }).click()
    const tool = page.getByRole('radio', { name: /^Edit text/ })
    await expect(tool).toBeEnabled({ timeout: 10_000 })
    await tool.click()

    const word = page.locator('.pdf-viewer__text-layer span', { hasText: 'quick' }).first()
    const box = (await word.boundingBox())!
    await page.mouse.click(box.x + 4, box.y + box.height / 2)
    const editor = page.getByRole('textbox', { name: 'Paragraph text' })
    await expect(editor).toBeFocused()
    await expect(editor).toContainText('The quick brown fox')

    // Typing, spaces included (Space is also the viewer's next-page key).
    await selectInEditor(page, 'quick brown')
    await page.keyboard.type('swift auburn')
    await selectInEditor(page, 'Revenue')
    await page.keyboard.press('Control+b')
    await page.keyboard.press('Control+Enter')
    await expect(editor).toHaveCount(0)

    // New text on empty space, below the table.
    const pageBox = (await page.locator('.pdf-viewer__page-wrapper[data-page="1"]').boundingBox())!
    await page.mouse.click(pageBox.x + 100, pageBox.y + 330)
    await expect(editor).toBeFocused()
    await page.keyboard.type('Added with Atlas')

    // Saved while that editor is still open: the save commits it first.
    await page.keyboard.press('Control+s')
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeDisabled({ timeout: 15_000 })

    const text = await savedText(file)
    expect(text).toContain('The swift auburn fox')
    expect(text).not.toContain('quick brown')
    expect(text).toContain('Added with Atlas')
    // Everything that wasn't touched is still there.
    expect(text).toContain('annual budget')
    expect(text).toContain('A short second paragraph.')
    expect(errors, errors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})

test('an edited paragraph can be re-opened, changed again and undone', async () => {
  const { app, page, errors } = await launch(FIXTURE)
  try {
    await page.getByRole('button', { name: 'Edit', exact: true }).click()
    const tool = page.getByRole('radio', { name: /^Edit text/ })
    await expect(tool).toBeEnabled({ timeout: 10_000 })
    await tool.click()

    const target = page.locator('.pdf-viewer__text-layer span', { hasText: 'short second' }).first()
    const box = (await target.boundingBox())!
    await page.mouse.click(box.x + 4, box.y + box.height / 2)
    const editor = page.getByRole('textbox', { name: 'Paragraph text' })
    await selectInEditor(page, 'short')
    await page.keyboard.type('brief')
    await page.keyboard.press('Control+Enter')
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeEnabled()

    // The page is redrawn from the edited file; the text layer says so.
    await expect(page.locator('.pdf-viewer__text-layer span', { hasText: 'brief' }).first()).toBeVisible({ timeout: 10_000 })

    // Re-open: the editor shows the edit, not the original.
    await page.mouse.click(box.x + 4, box.y + box.height / 2)
    await expect(editor).toContainText('A brief second paragraph.')
    await page.keyboard.press('Escape')

    await page.keyboard.press('Control+z')
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeDisabled()
    expect(errors, errors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})
