import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from 'docx'

/**
 * DOCX-1 — ArrowUp/ArrowDown/PageUp/PageDown silently teleported the
 * insertion point to the very start/end of the document instead of moving
 * one visual line, because Atlas's own range model didn't implement
 * vertical movement at all (`Input.ts` returned `null` and let the key fall
 * through to the browser's native contentEditable handling, which this
 * layout breaks since every run is `display: inline-block`). One keypress
 * was enough: whatever the user typed next landed nowhere near the visible
 * caret, with no error, dialog or visual indication anything was wrong. See
 * `atlas-night/docx-report.md`'s DOCX-1 section for the original repro.
 *
 * These specs drive the real Electron app and verify the SAVED bytes
 * (`word/document.xml`), not just on-screen text, matching how the bug was
 * originally found.
 */

const projectRoot = process.cwd()

async function createFixture(): Promise<string> {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Heading line zero')] }),
          new Paragraph({ children: [new TextRun('Paragraph one alpha bravo charlie.')] }),
          new Paragraph({ children: [new TextRun('Paragraph two delta echo foxtrot.')] }),
          new Paragraph({ children: [new TextRun('Paragraph three golf hotel india.')] }),
          new Paragraph({ children: [new TextRun('Paragraph four juliet kilo lima.')] }),
          new Paragraph({ children: [new TextRun('Paragraph five mike november oscar.')] }),
        ],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-vertical-caret-'))
  const file = path.join(dir, 'vertical-caret.docx')
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

const paragraphText = (page: Page, index: number): Promise<string> =>
  page.evaluate(
    (p) => Array.from(document.querySelectorAll(`[data-paragraph-path="${p}"]`)).map((el) => el.textContent ?? '').join(''),
    index,
  )

async function saveAndReadDocumentXml(page: Page, file: string): Promise<string> {
  const before = fs.statSync(file).mtimeMs
  await page.getByRole('button', { name: 'Save', exact: true }).click()
  await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
  await page.waitForTimeout(400)
  const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
  return zip.file('word/document.xml')!.async('string')
}

// Each test saves into its own fixture file (created fresh per test rather
// than shared via `beforeAll`) — a shared file would carry one test's saved
// edits into the next test's paragraph text and produce false positives/
// negatives unrelated to DOCX-1 itself.

test('DOCX-1: a single ArrowDown moves the caret one paragraph down, not to the end of the document', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    const p1 = page.locator('[data-paragraph-path="1"]').first()
    const box = await p1.boundingBox()
    if (box === null) throw new Error('paragraph 1 has no bounding box')
    await page.mouse.click(box.x + 10, box.y + box.height / 2)
    await page.waitForTimeout(150)

    await page.keyboard.press('ArrowDown')
    await page.waitForTimeout(150)
    await page.keyboard.type('CARETMARK', { delay: 20 })
    await page.waitForTimeout(200)

    // On-screen: the marker must land in paragraph 2, not paragraph 5 (the
    // pre-fix bug always appended to the very last paragraph).
    expect(await paragraphText(page, 2)).toContain('CARETMARK')
    expect(await paragraphText(page, 5)).not.toContain('CARETMARK')

    const xml = await saveAndReadDocumentXml(page, fixture)
    expect(xml).toContain('CARETMARK')
    // Saved bytes confirm it too, not just the live DOM.
    expect(xml.indexOf('CARETMARK')).toBeLessThan(xml.indexOf('Paragraph five'))
  } finally {
    kill(app)
  }
})

test('DOCX-1: ArrowUp from paragraph 3 moves the caret up one line, not to the start of the document', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    const p3 = page.locator('[data-paragraph-path="3"]').first()
    const box = await p3.boundingBox()
    if (box === null) throw new Error('paragraph 3 has no bounding box')
    await page.mouse.click(box.x + 10, box.y + box.height / 2)
    await page.waitForTimeout(150)

    await page.keyboard.press('ArrowUp')
    await page.waitForTimeout(150)
    await page.keyboard.type('UPMARK', { delay: 20 })
    await page.waitForTimeout(200)

    // Pre-fix, this landed in front of "Heading line zero" (paragraph 0).
    expect(await paragraphText(page, 2)).toContain('UPMARK')
    expect(await paragraphText(page, 0)).not.toContain('UPMARK')

    const xml = await saveAndReadDocumentXml(page, fixture)
    expect(xml).toContain('UPMARK')
  } finally {
    kill(app)
  }
})

test('DOCX-1: ArrowUp on the first line stays on that line instead of erroring or moving further', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    const p0 = page.locator('[data-paragraph-path="0"]').first()
    const box = await p0.boundingBox()
    if (box === null) throw new Error('paragraph 0 has no bounding box')
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2)
    await page.waitForTimeout(150)

    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(err.message || String(err)))

    await page.keyboard.press('ArrowUp')
    await page.waitForTimeout(150)
    await page.keyboard.type('TOPMARK', { delay: 20 })
    await page.waitForTimeout(200)

    expect(pageErrors).toEqual([])
    expect(await paragraphText(page, 0)).toContain('TOPMARK')
  } finally {
    kill(app)
  }
})

test('DOCX-1: ArrowDown on the last line stays on that line instead of erroring or moving further', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    const p5 = page.locator('[data-paragraph-path="5"]').first()
    const box = await p5.boundingBox()
    if (box === null) throw new Error('paragraph 5 has no bounding box')
    await page.mouse.click(box.x + box.width - 5, box.y + box.height / 2)
    await page.waitForTimeout(150)

    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(err.message || String(err)))

    await page.keyboard.press('ArrowDown')
    await page.waitForTimeout(150)
    await page.keyboard.type('BOTTOMMARK', { delay: 20 })
    await page.waitForTimeout(200)

    expect(pageErrors).toEqual([])
    expect(await paragraphText(page, 5)).toContain('BOTTOMMARK')
  } finally {
    kill(app)
  }
})

test('DOCX-1: Shift+ArrowDown extends the selection instead of collapsing it', async () => {
  const fixture = await createFixture()
  const { app, page } = await launch(fixture)
  try {
    const p1 = page.locator('[data-paragraph-path="1"]').first()
    const box = await p1.boundingBox()
    if (box === null) throw new Error('paragraph 1 has no bounding box')
    // Click right before "alpha" (after "Paragraph one ").
    await page.mouse.click(box.x + 10, box.y + box.height / 2)
    await page.waitForTimeout(150)
    await page.keyboard.press('Home')
    await page.waitForTimeout(100)

    await page.keyboard.down('Shift')
    await page.keyboard.press('ArrowDown')
    await page.keyboard.up('Shift')
    await page.waitForTimeout(150)

    const selectionText = await page.evaluate(() => window.getSelection()?.toString() ?? '')
    // A real selection spanning from paragraph 1's start into paragraph 2,
    // not a collapsed caret (empty string) and not the selection jumping to
    // the document's end.
    expect(selectionText.length).toBeGreaterThan(0)
    expect(selectionText).toContain('Paragraph one')
  } finally {
    kill(app)
  }
})
