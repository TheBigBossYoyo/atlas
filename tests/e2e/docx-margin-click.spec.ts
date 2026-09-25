import path from 'node:path'
import fs from 'node:fs'
import os from 'node:os'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * DOCX-2 — clicking inside a paragraph's box but outside its rendered text
 * (a click whose native event target is an ANCESTOR of the editor surface,
 * not a descendant of it — e.g. once the Comments panel narrows the page
 * column enough that part of a line's own layout position falls outside the
 * surface's clipped/scrollable box) left focus and the native `Selection`
 * looking perfectly normal while Atlas's own `range` state was never
 * updated, since `handleSurfaceMouseDown`'s `onMouseDown` only fires for a
 * click whose target is the surface or a DESCENDANT of it. Every keystroke
 * after such a click then silently no-opped behind `Input.ts`'s
 * `range === null` guard — no error, no visual sign anything was wrong.
 *
 * Fixed with a native `mousedown` listener on the outer `.docx-viewer`
 * container in `DocxViewer.tsx` (see its own doc comment), which resolves
 * the click via the same geometry-based `positionFromClientPoint` a normal
 * click already uses — so it works regardless of where the native event's
 * target landed. See atlas-night/docx-report.md's DOCX-2 section for the
 * original repro.
 */

const projectRoot = process.cwd()

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const page = await app.firstWindow()
  await page.waitForSelector('[data-paragraph-path]', { timeout: 20_000 })
  await page.waitForTimeout(1000)
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function saveAndReadDocumentXml(page: Page, file: string): Promise<string> {
  const before = fs.statSync(file).mtimeMs
  await page.getByRole('button', { name: 'Save', exact: true }).click()
  await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
  await page.waitForTimeout(400)
  const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
  return zip.file('word/document.xml')!.async('string')
}

test('DOCX-2: clicking the page margin next to a paragraph (with the Comments panel open) still places a working caret', async () => {
  const src = path.join(projectRoot, 'src/docx/__fixtures__/corpus/comments-with-reply.docx')
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-margin-click-'))
  const file = path.join(dir, 'comments-with-reply.docx')
  fs.copyFileSync(src, file)

  const { app, page } = await launch(file)
  try {
    // Opening a document with existing comments auto-opens the Comments
    // panel, narrowing the page column — the real layout condition the
    // original bug report reproduced against.
    const p0 = page.locator('[data-paragraph-path="0"]').first()
    const box = await p0.boundingBox()
    if (box === null) throw new Error('paragraph 0 has no bounding box')

    // 20px in from paragraph 0's own reported left edge: inside its logical
    // box, but the exact coordinate the original report found unclickable
    // once the Comments panel narrowed the page.
    await page.mouse.click(box.x + 20, box.y + box.height / 2)
    await page.waitForTimeout(200)

    await page.keyboard.type('MARGINCLICK', { delay: 20 })
    await page.waitForTimeout(300)

    const text = await page.evaluate(
      (p) => Array.from(document.querySelectorAll(`[data-paragraph-path="${p}"]`)).map((el) => el.textContent ?? '').join(''),
      0,
    )
    // The pre-fix bug: this stayed exactly "Fixture: comments-with-reply."
    // (the keystroke discarded silently, no error, caret looked normal).
    expect(text).toContain('MARGINCLICK')

    const xml = await saveAndReadDocumentXml(page, file)
    expect(xml).toContain('MARGINCLICK')
  } finally {
    kill(app)
  }
})
