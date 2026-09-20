/**
 * D23-PERF-2 — the DOCX page-virtualization surface at scale: a real 36-page
 * document with images, driven through scroll/type/Find/undo-redo/save/
 * reopen. No existing e2e fixture went past 2 pages, so virtualization's
 * mount/unmount churn (`PageStack.tsx`) was previously only unit-tested
 * against synthetic offsets, never against a real editor session.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync, execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

let largeDocxPath: string

test.beforeAll(() => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-scale-'))
  largeDocxPath = path.join(dir, 'large36.docx')
  execFileSync(process.execPath, [path.join(projectRoot, 'tests', 'e2e', 'fixtures', 'generateLargeDocx.mjs'), largeDocxPath, '36'], {
    stdio: 'inherit',
  })
})

test('36-page docx with images: scroll to end, type, find, undo/redo, save+reopen preserves content', async () => {
  const app = await electron.launch({
    args: ['.', largeDocxPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))

    await page.waitForSelector('[data-paragraph-path]', { state: 'attached', timeout: 20_000 })
    await page.waitForTimeout(1000)

    // Scroll fast to the end of the document.
    for (let i = 0; i < 10; i++) {
      await page.mouse.wheel(0, 4000)
      await page.waitForTimeout(50)
    }
    await page.waitForTimeout(500)

    // Find the end-of-document marker to confirm the whole doc actually loaded and virtualization didn't drop content.
    await page.locator('[data-paragraph-path]').first().click({ position: { x: 5, y: 5 } })
    await page.keyboard.press('Control+f')
    const findInput = page.locator('.docx-find__input').first()
    await expect(findInput).toBeVisible({ timeout: 5000 })
    await findInput.fill('END-OF-DOCUMENT-MARKER-XYZZY')
    await page.waitForTimeout(500)
    const statusText = await page.locator('.docx-find__status').innerText()
    console.log('find status text:', statusText)
    // A real, non-zero match count proves the end-of-document text is
    // actually present in the live model, not dropped by virtualization
    // (Find must search the full document, not just mounted pages).
    expect(statusText.length).toBeGreaterThan(0)
    expect(statusText.toLowerCase()).not.toMatch(/no match|aucune correspondance|^0 /)
    await page.keyboard.press('Escape')
    await page.waitForTimeout(300)

    // Already scrolled to the end by the wheel loop above (real user-style
    // scrolling — see the "reopen" test below for why a programmatic
    // `scrollTop =` assignment is not a reliable substitute here). Type near
    // the end marker.
    const lastParagraph = page.locator('[data-paragraph-path]').last()
    await lastParagraph.click({ position: { x: 5, y: 5 } })
    await page.keyboard.press('End')
    await page.keyboard.type(' APPENDED-TAIL-TEXT')
    await page.waitForTimeout(300)

    const lastText = await lastParagraph.textContent()
    console.log('last paragraph text after typing:', lastText)
    expect(lastText).toContain('APPENDED-TAIL-TEXT')

    // Undo the typing, then redo it.
    await page.keyboard.press('Control+z')
    await page.waitForTimeout(300)
    const afterUndo = await page.locator('[data-paragraph-path]').last().textContent()
    expect(afterUndo).not.toContain('APPENDED-TAIL-TEXT')

    await page.keyboard.press('Control+y')
    await page.waitForTimeout(300)
    const afterRedo = await page.locator('[data-paragraph-path]').last().textContent()
    expect(afterRedo).toContain('APPENDED-TAIL-TEXT')

    // Save, then verify on disk, then reopen fresh and check content/paragraph count survived.
    await page.keyboard.press('Control+s')
    await page.waitForTimeout(1500)

    expect(pageErrors, pageErrors.join('\n')).toEqual([])

    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(largeDocxPath))
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml).toContain('APPENDED-TAIL-TEXT')
    expect(documentXml).toContain('END-OF-DOCUMENT-MARKER-XYZZY')
    // 36 images embedded (the fixture generator's `docx` library dedupes the
    // identical 1x1 PNG bytes into ONE shared media file referenced by 36
    // relationships, not 36 separate files — verified against the freshly
    // generated fixture before ever opening it in Atlas). Confirm none of the
    // 36 blip references were lost on save.
    const mediaFiles = Object.keys(zip.files).filter((f) => f.startsWith('word/media/') && !f.endsWith('/'))
    console.log('media files after save:', mediaFiles.length)
    expect(mediaFiles.length).toBeGreaterThanOrEqual(1)
    const blipRefCount = (documentXml.match(/r:embed=/g) ?? []).length
    console.log('blip (image) references after save:', blipRefCount)
    expect(blipRefCount).toBeGreaterThanOrEqual(36)
  } finally {
    kill(app)
  }
})

test('reopen the saved 36-page docx: nothing moved or vanished', async () => {
  const app = await electron.launch({
    args: ['.', largeDocxPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))
    await page.waitForSelector('[data-paragraph-path]', { state: 'attached', timeout: 20_000 })
    await page.waitForTimeout(1500)

    // Virtualization means only pages near the viewport are actually mounted
    // — check page 1 (already in view), then scroll to the end and check the
    // last page, rather than assuming the whole document's text is in the DOM
    // at once.
    const topText = await page.evaluate(() => document.querySelector('.docx-page-stack')?.textContent ?? '')
    expect(topText).toContain('Page 1 heading')

    // `.docx-viewer` (not `.docx-page-stack`) is the real scroll container
    // (see `containerRef` in DocxViewer.tsx). A real mouse-wheel scroll (not
    // a programmatic `scrollTop =` assignment, which this Electron/CI
    // environment does not reliably turn into a native `scroll` event) gets
    // virtualization's scroll listener to actually mount the now-visible
    // pages.
    const box = await page.locator('.docx-viewer').boundingBox()
    if (box) {
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
      for (let i = 0; i < 60; i++) {
        await page.mouse.wheel(0, 800)
        await page.waitForTimeout(20)
      }
    }
    await page.waitForTimeout(500)
    const bottomText = await page.evaluate(() => document.querySelector('.docx-page-stack')?.textContent ?? '')
    expect(bottomText).toContain('Page 36 heading')
    expect(bottomText).toContain('END-OF-DOCUMENT-MARKER-XYZZY')
    expect(bottomText).toContain('APPENDED-TAIL-TEXT')

    expect(pageErrors, pageErrors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})
