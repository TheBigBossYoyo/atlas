import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { AlignmentType, Document, Packer, Paragraph, TextRun } from 'docx'

/**
 * USR-01…USR-13 — owner-reported DOCX editor regressions. Every scenario
 * drives the real Electron app with real mouse/keyboard input, because the
 * unit suite passed while typing into a DOCX was completely broken.
 */

const projectRoot = process.cwd()

// D29 follow-up — an edit is now spliced into the paragraph's own runs
// (keeping whatever text/formatting the edit didn't touch, rather than
// flattening the whole paragraph into one new run), so a shared suffix like
// "Draft header" -> "Final header"'s " header" can legitimately end up in
// its own `<w:t>` after the rewritten run. Strip tags before asserting on
// text content so these checks don't depend on exactly how many runs the
// diff split the result across.
const textOnly = (xml: string): string => xml.replace(/<[^>]+>/g, '')

async function createFixture(): Promise<string> {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun('Atlas editor title line')] }),
          new Paragraph({
            children: [
              new TextRun(
                'The quick brown fox jumps over the lazy dog while the editor keeps every keystroke and the find panel finds the words.',
              ),
            ],
          }),
          new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun('Second paragraph stays left aligned.')] }),
        ],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-editor-'))
  const file = path.join(dir, 'editor.docx')
  fs.writeFileSync(file, await Packer.toBuffer(doc))
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
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

const selectionLength = (page: Page): Promise<number> => page.evaluate(() => window.getSelection()?.toString().length ?? 0)

async function doubleClickWord(page: Page, paragraph: number, word: number): Promise<void> {
  await page.locator(`[data-paragraph-path="${paragraph}"] .docx-run`).nth(word).dblclick()
  await page.waitForTimeout(200)
}

let fixture = ''
test.beforeAll(async () => {
  fixture = await createFixture()
})

test('USR-01: typing inserts text at the caret', async () => {
  const { app, page } = await launch(fixture)
  try {
    const line = page.locator('[data-paragraph-path="1"]').first()
    const box = await line.boundingBox()
    if (box === null) throw new Error('line not rendered')
    await page.mouse.click(box.x + 20, box.y + box.height / 2)
    await page.keyboard.type('XYZ', { delay: 30 })
    await expect.poll(() => paragraphText(page, 1)).toContain('XYZ')
  } finally {
    kill(app)
  }
})

test('USR-02/USR-07: Ctrl+A selects everything, Ctrl+Z/Ctrl+Y undo and redo', async () => {
  const { app, page } = await launch(fixture)
  try {
    await page.locator('[data-paragraph-path="1"] .docx-run').first().click()
    await page.keyboard.type('Q')
    await expect.poll(() => paragraphText(page, 1)).toContain('Q')
    await page.keyboard.press('Control+z')
    await expect.poll(() => paragraphText(page, 1)).not.toContain('Q')
    await page.keyboard.press('Control+y')
    await expect.poll(() => paragraphText(page, 1)).toContain('Q')

    await page.keyboard.press('Control+a')
    const total = await page.evaluate(() => document.querySelector('.docx-viewer__surface')?.textContent?.length ?? 0)
    await expect.poll(() => selectionLength(page)).toBeGreaterThanOrEqual(total - 2)
  } finally {
    kill(app)
  }
})

test('USR-03/USR-09: underline toggles on and off and the selection survives formatting', async () => {
  const { app, page } = await launch(fixture)
  try {
    await doubleClickWord(page, 1, 1)
    const underlined = (): Promise<number> =>
      page.evaluate(
        () =>
          Array.from(document.querySelectorAll('[data-paragraph-path="1"] .docx-run')).filter(
            (el) => getComputedStyle(el).textDecorationLine.includes('underline'),
          ).length,
      )
    const selected = await selectionLength(page)
    expect(selected).toBeGreaterThan(0)

    await page.keyboard.press('Control+u')
    await expect.poll(underlined).toBeGreaterThan(0)
    await expect.poll(() => selectionLength(page)).toBe(selected)

    await page.keyboard.press('Control+u')
    await expect.poll(underlined).toBe(0)

    await page.keyboard.press('Control+b')
    await expect.poll(() => selectionLength(page)).toBe(selected)
  } finally {
    kill(app)
  }
})

test('USR-04: aligning a triple-clicked line does not re-align the next paragraph', async () => {
  const { app, page } = await launch(fixture)
  try {
    const left = (p: number): Promise<number> =>
      page.evaluate((pp) => document.querySelector(`[data-paragraph-path="${pp}"]`)?.getBoundingClientRect().left ?? 0, p)
    const before = await left(1)
    await page.locator('[data-paragraph-path="0"]').first().click({ clickCount: 3 })
    await page.getByRole('button', { name: 'Align right', exact: true }).click()
    await expect.poll(() => left(0)).toBeGreaterThan(before)
    expect(await left(1)).toBe(before)
  } finally {
    kill(app)
  }
})

test('USR-08: highlight applies a background to the selected word', async () => {
  const { app, page } = await launch(fixture)
  try {
    await doubleClickWord(page, 1, 2)
    await page.getByRole('button', { name: 'Highlight Color', exact: true }).click()
    await page.locator('.docx-toolbar__popover [aria-label^="Color #"]').first().click()
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            Array.from(document.querySelectorAll('[data-paragraph-path="1"] .docx-run')).filter(
              (el) => getComputedStyle(el).backgroundColor !== 'rgba(0, 0, 0, 0)',
            ).length,
        ),
      )
      .toBeGreaterThan(0)
  } finally {
    kill(app)
  }
})

test('USR-06: find keeps focus in the Find box and moves through matches', async () => {
  const { app, page } = await launch(fixture)
  try {
    await page.locator('[data-paragraph-path="1"] .docx-run').first().click()
    await page.keyboard.press('Control+f')
    const input = page.getByLabel('Find', { exact: true })
    await input.fill('the')
    await page.keyboard.press('Enter')
    await page.keyboard.press('Enter')
    await expect(input).toBeFocused()
    await expect(page.getByText(/\d+ of \d+/)).toBeVisible()
    const beforeText = await paragraphText(page, 1)
    expect(beforeText).toContain('The quick brown fox')
  } finally {
    kill(app)
  }
})

test('USR-05/USR-13: clicking beside a line puts the caret on that line; page counts agree', async () => {
  const { app, page } = await launch(fixture)
  try {
    const line = page.locator('[data-paragraph-path="2"]').first()
    const box = await line.boundingBox()
    if (box === null) throw new Error('line not rendered')
    await page.mouse.click(box.x + box.width + 80, box.y + box.height / 2)
    await page.keyboard.type('!')
    await expect.poll(() => paragraphText(page, 2)).toMatch(/!\s*$/)

    const meta = await page.locator('.docx-viewer__meta').innerText()
    const status = await page.locator('.statusbar').innerText()
    const metaPages = Number(/Pages: (\d+)/.exec(meta)?.[1])
    expect(status).toContain(`${metaPages} ${metaPages === 1 ? 'page' : 'pages'}`)
  } finally {
    kill(app)
  }
})

test('D29: the header and footer can be edited and saved', async () => {
  const { Document, Packer, Paragraph, TextRun, Header, Footer } = await import('docx')
  const doc = new Document({
    sections: [
      {
        headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun('Draft header')] })] }) },
        footers: { default: new Footer({ children: [new Paragraph({ children: [new TextRun('Page footer')] })] }) },
        children: [new Paragraph({ children: [new TextRun('Body text')] })],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-hf-'))
  const file = path.join(dir, 'with-header.docx')
  fs.writeFileSync(file, await Packer.toBuffer(doc))

  const { app, page } = await launch(file)
  try {
    // A narrow window wraps the document toolbar; give it room so the action is on one row.
    await page.setViewportSize({ width: 1400, height: 900 })
    await page.waitForTimeout(300)
    await page.getByRole('button', { name: 'Header and footer' }).click()
    const headerField = page.getByRole('textbox', { name: 'Header' })
    await expect(headerField).toHaveValue('Draft header')

    await headerField.fill('Final header')
    await page.getByRole('textbox', { name: 'Footer' }).click() // blur commits
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeVisible()

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(400)

    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
    const headerPart = Object.keys(zip.files).find((name) => /^word\/header\d+\.xml$/.test(name))!
    expect(textOnly(await zip.file(headerPart)!.async('string'))).toContain('Final header')
  } finally {
    kill(app)
  }
})

test('D29 follow-up: Ctrl+S while the header field is still focused saves what was typed, not the stale text', async () => {
  const { Document, Packer, Paragraph, TextRun, Header } = await import('docx')
  const doc = new Document({
    sections: [
      {
        headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun('Draft header')] })] }) },
        children: [new Paragraph({ children: [new TextRun('Body text')] })],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-hf-ctrls-'))
  const file = path.join(dir, 'with-header.docx')
  fs.writeFileSync(file, await Packer.toBuffer(doc))

  const { app, page } = await launch(file)
  try {
    await page.setViewportSize({ width: 1400, height: 900 })
    await page.waitForTimeout(300)
    await page.getByRole('button', { name: 'Header and footer' }).click()
    const headerField = page.getByRole('textbox', { name: 'Header' })
    await expect(headerField).toHaveValue('Draft header')

    // `fill` focuses the field and leaves it focused — never blurring it, so
    // the only way the typed text can reach the saved file is if Ctrl+S
    // itself commits the still-pending edit (see `flushHeaderFooterEdits` in
    // DocxViewer.tsx) rather than saving the stale pre-edit `documentModel`.
    await headerField.fill('Typed without blurring')
    await expect(headerField).toBeFocused()

    const before = fs.statSync(file).mtimeMs
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(400)

    // The field still shows the typed text (the commit didn't blur/reset it)...
    await expect(headerField).toHaveValue('Typed without blurring')
    // ...and it's what actually landed on disk.
    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
    const headerPart = Object.keys(zip.files).find((name) => /^word\/header\d+\.xml$/.test(name))!
    const headerText = textOnly(await zip.file(headerPart)!.async('string'))
    expect(headerText).toContain('Typed without blurring')
    expect(headerText).not.toContain('Draft header')
  } finally {
    kill(app)
  }
})

test('Ctrl+Shift+S (Save As) actually opens the native Save dialog for a DOCX, not a silent no-op', async () => {
  // Found by driving the real app: the Shortcuts modal documents Ctrl+Shift+S
  // as a global "Save As", and App.tsx's saveFileAs() already routes plain
  // Ctrl+S to whichever non-markdown viewer is active — but nothing wired an
  // equivalent path for Save As, so it silently did nothing for a DOCX. Worse,
  // DocxViewer's own `useViewerShortcuts` combo-claiming filter treated any
  // Ctrl+S (shifted or not) as "reserved" and swallowed the keystroke before
  // it could even reach the shell dispatcher, independent of the missing
  // routing. Both are fixed: the viewer only claims the unshifted combo now,
  // and `saveFileAs()` falls back to a new `registerSaveAs`/`saveAs()`
  // capability every save-capable viewer (DOCX, spreadsheet, slides)
  // registers, mirroring `registerSave`/`save()`.
  const { app, page } = await launch(await createFixture())
  try {
    let saveDialogCalls = 0
    await app.evaluate(({ dialog }) => {
      const original = dialog.showSaveDialog.bind(dialog)
      dialog.showSaveDialog = (async (...args: Parameters<typeof dialog.showSaveDialog>) => {
        ;(globalThis as { __saveAsCalls?: number }).__saveAsCalls =
          ((globalThis as { __saveAsCalls?: number }).__saveAsCalls ?? 0) + 1
        return original(...args)
      }) as typeof dialog.showSaveDialog
    })

    await page.locator('.docx-page__column').first().click({ position: { x: 5, y: 5 } })
    await page.keyboard.press('Control+Shift+S')
    await page.waitForTimeout(500)

    saveDialogCalls = await app.evaluate(() => (globalThis as { __saveAsCalls?: number }).__saveAsCalls ?? 0)
    expect(saveDialogCalls).toBe(1)
  } finally {
    kill(app)
  }
})

test('after Save As, the tab follows the document to its new file', async () => {
  // Found by driving the real app: Save As wrote the new file but left the
  // tab pointing at the file the document was opened from, so leaving the tab
  // and coming back re-read the ORIGINAL file over the user's work, and the
  // next Ctrl+S wrote to neither file.
  const source = await createFixture()
  const target = path.join(path.dirname(source), 'saved-as.docx')
  const other = path.join(path.dirname(source), 'other.md')
  fs.writeFileSync(other, '# Other document\n')

  const app = await electron.launch({
    args: ['.', source],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    await page.waitForSelector('[data-paragraph-path]', { timeout: 20_000 })
    await page.waitForTimeout(800)

    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath })) as typeof dialog.showSaveDialog
    }, target)

    await page.locator('.docx-page__column').first().click({ position: { x: 5, y: 5 } })
    await page.keyboard.press('Control+End')
    await page.keyboard.type('Saved as a new file.')
    await page.keyboard.press('Control+Shift+S')

    await expect.poll(() => fs.existsSync(target), { timeout: 15_000 }).toBe(true)
    // The tab, and the title bar, now name the file the document was written to.
    await expect(page.getByRole('tab', { name: 'saved-as.docx' })).toBeVisible({ timeout: 15_000 })

    // Leave and come back: the document must still be the saved one.
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, other)
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    await expect(page.getByRole('tab', { name: 'other.md' })).toHaveAttribute('aria-selected', 'true', { timeout: 20_000 })
    await page.getByRole('tab', { name: 'saved-as.docx' }).click()
    await expect(page.locator('[data-paragraph-path]').first()).toBeVisible({ timeout: 20_000 })
    await expect(page.locator('.docx-page-stack')).toContainText('Saved as a new file.', { timeout: 20_000 })

    // And a plain save now writes to the new file, not the old one.
    const sourceBefore = fs.readFileSync(source)
    await page.locator('.docx-page__column').first().click({ position: { x: 5, y: 5 } })
    await page.keyboard.press('Control+End')
    await page.keyboard.type(' Edited again.')
    const targetBefore = fs.readFileSync(target)
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.readFileSync(target).equals(targetBefore), { timeout: 15_000 }).toBe(false)
    expect(fs.readFileSync(source).equals(sourceBefore)).toBe(true)
  } finally {
    kill(app)
  }
})

test('DOCX-3: typing continues to work immediately after Save As, without clicking back into the document', async () => {
  // Found by driving the real app: the native Save dialog Save As opens
  // takes OS-level focus away from the whole Electron window; nothing gave
  // it back to the editor once the dialog closed, so `document.activeElement`
  // was left as `<body>` and every keystroke typed right after Save As
  // vanished silently — no error, no visual sign, until the user noticed and
  // clicked back into the page.
  //
  // FIX-1/2: src/viewers/DocxViewer.tsx's handleSaveInternal now re-focuses
  // editorRootRef and resyncs the selection once Save As's saveBinaryFile
  // call resolves (verified directly: a temporary console.log right after
  // that call showed `document.activeElement` correctly back on the editor
  // surface).
  //
  // FIX-2/2 — NOT done here, outside this task's file ownership: a moment
  // later, once `reportSavedPath` (called just before the fix above, to tell
  // the shell the tab's file renamed) updates the active tab's path,
  // `src/components/ViewerRouter.tsx` remounts a BRAND NEW DocxEditor
  // instance — `<ViewerErrorBoundary key={file.path} ...>` there keys the
  // whole viewer subtree by `file.path`, and Save As always changes it — so
  // the freshly restored focus is immediately discarded along with the old
  // instance. Confirmed via a temporary render-count console.log: the
  // component remounts (a fresh `file.path` value in a fresh render) within
  // ~300ms of the focus fix running. See this task's final report for the
  // exact fix ViewerRouter.tsx needs (it needs to distinguish "the current
  // document was renamed by its own Save As" from "the user opened a
  // different file" before deciding whether to remount).
  //
  // Pinned via `test.fail()` rather than deleted or weakened: this documents
  // the exact remaining gap and will flip to an unexpected PASS (Playwright
  // then reports it as a failure needing promotion back to a plain `test`)
  // the moment ViewerRouter.tsx's half of the fix lands.
  test.fail()

  const source = await createFixture()
  const target = path.join(path.dirname(source), 'saveas-continue-typing.docx')

  const app = await electron.launch({
    args: ['.', source],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    await page.waitForSelector('[data-paragraph-path]', { timeout: 20_000 })
    await page.waitForTimeout(800)

    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath })) as typeof dialog.showSaveDialog
    }, target)

    const line = page.locator('[data-paragraph-path="1"]').first()
    const box = await line.boundingBox()
    if (box === null) throw new Error('paragraph 1 has no bounding box')
    await page.mouse.click(box.x + 10, box.y + box.height / 2)
    await page.keyboard.type('Before save-as. ')

    await page.keyboard.press('Control+Shift+S')
    await expect.poll(() => fs.existsSync(target), { timeout: 15_000 }).toBe(true)
    await expect(page.getByRole('tab', { name: 'saveas-continue-typing.docx' })).toBeVisible({ timeout: 15_000 })

    // No click here — this is the exact repro: type right away.
    await page.waitForTimeout(300)
    await page.keyboard.type('POSTSAVEASMARK')
    await page.waitForTimeout(300)

    expect(await paragraphText(page, 1)).toContain('POSTSAVEASMARK')

    const before = fs.statSync(target).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(target).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(300)

    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(target))
    const xml = textOnly(await zip.file('word/document.xml')!.async('string'))
    expect(xml).toContain('POSTSAVEASMARK')
  } finally {
    kill(app)
  }
})
