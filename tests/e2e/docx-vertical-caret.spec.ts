import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableLayoutType, TableRow, TextRun } from 'docx'

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

// Enough paragraphs, each padded long enough to wrap onto several visual
// lines, that the document is guaranteed to spill onto a second page
// regardless of the exact page size/margins Atlas renders with — verifying
// DOCX-1's fix actually crosses a `.docx-page` boundary, not just a
// paragraph boundary within one page.
async function createMultiPageFixture(): Promise<string> {
  const filler = 'alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar. '
  const paragraphs = Array.from(
    { length: 40 },
    (_unused, index) => new Paragraph({ children: [new TextRun(`Paragraph ${index} ${filler.repeat(3)}`)] }),
  )
  const doc = new Document({ sections: [{ children: paragraphs }] })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-vertical-caret-multipage-'))
  const file = path.join(dir, 'multipage.docx')
  fs.writeFileSync(file, await Packer.toBuffer(doc))
  return file
}

async function createTableFixture(): Promise<string> {
  const cell = (text: string) => new TableCell({ children: [new Paragraph({ children: [new TextRun(text)] })] })
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun('Intro paragraph before the table.')] }),
          new Table({
            columnWidths: [3000, 3000, 3000],
            layout: TableLayoutType.FIXED,
            rows: [
              new TableRow({ children: [cell('A1'), cell('B1'), cell('C1')] }),
              new TableRow({ children: [cell('A2'), cell('B2'), cell('C2')] }),
            ],
          }),
          new Paragraph({ children: [new TextRun('Paragraph after the table.')] }),
        ],
      },
    ],
  })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-vertical-caret-table-'))
  const file = path.join(dir, 'table.docx')
  fs.writeFileSync(file, await Packer.toBuffer(doc))
  return file
}

const WINDOW_SIZES: ReadonlyArray<{ readonly label: string; readonly width: number; readonly height: number }> = [
  { label: '1024x768', width: 1024, height: 768 },
  { label: '1400x900', width: 1400, height: 900 },
]

async function launch(
  file: string,
  size?: { readonly width: number; readonly height: number },
): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const page = await app.firstWindow()
  if (size !== undefined) {
    await app.evaluate(
      ({ BrowserWindow }, s) => BrowserWindow.getAllWindows()[0].setSize(s.width, s.height),
      size,
    )
  }
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

// DOCX-SMALL-WINDOW-2 — at a small enough window (confirmed: CI's
// e2e-windows runner and any window at 1024x768 on this laptop), only the
// page(s) holding the current caret are guaranteed mounted (`pinnedPageIndices`
// in DocxViewer.tsx); the NEXT page can still be a bare placeholder with
// zero `.docx-page__line` elements, since plain keyboard vertical movement
// never scrolled it into the scroll-driven virtualization's visible+buffer
// range. `positionOnAdjacentLine`'s geometry search then found nothing and
// the caret got stuck at the last MOUNTED line, indistinguishable from a
// genuine document boundary. Fixed by pinning the target page and retrying
// the same move once it mounts (see `pendingVerticalMove`'s doc comment in
// DocxViewer.tsx). Parametrized across two window sizes so this stays
// proven size-independent rather than re-pinned to one laptop's default.
for (const { label, width, height } of WINDOW_SIZES) {
  test(`DOCX-1: ArrowDown repeated enough times crosses a real page boundary and lands on page 2 [${label}]`, async () => {
    const fixture = await createMultiPageFixture()
    const { app, page } = await launch(fixture, { width, height })
    try {
      const pageCountBefore = await page.locator('.docx-page').count()
      expect(pageCountBefore).toBeGreaterThan(1)

      const p0 = page.locator('[data-paragraph-path="0"]').first()
      const box = await p0.boundingBox()
      if (box === null) throw new Error('paragraph 0 has no bounding box')
      await page.mouse.click(box.x + 10, box.y + box.height / 2)
      await page.waitForTimeout(150)

      // Repeatedly press ArrowDown until the caret's own client rect sits
      // inside the SECOND `.docx-page` element — real vertical movement,
      // line by line, across the page break, not a jump to the doc's end.
      let landedOnPageTwo = false
      for (let i = 0; i < 200; i += 1) {
        await page.keyboard.press('ArrowDown')
        const onPageTwo = await page.evaluate(() => {
          const pages = Array.from(document.querySelectorAll('.docx-page'))
          const sel = window.getSelection()
          const node = sel?.focusNode
          if (pages.length < 2 || !node) return false
          return pages[1].contains(node.nodeType === Node.TEXT_NODE ? node.parentElement : (node as Element))
        })
        if (onPageTwo) {
          landedOnPageTwo = true
          break
        }
      }
      expect(landedOnPageTwo).toBe(true)

      await page.keyboard.type('PAGE2MARK', { delay: 20 })
      await page.waitForTimeout(200)
      const xml = await saveAndReadDocumentXml(page, fixture)
      expect(xml).toContain('PAGE2MARK')
      // Must NOT have landed at the very end of the document (the pre-fix
      // teleport-to-end bug) — the marker sits well before the last paragraph.
      expect(xml.indexOf('PAGE2MARK')).toBeLessThan(xml.indexOf('Paragraph 39'))
    } finally {
      kill(app)
    }
  })
}

test('DOCX-1: ArrowDown moves the caret into a table cell below, and out the other side', async () => {
  const fixture = await createTableFixture()
  const { app, page } = await launch(fixture)
  try {
    // Click into the intro paragraph, immediately above the table.
    const intro = page.locator('[data-paragraph-path="0"]').first()
    const introBox = await intro.boundingBox()
    if (introBox === null) throw new Error('intro paragraph has no bounding box')
    await page.mouse.click(introBox.x + 10, introBox.y + introBox.height / 2)
    await page.waitForTimeout(150)

    await page.keyboard.press('ArrowDown')
    await page.waitForTimeout(150)
    await page.keyboard.type('INTABLE', { delay: 20 })
    await page.waitForTimeout(200)

    const cellText = await page.evaluate(() => {
      const table = document.querySelector('.docx-page__table')
      const td = Array.from(table?.querySelectorAll('td') ?? []).find((c) => c.textContent?.includes('INTABLE'))
      return td?.textContent ?? null
    })
    expect(cellText).not.toBeNull()
    // The A1 cell's original text was "A1" — the typed marker lands
    // somewhere inside it (exact caret offset isn't the point here, landing
    // in the right cell is), so "A1" surrounds "INTABLE" rather than being
    // adjacent to it.
    expect(cellText).toMatch(/^A.*INTABLE.*1$/)

    // From inside the first row, ArrowDown enough times should reach the
    // paragraph after the table (out the other side), not skip past it.
    for (let i = 0; i < 10; i += 1) {
      await page.keyboard.press('ArrowDown')
      await page.waitForTimeout(80)
    }
    await page.keyboard.type('AFTERTABLE', { delay: 20 })
    await page.waitForTimeout(200)

    const xml = await saveAndReadDocumentXml(page, fixture)
    expect(xml).toContain('INTABLE')
    expect(xml).toContain('AFTERTABLE')
    // The after-table marker must land after the in-table one in document
    // order (moving further down, not teleporting to the doc's start/end).
    expect(xml.indexOf('AFTERTABLE')).toBeGreaterThan(xml.indexOf('INTABLE'))
  } finally {
    kill(app)
  }
})
