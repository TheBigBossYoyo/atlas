import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { Document, Packer, Paragraph, Table, TableCell, TableLayoutType, TableRow, TextRun } from 'docx'

/**
 * DOCX-17 — the caret could not be placed in any table cell, including a
 * pre-existing, correctly rendered table from a normal `.docx`: clicking a
 * cell (or Tab/ArrowDown into one) landed the caret in the paragraph after
 * the table, and cell merge was unreachable as a direct consequence.
 *
 * Root cause (see `src/docx/render/PageView.tsx`'s and
 * `src/docx/editor/Cursor.ts`'s own "Table-cell paragraph-path resolution"
 * sections): table cell lines rendered with no `data-paragraph-path`
 * attribute at all — `PageView.tsx`'s table-row renderer never threaded one
 * through, because the layout pipeline's per-cell line list
 * (`LaidOutCell.contentLines`) has no paragraph identity attached to pass
 * down in the first place — so every caret/selection helper that keys off
 * that attribute silently failed to resolve ANY position inside a table.
 * This only reproduces against the real rendered DOM (the unit suite mocks
 * past exactly this), hence a real-app e2e test.
 */

const projectRoot = process.cwd()

async function createTableFixture(): Promise<string> {
  // Wide fixed columns (twips) — like `scripts/generate-docx-corpus.mjs`'s
  // `table-fixed-grid-merged-cells` fixture — so a cell's short "A1"-style
  // text sits flush left in a much wider column: clicking near the cell's
  // OWN left edge (not its text's) reliably lands the caret at the start of
  // that text regardless of exactly how wide the glyphs measure.
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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-docx-table-cursor-'))
  const file = path.join(dir, 'table.docx')
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

async function cellRect(page: Page, text: string): Promise<{ x: number; y: number; w: number; h: number }> {
  const rect = await page.evaluate((cellText) => {
    const table = document.querySelector('.docx-page__table')
    const td = Array.from(table?.querySelectorAll('td') ?? []).find((c) => c.textContent?.trim() === cellText)
    if (!td) {
      throw new Error(`no cell with text ${cellText}`)
    }
    const box = td.getBoundingClientRect()
    return { x: box.x, y: box.y, w: box.width, h: box.height }
  }, text)
  return rect
}

/**
 * Clicks near the cell's OWN left edge rather than its center — the column
 * is much wider than its short "A1"-style text (see `createTableFixture`),
 * so a center click can land past the end of the text or, worse, roughly at
 * its horizontal midpoint (ambiguous between the first and second
 * character). A small fixed inset from the cell's left edge instead
 * reliably lands the caret at the very start of that text, independent of
 * exactly how many pixels the glyphs measure.
 */
async function clickCell(page: Page, text: string): Promise<void> {
  const rect = await cellRect(page, text)
  await page.mouse.click(rect.x + 4, rect.y + rect.h / 2)
  await page.waitForTimeout(200)
}

/** Whether the native DOM selection (the visual caret) is currently inside a
 * `<td>` — i.e. whether `positionToDomRange`/`syncSelectionToDom` actually
 * found a DOM location for the model's current table-cell position. */
async function selectionInsideTableCell(page: Page): Promise<boolean> {
  return page.evaluate(() => {
    const sel = window.getSelection()
    const node = sel?.anchorNode ?? null
    const el = node === null ? null : node.nodeType === 3 ? node.parentElement : (node as Element)
    return el?.closest('td') !== null
  })
}

test('clicking a pre-existing table cell places the caret in that cell, and typing lands there', async () => {
  const file = await createTableFixture()
  const { app, page } = await launch(file)
  try {
    await clickCell(page, 'B1')
    expect(await selectionInsideTableCell(page)).toBe(true)

    await page.keyboard.type('XY')
    await page.waitForTimeout(200)

    const cells = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.docx-page__table td')).map((td) => td.textContent),
    )
    expect(cells).toContain('XYB1')
    // Nothing else moved — DOCX-17 lost the keystrokes to whatever paragraph
    // happened to be nearest instead (the paragraph after the table).
    const afterParagraph = await page.evaluate(() => {
      const lines = Array.from(document.querySelectorAll('[data-paragraph-path]'))
      const last = lines.find((el) => el.textContent?.includes('Paragraph after the table'))
      return last?.textContent ?? null
    })
    expect(afterParagraph).toBe('Paragraph after the table.')
  } finally {
    kill(app)
  }
})

test('Tab moves to the next table cell and Shift+Tab to the previous one', async () => {
  const file = await createTableFixture()
  const { app, page } = await launch(file)
  try {
    await clickCell(page, 'A1')
    expect(await selectionInsideTableCell(page)).toBe(true)

    // Tab in a cell selects that cell's whole content (Word parity), so
    // typing replaces the placeholder text with the new value.
    await page.keyboard.press('Tab')
    await page.waitForTimeout(150)
    await page.keyboard.type('B1-edited')
    await page.waitForTimeout(150)

    let cells = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.docx-page__table td')).map((td) => td.textContent),
    )
    expect(cells).toContain('B1-edited')
    expect(cells).not.toContain('B1')

    await page.keyboard.press('Shift+Tab')
    await page.waitForTimeout(150)
    await page.keyboard.type('A1-edited')
    await page.waitForTimeout(150)

    cells = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.docx-page__table td')).map((td) => td.textContent),
    )
    expect(cells).toContain('A1-edited')
    expect(cells).not.toContain('A1')
  } finally {
    kill(app)
  }
})

test('clicking a table cell makes cell merge reachable from the Edit Table toolbar button', async () => {
  // DXE-14's "Edit Table" toolbar button (and the equivalent right-click
  // menu) is disabled whenever `ToolbarState.insideTable` is false, which is
  // computed straight from `findEnclosingTable(document,
  // range.focus.paragraphPath)` — i.e. directly from whatever `range` the
  // last click/Tab left behind. DOCX-17 meant `range.focus.paragraphPath`
  // could never actually address a table cell, so this button could never
  // be reached from inside a table no matter how a cell was clicked.
  const file = await createTableFixture()
  const { app, page } = await launch(file)
  try {
    await page.getByRole('tab', { name: 'Insert' }).click()
    const editTableButton = page.getByRole('button', { name: 'Edit Table' })
    await expect(editTableButton).toBeDisabled()

    await clickCell(page, 'A2')
    expect(await selectionInsideTableCell(page)).toBe(true)
    await expect(editTableButton).toBeEnabled()

    await editTableButton.click()
    await expect(page.locator('.docx-toolbar__menu-item', { hasText: 'Merge Right' })).toBeVisible()
  } finally {
    kill(app)
  }
})

test('typed text inside a table cell survives save, close, and reopen', async () => {
  const file = await createTableFixture()
  const { app, page } = await launch(file)
  try {
    await clickCell(page, 'C2')
    expect(await selectionInsideTableCell(page)).toBe(true)
    await page.keyboard.type('saved-')
    await page.waitForTimeout(200)

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(400)
  } finally {
    kill(app)
  }

  // Verify directly in the saved archive — this is the part of the
  // acceptance criteria that doesn't depend on the editor's own rendering.
  const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
  const documentXml = await zip.file('word/document.xml')!.async('string')
  expect(documentXml).toContain('saved-C2')

  // ...and reopening the saved file in the app shows the same text back,
  // inside the same table cell (not merged into an adjacent one, not
  // dropped, not duplicated elsewhere).
  const { app: app2, page: page2 } = await launch(file)
  try {
    const cells = await page2.evaluate(() =>
      Array.from(document.querySelectorAll('.docx-page__table td')).map((td) => td.textContent),
    )
    expect(cells).toContain('saved-C2')
    expect(cells.filter((text) => text === 'saved-C2')).toHaveLength(1)
  } finally {
    kill(app2)
  }
})
