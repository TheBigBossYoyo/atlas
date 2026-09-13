import path from 'node:path'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test } from '@playwright/test'

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

const fixtures = [
  { format: 'markdown', fileName: 'sample.md' },
  { format: 'text', fileName: 'sample.txt' },
  { format: 'code', fileName: 'sample.ts' },
  { format: 'csv', fileName: 'sample.csv' },
  { format: 'tsv', fileName: 'sample.tsv' },
  { format: 'xlsx', fileName: 'sample.xlsx' },
  { format: 'ods', fileName: 'sample.ods' },
  { format: 'docx', fileName: 'sample.docx' },
  { format: 'rtf', fileName: 'sample.rtf' },
  { format: 'odt', fileName: 'sample.odt' },
  { format: 'pdf', fileName: 'sample.pdf' },
  { format: 'pptx', fileName: 'sample.pptx' },
  { format: 'odp', fileName: 'sample.odp' },
] as const

test.beforeAll(() => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
})

for (const fixture of fixtures) {
  test(`loads ${fixture.format} fixture`, async () => {
    const electronApp = await electron.launch({
      args: ['.', path.join(fixtureDir, fixture.fileName)],
      cwd: projectRoot,
      env: {
        ...process.env,
        CI: '1',
        PLAYWRIGHT: '1',
      },
    })

    try {
      const page = await electronApp.firstWindow()
      const consoleErrors: string[] = []

      page.on('console', (message) => {
        if (message.type() === 'error') {
          consoleErrors.push(message.text())
        }
      })

      await expect(page.locator(`[data-viewer="${fixture.format}"]`)).toHaveCount(1)
      await expect(page.locator('.statusbar__file')).toHaveText(fixture.fileName, { timeout: 15_000 })
      await page.waitForTimeout(500)
      expect(consoleErrors, consoleErrors.join('\n')).toEqual([])
    } finally {
      await electronApp.close()
    }
  })
}

// Wave 3-P / PDF-16 / P12: exercises the rewritten PdfViewer end-to-end
// against a real multi-page, mixed-portrait/landscape PDF with a real
// outline — page count, zoom, page navigation, and in-viewer find — beyond
// what the single-page `sample.pdf` fixture in the per-format loop above
// can (it never scrolls, zooms, or searches).
test('PDF viewer: page count, zoom, navigation, and find on a multi-page document', async () => {
  const electronApp = await electron.launch({
    args: ['.', path.join(fixtureDir, 'sample-multipage.pdf')],
    cwd: projectRoot,
    env: {
      ...process.env,
      CI: '1',
      PLAYWRIGHT: '1',
    },
  })

  try {
    const page = await electronApp.firstWindow()
    const consoleErrors: string[] = []
    page.on('console', (message) => {
      if (message.type() === 'error') {
        consoleErrors.push(message.text())
      }
    })

    await expect(page.locator('[data-viewer="pdf"]')).toHaveCount(1)
    await expect(page.locator('.statusbar__file')).toHaveText('sample-multipage.pdf', {
      timeout: 15_000,
    })

    // Page count: the toolbar's "/ N" indicator and one page-wrapper per page.
    await expect(page.locator('.pdf-viewer__page-indicator')).toContainText('/ 5', {
      timeout: 15_000,
    })
    await expect(page.locator('.pdf-viewer__page-wrapper')).toHaveCount(5)

    // Zoom: default 100%, Zoom In steps to 125%.
    const zoomButton = page.locator('.pdf-viewer__zoom-button')
    await expect(zoomButton).toContainText('100%')
    await page.locator('[title="Zoom In (Ctrl++)"]').click()
    await expect(zoomButton).toContainText('125%')

    // Navigation: jump to page 3 via the page-number input.
    const pageInput = page.getByLabel('Page number')
    await pageInput.fill('3')
    await pageInput.press('Enter')
    await expect(pageInput).toHaveValue('3', { timeout: 15_000 })
    await expect(page.locator('.pdf-viewer__page-wrapper[data-page="3"] canvas')).toBeVisible()

    // Find: Ctrl+F opens the find bar; the query matches text that only
    // exists on page 3 of this fixture.
    await page.keyboard.press('Control+f')
    const findInput = page.getByPlaceholder('Find in document…')
    await expect(findInput).toBeVisible()
    await findInput.fill('Findable needle')
    await expect(page.locator('.pdf-viewer__find-count')).toHaveText('1 of 1', { timeout: 15_000 })

    expect(consoleErrors, consoleErrors.join('\n')).toEqual([])
  } finally {
    await electronApp.close()
  }
})

// Regression coverage for the mermaid dependency bump (P1.16): the markdown
// fixture embeds a fenced ```mermaid block, so this asserts the diagram
// actually renders to a real <svg> (not the Mermaid.tsx error fallback)
// with no console errors, catching a future mermaid/markdown regression
// that the per-format loop above wouldn't (it never inspects diagram output).
test('renders a mermaid diagram in the markdown fixture', async () => {
  const electronApp = await electron.launch({
    args: ['.', path.join(fixtureDir, 'sample.md')],
    cwd: projectRoot,
    env: {
      ...process.env,
      CI: '1',
      PLAYWRIGHT: '1',
    },
  })

  try {
    const page = await electronApp.firstWindow()
    const consoleErrors: string[] = []

    page.on('console', (message) => {
      if (message.type() === 'error') {
        consoleErrors.push(message.text())
      }
    })

    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1)
    await expect(page.locator('.mermaid--error')).toHaveCount(0)
    await expect(page.locator('.mermaid svg')).toBeVisible({ timeout: 15_000 })
    expect(consoleErrors, consoleErrors.join('\n')).toEqual([])
  } finally {
    await electronApp.close()
  }
})
