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
