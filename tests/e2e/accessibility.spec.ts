/**
 * Accessibility audit — WCAG 2.1 A/AA, via axe-core, across every viewer.
 *
 * WHY THIS EXISTS. Atlas had never had one. There was keyboard coverage
 * (`keyboard-only.spec.ts`, `focus-management.spec.ts`) and individual aria
 * labels added as bugs were found (FROZENROWS-I18N-1, the comments pane's
 * "Resolved" badge being text rather than opacity), but nothing had ever checked
 * the whole surface, and nothing would notice a regression.
 *
 * WHAT THIS DOES AND DOES NOT PROVE. axe-core finds machine-checkable violations:
 * a control with no accessible name, a contrast ratio below 4.5:1, a broken
 * landmark or heading structure, a role with missing required attributes. It
 * cannot tell you whether a screen reader makes the document READABLE — whether
 * page boundaries are announced, whether the caret position is spoken, whether
 * the comments pane reads in a sensible order. Those need a real screen reader
 * and a person, which has still not happened; this is the floor, not the ceiling,
 * and `docs/KNOWN_LIMITATIONS.md` says so.
 *
 * Each viewer is audited in its own test with its own Electron launch, so one
 * viewer's violations are reported with that viewer named rather than as a single
 * opaque failure. Violations are printed in full — rule id, impact, and the
 * offending selector — because "axe found 3 issues" is not actionable.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

/**
 * The tags axe should run. `wcag2a`/`wcag2aa`/`wcag21a`/`wcag21aa` are the
 * conformance sets; `best-practice` is deliberately excluded, since it flags
 * house-style opinions (e.g. "prefer a landmark") that are not accessibility
 * failures and would bury the ones that are.
 */
const WCAG_TAGS = ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function launch(file?: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: file === undefined ? ['.'] : ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const page = await app.firstWindow()
  await page.waitForSelector('body', { timeout: 30_000 })
  return { app, page }
}

type Violation = {
  readonly id: string
  readonly impact?: string | null
  readonly help: string
  readonly nodes: ReadonlyArray<{ readonly target: ReadonlyArray<unknown>; readonly failureSummary?: string }>
}

/** One line per offending node, so a failure says what to change and where. */
function describe(violations: ReadonlyArray<Violation>): string {
  return violations
    .flatMap((violation) =>
      violation.nodes.map(
        (node) =>
          `[${violation.impact ?? 'unknown'}] ${violation.id}: ${violation.help} @ ${node.target.join(' ')}` +
          (node.failureSummary === undefined ? '' : ` — ${node.failureSummary.replace(/\s+/g, ' ')}`),
      ),
    )
    .join('\n')
}

/**
 * axe-core's source, injected by hand.
 *
 * `@axe-core/playwright` cannot be used here: it calls `browserContext.newPage()`
 * while working out frame support, and Electron answers `Target.createTarget: Not
 * supported`, so every audit fails before running. Injecting the library and
 * calling `axe.run` is what that wrapper does anyway, minus the part Electron
 * rejects.
 */
const AXE_SOURCE = fs.readFileSync(path.join(projectRoot, 'node_modules', 'axe-core', 'axe.min.js'), 'utf8')

async function audit(page: Page): Promise<ReadonlyArray<Violation>> {
  await page.evaluate(AXE_SOURCE)
  return (await page.evaluate(
    async (tags) =>
      (
        await (window as unknown as { axe: { run: (options: unknown) => Promise<{ violations: unknown[] }> } }).axe.run({
          runOnly: { type: 'tag', values: tags },
          // A `resultTypes` narrowing would hide the selectors that make a
          // failure actionable, so everything is kept.
        })
      ).violations,
    WCAG_TAGS,
  )) as unknown as ReadonlyArray<Violation>
}

function tmpXlsx(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-a11y-'))
  const file = path.join(dir, 'audit.xlsx')
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['Region', 'Units', 'Revenue'],
      ['North', 12, 4200],
      ['South', 8, 3100],
    ]),
    'Sales',
  )
  fs.writeFileSync(file, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)
  return file
}

/** The fixtures already in the repo, so this audits what the other specs audit. */
function fixture(name: string): string {
  const file = path.join(fixtureDir, name)
  if (!fs.existsSync(file)) throw new Error(`missing fixture ${name} — run tests/e2e/fixtures/generate.mjs`)
  return file
}

test.describe('accessibility (axe-core, WCAG 2.1 A/AA)', () => {
  test('the empty shell — no document open', async () => {
    const { app, page } = await launch()
    try {
      await page.waitForTimeout(1500)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the Word viewer', async () => {
    const { app, page } = await launch(fixture('sample.docx'))
    try {
      await page.waitForSelector('[data-paragraph-path]', { timeout: 30_000 })
      await page.waitForTimeout(1000)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the Word viewer with the comments pane open', async () => {
    // The pane was added late and grew an un-resolve affordance later still, so
    // it is the newest markup in the app and the least likely to have been
    // checked by hand.
    const { app, page } = await launch(fixture('sample.docx'))
    try {
      await page.waitForSelector('[data-paragraph-path]', { timeout: 30_000 })
      await page.keyboard.press('Control+Alt+m')
      await page.waitForTimeout(1200)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the spreadsheet viewer', async () => {
    const { app, page } = await launch(tmpXlsx())
    try {
      await page.waitForSelector('.spreadsheet-viewer__grid canvas', { timeout: 30_000 })
      await page.waitForTimeout(1200)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the markdown viewer', async () => {
    const { app, page } = await launch(fixture('sample.md'))
    try {
      await page.waitForTimeout(2000)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the slides viewer', async () => {
    const { app, page } = await launch(fixture('sample.pptx'))
    try {
      await page.waitForTimeout(2500)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the PDF viewer', async () => {
    const { app, page } = await launch(fixture('sample.pdf'))
    try {
      await page.waitForTimeout(2500)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the code viewer', async () => {
    const { app, page } = await launch(fixture('sample.ts'))
    try {
      await page.waitForTimeout(2000)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })

  test('the keyboard shortcuts dialog', async () => {
    // A modal is where aria goes wrong quietly: no dialog role, no label, focus
    // not trapped, nothing that tells a screen reader the page behind it is gone.
    const { app, page } = await launch(fixture('sample.md'))
    try {
      await page.waitForTimeout(1500)
      await page.keyboard.press('Control+/')
      await page.waitForTimeout(800)
      const violations = await audit(page)
      expect(violations, describe(violations)).toEqual([])
    } finally {
      kill(app)
    }
  })
})
