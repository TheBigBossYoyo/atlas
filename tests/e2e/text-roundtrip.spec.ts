import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * NIGHT/text-roundtrip — text-class files (markdown, code/text, CSV/TSV)
 * must round-trip byte-faithfully: encoding, BOM, line endings, and (for
 * CSV) the delimiter. Verified live against the real app, asserting on the
 * SAVED BYTES read back from disk (never the screen) — see
 * `electron/lib/textDecoding.cjs`, `src/utils/textDecoding.ts`,
 * `electron/main.cjs`'s `readMarkdownFile`/`save-file` handler, `App.tsx`'s
 * markdown save path, and `CodeViewer.tsx`'s save path for the fix itself.
 *
 * SHELL-1/SHELL-2 (markdown: CRLF -> LF, BOM stripped) and the wider
 * UTF-16/Windows-1252-silently-re-encoded-as-UTF-8 gap are FULLY fixed and
 * proven below.
 *
 * SHEET-6 (CSV delimiter rewritten) is fully fixed and proven below.
 * SHEET-7/SHEET-8 (CSV BOM stripped / LF -> CRLF) are NOT fixed by this
 * commit — `documentToDelimitedText` (the only file in
 * `src/viewers/spreadsheet/` this fix owns) has no way to learn a per-file
 * BOM/newline convention without a one-line change to
 * `useSpreadsheetEditor.ts`'s `writeToDisk`, which is out of this fix's
 * ownership (`src/viewers/spreadsheet/**` other than that one function is
 * explicitly do-not-touch). The CSV test below documents this honestly: it
 * asserts the delimiter fix (real) and the BOM-still-stripped/LF-still-CRLF
 * gap (real, unfixed, characterized rather than silently ignored) in the
 * same test, rather than a passing test with a false "byte-faithful"
 * implication. See this fix's report for the exact diff that would close
 * the gap.
 */

const projectRoot = process.cwd()

function tempFile(name: string, bytes: Buffer): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-text-roundtrip-'))
  const file = path.join(dir, name)
  fs.writeFileSync(file, bytes)
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const page = await app.firstWindow()
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

test('SHELL-1/SHELL-2: a CRLF markdown file with a UTF-8 BOM keeps its BOM and CRLF on save, including untouched lines', async () => {
  const original = Buffer.concat([
    Buffer.from([0xef, 0xbb, 0xbf]),
    Buffer.from('# Title\r\n\r\nLine one.\r\nLine two.\r\nLine three.\r\n', 'utf-8'),
  ])
  const file = tempFile('crlf-bom.md', original)
  const { app, page } = await launch(file)
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    await page.keyboard.press('Control+2')
    const textarea = page.locator('.editor-panel__textarea')
    await expect(textarea).toBeVisible({ timeout: 10_000 })

    await textarea.click()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('Extra line appended.')

    await page.keyboard.press('Control+s')
    await expect
      .poll(() => fs.readFileSync(file).toString('utf-8'), { timeout: 10_000 })
      .toContain('Extra line appended.')

    const raw = fs.readFileSync(file)
    expect(raw.subarray(0, 3), 'BOM must survive the save').toEqual(Buffer.from([0xef, 0xbb, 0xbf]))
    expect(raw.toString('utf-8')).toBe(
      '﻿# Title\r\n\r\nLine one.\r\nLine two.\r\nLine three.\r\nExtra line appended.',
    )
  } finally {
    kill(app)
  }
})

test('a UTF-16LE code file keeps its encoding and BOM on save', async () => {
  const source = 'const greeting = "hello"\r\n'
  const original = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(source, 'utf16le')])
  const file = tempFile('utf16le.js', original)
  const { app, page } = await launch(file)
  try {
    await page.waitForSelector('.cm-content', { timeout: 20_000 })
    await page.waitForTimeout(400)
    await expect(page.locator('.cm-content')).toContainText('greeting')

    await page.locator('.cm-content').click()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('const answer = 42\n')

    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect
      .poll(() => fs.readFileSync(file).includes(Buffer.from('answer', 'utf16le')), { timeout: 10_000 })
      .toBe(true)

    const raw = fs.readFileSync(file)
    expect(raw.subarray(0, 2), 'UTF-16LE BOM must survive the save').toEqual(Buffer.from([0xff, 0xfe]))
    const decoded = raw.subarray(2).toString('utf16le')
    // The untouched first line's CRLF must survive; the exact whitespace
    // CodeMirror inserts around the typed Enter is not the point of this
    // test (encoding/BOM round-trip is), so this checks the substance
    // (CRLF preserved, both lines present) rather than a byte-exact match.
    expect(decoded.startsWith('const greeting = "hello"\r\nconst answer = 42')).toBe(true)
  } finally {
    kill(app)
  }
})

test('a Windows-1252 code file (invalid UTF-8, no BOM) re-encodes back to cp1252, not UTF-8', async () => {
  // `// café` with the accented `é` as the single cp1252 byte 0xE9 (invalid
  // UTF-8 on its own — the exact shape that forces the Windows-1252
  // fallback in `decodeTextBufferWithMeta`).
  const original = Buffer.from([0x2f, 0x2f, 0x20, 0x63, 0x61, 0x66, 0xe9, 0x0a])
  const file = tempFile('cp1252.js', original)
  const { app, page } = await launch(file)
  try {
    await page.waitForSelector('.cm-content', { timeout: 20_000 })
    await page.waitForTimeout(400)
    await expect(page.locator('.cm-content')).toContainText('café')

    await page.locator('.cm-content').click()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('const x = 1\n')

    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect
      .poll(() => fs.readFileSync(file).includes(Buffer.from('const x = 1')), { timeout: 10_000 })
      .toBe(true)

    const raw = fs.readFileSync(file)
    // Pre-fix behavior would have re-encoded as UTF-8, where é is the TWO
    // bytes 0xC3 0xA9 — assert the cp1252 single-byte form survives instead.
    expect(raw.includes(Buffer.from([0xc3, 0xa9])), 'must not re-encode é as UTF-8').toBe(false)
    expect(raw.includes(Buffer.from([0xe9]))).toBe(true)
    expect(raw.toString('latin1')).toBe('// café\nconst x = 1\n')
  } finally {
    kill(app)
  }
})

const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const ROW_MARKER_WIDTH = 48
const COLUMN_WIDTH = 120

async function launchCsv(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const page = await app.firstWindow()
  await page.waitForSelector('.csv-viewer__grid canvas', { timeout: 20_000 })
  await page.waitForTimeout(600)
  return { app, page }
}

/** Edits the first data cell (row 0, col 0) via the grid's overlay editor. */
async function editFirstCell(page: Page, text: string): Promise<void> {
  const box = (await page.locator('.csv-viewer__grid canvas').first().boundingBox())!
  await page.mouse.click(box.x + ROW_MARKER_WIDTH + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT / 2)
  await page.keyboard.type(text.slice(0, 1))
  await expect(page.locator('#portal textarea')).toBeFocused()
  await page.keyboard.type(text.slice(1), { delay: 20 })
  await page.keyboard.press('Enter')
  await expect(page.locator('#portal textarea')).toHaveCount(0)
}

test('SHEET-6: a semicolon-delimited CSV with a UTF-8 BOM keeps its semicolon delimiter on save (BOM/newline are a documented, unfixed gap)', async () => {
  const original = Buffer.concat([
    Buffer.from([0xef, 0xbb, 0xbf]),
    Buffer.from('Name;Amount\r\nAlice;1,5\r\nBob;2,0\r\n', 'utf-8'),
  ])
  const file = tempFile('semicolon-bom.csv', original)
  const { app, page } = await launchCsv(file)
  try {
    await editFirstCell(page, 'EDITED')
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.readFileSync(file, 'utf-8'), { timeout: 10_000 }).toContain('EDITED')

    const savedText = fs.readFileSync(file, 'utf-8')
    // Fixed (SHEET-6): the delimiter is still `;`, not silently rewritten to `,`.
    expect(savedText).toContain(';')
    expect(savedText.replace(/EDITED/g, '')).not.toContain('Alice,')
    const firstLine = savedText.replace(/^﻿/, '').split(/\r?\n/)[0]
    expect(firstLine.split(';').length).toBeGreaterThan(1)

    // Documented, unfixed gap (see module header): BOM/newline are not
    // reachable from `documentToDelimitedText` without an unowned call-site
    // change, so today they still fall back to "no BOM" / Papa's own
    // default newline rather than the source file's actual convention.
    const raw = fs.readFileSync(file)
    expect(raw.subarray(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf])), 'BOM preservation is NOT fixed yet').toBe(false)
  } finally {
    kill(app)
  }
})

test('a comma-delimited, LF-only CSV keeps its delimiter and content on save (newline preservation is the same documented, unfixed gap)', async () => {
  const original = Buffer.from('Name,Amount\nAlice,10\nBob,20\n', 'utf-8')
  const file = tempFile('lf-comma.csv', original)
  const { app, page } = await launchCsv(file)
  try {
    await editFirstCell(page, 'EDITED')
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.readFileSync(file, 'utf-8'), { timeout: 10_000 }).toContain('EDITED')

    const savedText = fs.readFileSync(file, 'utf-8')
    expect(savedText).toContain(',')
    expect(savedText).toContain('EDITED')
    expect(savedText).toContain('Bob')
  } finally {
    kill(app)
  }
})
