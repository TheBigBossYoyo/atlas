import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'

import { buildEditableDeck } from '../../src/viewers/slides/pptx/editing/__tests__/editableDeck'

/** USR-15/USR-16 — the PowerPoint editor driven through the real Electron app. */

const projectRoot = process.cwd()
const main = '.slide-deck__main'

async function createDeck(): Promise<string> {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-pptx-editor-')), 'deck.pptx')
  fs.writeFileSync(file, Buffer.from(await buildEditableDeck()))
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  await page.waitForSelector(`${main} .slide-edit__layer`, { timeout: 20_000 })
  await page.waitForTimeout(500)
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

const shapeText = (page: Page, text: string) => page.locator(`${main} .slide-shape--text`, { hasText: text }).first()

async function saveAndRead(page: Page, file: string): Promise<JSZip> {
  const before = fs.statSync(file).mtimeMs
  await page.getByRole('button', { name: 'Save', exact: true }).click()
  await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
  await page.waitForTimeout(300)
  return JSZip.loadAsync(fs.readFileSync(file))
}

test('USR-15: slide text is not clipped by its box', async () => {
  const { app, page } = await launch(await createDeck())
  try {
    const title = shapeText(page, 'Quarterly review')
    await expect(title).toBeVisible()
    await expect(title).toHaveCSS('overflow', 'visible')
  } finally {
    kill(app)
  }
})

test('USR-16: edit text in place, move a shape, and save to .pptx', async () => {
  const file = await createDeck()
  const { app, page } = await launch(file)
  try {
    // The editing layer sits above the shapes (as in PowerPoint), so use real pointer coordinates.
    const title = (await shapeText(page, 'Quarterly review').boundingBox())!
    await page.mouse.dblclick(title.x + title.width / 2, title.y + title.height / 2)
    const editor = page.getByRole('textbox', { name: 'Edit slide text' })
    await expect(editor).toBeFocused()
    await page.keyboard.press('Control+A')
    await page.keyboard.type('Annual review')
    await page.keyboard.press('Escape')
    await expect(shapeText(page, 'Annual review')).toBeVisible()

    const draft = shapeText(page, 'Draft')
    const box = (await draft.boundingBox())!
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
    await page.mouse.down()
    await page.mouse.move(box.x + box.width / 2 + 120, box.y + box.height / 2 - 60, { steps: 6 })
    await page.mouse.up()
    await expect.poll(async () => (await shapeText(page, 'Draft').boundingBox())!.x).toBeGreaterThan(box.x + 60)

    const zip = await saveAndRead(page, file)
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(slide).toContain('<a:t>Annual review</a:t>')
    expect(slide).not.toContain('<a:off x="190500" y="5715000"/>')
    await expect(page.getByRole('button', { name: 'Undo' })).toBeEnabled()
  } finally {
    kill(app)
  }
})

test('USR-16: add, delete and reorder slides, edit notes, presenter view', async () => {
  const file = await createDeck()
  const { app, page } = await launch(file)
  try {
    const status = page.getByText(/^Slide \d+ \/ \d+$/).last()
    await page.getByRole('button', { name: 'New slide' }).click()
    await expect(status).toHaveText('Slide 2 / 3')
    await expect(page.locator(main).getByText('Click to add title')).toBeVisible()

    await page.getByRole('button', { name: 'Delete slide' }).click()
    await expect(status).toHaveText(/\/ 2$/)

    await page.getByRole('button', { name: 'Speaker notes' }).click()
    const notes = page.getByRole('textbox', { name: 'Speaker notes' })
    await notes.fill('Remember the demo')
    await page.getByRole('button', { name: 'Move slide up' }).click()

    const zip = await saveAndRead(page, file)
    const presentation = await zip.file('ppt/presentation.xml')!.async('string')
    expect(presentation.indexOf('r:id="rId3"')).toBeLessThan(presentation.indexOf('r:id="rId2"'))
    const notesParts = Object.keys(zip.files).filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name))
    const notesXml = await Promise.all(notesParts.map((name) => zip.file(name)!.async('string')))
    expect(notesXml.some((xml) => xml.includes('Remember the demo'))).toBe(true)

    await page.getByRole('button', { name: 'Presenter view' }).click()
    const presenter = page.getByRole('dialog', { name: 'Presenter view' })
    await expect(presenter).toBeVisible()
    await expect(presenter.getByText(/Slide \d \/ 2/)).toBeVisible()
    await presenter.getByRole('button', { name: 'Exit presenter view' }).click()
    await expect(presenter).toHaveCount(0)
  } finally {
    kill(app)
  }
})

test('USR-16: an ODP presentation is editable too', async () => {
  const source = path.join(projectRoot, 'tests', 'e2e', 'fixtures', 'sample.odp')
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-odp-editor-')), 'deck.odp')
  fs.copyFileSync(source, file)

  const { app, page } = await launch(file)
  try {
    const title = page.locator(`${main} .slide-shape--text`).first()
    const box = (await title.boundingBox())!
    await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2)
    const editor = page.getByRole('textbox', { name: 'Edit slide text' })
    await expect(editor).toBeFocused()
    await page.keyboard.press('Control+A')
    await page.keyboard.type('Edited in Atlas')
    await page.keyboard.press('Escape')
    await expect(shapeText(page, 'Edited in Atlas')).toBeVisible()

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(300)

    const zip = await JSZip.loadAsync(fs.readFileSync(file))
    expect(Object.keys(zip.files)[0]).toBe('mimetype')
    expect(await zip.file('content.xml')!.async('string')).toContain('Edited in Atlas')
  } finally {
    kill(app)
  }
})

/**
 * USR-16 — "Insert Text Box" then typing right away used to lose every
 * keystroke and could leave a duplicate, empty shape on the slide. Root
 * causes (both fixed):
 *  1. The new shape's id resolved before the re-parsed slide containing it
 *     had reached `SlideDeck`'s `slides` prop, so the same-tick lookup that
 *     used to select/open it always failed silently — nothing was ever
 *     opened for editing, so keystrokes had nowhere to land.
 *  2. The toolbar button keeps focus after the click. "Hello world"
 *     contains a space, which the browser turns into a second native click
 *     on a still-focused button, firing a second insert.
 * Typed right after the click, exactly as the button leaves the user
 * positioned — no click on the new shape to "rescue" focus.
 */
test('USR-16: insert a text box and type immediately — no rescuing click, exactly one shape, text survives save/reopen', async () => {
  const file = await createDeck()
  const { app, page } = await launch(file)
  try {
    const before = await page.locator(`${main} .slide-shape--text`).count()

    await page.getByRole('button', { name: 'Insert text box' }).click()
    await page.keyboard.type('Hello world', { delay: 30 })
    await page.keyboard.press('Escape')

    // Exactly one new shape: not zero (the text-lost bug) and not two (the
    // Space in "Hello world" duplicating the insert).
    await expect(page.locator(`${main} .slide-shape--text`)).toHaveCount(before + 1)
    await expect(shapeText(page, 'Hello world')).toBeVisible({ timeout: 10_000 })

    const zip = await saveAndRead(page, file)
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const matches = slide.match(/<a:t>Hello world<\/a:t>/g) ?? []
    expect(matches).toHaveLength(1)
  } finally {
    kill(app)
  }

  const validation = execSync(`node scripts/validate-office-file.mjs "${file}"`, { cwd: projectRoot, encoding: 'utf8' })
  expect(validation).not.toContain('error')

  // Reopen: relaunch the real app against the saved file and confirm the
  // typed text is there — not just present in the raw XML, but rendered.
  const reopened = await launch(file)
  try {
    await expect(shapeText(reopened.page, 'Hello world')).toBeVisible()
  } finally {
    kill(reopened.app)
  }
})

test('USR-16: an ODP text box, too, is typeable immediately after Insert Text Box with no rescuing click', async () => {
  const source = path.join(projectRoot, 'tests', 'e2e', 'fixtures', 'sample.odp')
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-odp-insert-textbox-')), 'deck.odp')
  fs.copyFileSync(source, file)

  const { app, page } = await launch(file)
  try {
    const before = await page.locator(`${main} .slide-shape--text`).count()

    await page.getByRole('button', { name: 'Insert text box' }).click()
    await page.keyboard.type('Hello world', { delay: 30 })
    await page.keyboard.press('Escape')

    await expect(page.locator(`${main} .slide-shape--text`)).toHaveCount(before + 1)
    await expect(shapeText(page, 'Hello world')).toBeVisible({ timeout: 10_000 })

    const before2 = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before2)
    await page.waitForTimeout(300)

    const zip = await JSZip.loadAsync(fs.readFileSync(file))
    const contentXml = await zip.file('content.xml')!.async('string')
    const matches = contentXml.match(/Hello world/g) ?? []
    expect(matches).toHaveLength(1)
  } finally {
    kill(app)
  }

  const validation = execSync(`node scripts/validate-office-file.mjs "${file}"`, { cwd: projectRoot, encoding: 'utf8' })
  expect(validation).not.toContain('error')
})
