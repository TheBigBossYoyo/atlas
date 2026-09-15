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
