/**
 * PDF-EDIT-1 — editing a PDF in the real app, then reading the saved file
 * back with pdf-lib and pdf.js (outside the app) to check what actually got
 * written: form values, a text box, a drawing, a highlight, a text
 * replacement, a page rotation and a page deletion.
 *
 * Before this, Atlas had no PDF save path at all: form fields could be typed
 * into but Ctrl+S did nothing, and the typed values vanished as soon as the
 * page scrolled out of view.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'
import { pathToFileURL } from 'node:url'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { PDFCheckBox, PDFDocument, PDFTextField, StandardFonts } from 'pdf-lib'
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist/legacy/build/pdf.mjs'

const projectRoot = process.cwd()
GlobalWorkerOptions.workerSrc = pathToFileURL(
  path.join(projectRoot, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs'),
).href

async function createFormPdf(): Promise<string> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (let index = 1; index <= 3; index += 1) {
    const page = doc.addPage([612, 792])
    page.drawText(`Page ${index} original`, { x: 72, y: 700, size: 18, font })
    page.drawText(`Highlight me on page ${index}`, { x: 72, y: 650, size: 14, font })
  }
  const form = doc.getForm()
  form.createTextField('name').addToPage(doc.getPage(0), { x: 72, y: 560, width: 220, height: 24 })
  form.createCheckBox('agree').addToPage(doc.getPage(0), { x: 72, y: 520, width: 16, height: 16 })
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-pdf-edit-'))
  const file = path.join(dir, 'form.pdf')
  fs.writeFileSync(file, await doc.save())
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page; errors: string[] }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.waitForSelector('.pdf-viewer__page-wrapper[data-page="1"] .pdf-viewer__text-layer span', { timeout: 20_000 })
  return { app, page, errors }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function pageBox(page: Page, pageNumber: number) {
  const box = await page.locator(`.pdf-viewer__page-wrapper[data-page="${pageNumber}"]`).boundingBox()
  if (!box) throw new Error(`page ${pageNumber} is not laid out`)
  return box
}

async function goToPage(page: Page, pageNumber: number): Promise<void> {
  const input = page.getByLabel('Page number')
  await input.fill(String(pageNumber))
  await input.press('Enter')
  await expect(input).toHaveValue(String(pageNumber))
  await page.waitForTimeout(700) // the smooth scroll settles before the current page is trusted
}

async function savedTextOf(bytes: Uint8Array): Promise<string[]> {
  const task = getDocument({ data: bytes.slice(0) })
  try {
    const pdf = await task.promise
    const texts: string[] = []
    for (let index = 1; index <= pdf.numPages; index += 1) {
      const content = await (await pdf.getPage(index)).getTextContent()
      texts.push(content.items.map((item) => ('str' in item ? item.str : '')).join(' '))
    }
    return texts
  } finally {
    await task.destroy()
  }
}

async function annotationSubtypes(bytes: Uint8Array): Promise<string[][]> {
  const task = getDocument({ data: bytes.slice(0) })
  try {
    const pdf = await task.promise
    const result: string[][] = []
    for (let index = 1; index <= pdf.numPages; index += 1) {
      const annotations = await (await pdf.getPage(index)).getAnnotations()
      result.push(annotations.map((annotation) => annotation.subtype as string).sort())
    }
    return result
  } finally {
    await task.destroy()
  }
}

let fixture = ''
test.beforeAll(async () => {
  fixture = await createFormPdf()
})

test('form values survive scrolling away and are written by Ctrl+S', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-pdf-form-')), 'form.pdf')
  fs.copyFileSync(fixture, file)
  const { app, page, errors } = await launch(file)
  try {
    const nameField = page.getByRole('textbox', { name: 'name' })
    await nameField.fill('Youssef')
    await page.getByRole('checkbox', { name: 'agree' }).check()

    // Page 1 leaves the render window and comes back: the overlay is rebuilt.
    await goToPage(page, 3)
    await page.locator('.pdf-viewer').evaluate((viewer) => viewer.scrollTo({ top: viewer.scrollHeight }))
    await page.waitForTimeout(300)
    await goToPage(page, 1)
    await expect(page.getByRole('textbox', { name: 'name' })).toHaveValue('Youssef')
    await expect(page.getByRole('checkbox', { name: 'agree' })).toBeChecked()

    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.readFileSync(file).length, { timeout: 10_000 }).not.toBe(fs.readFileSync(fixture).length)

    const form = (await PDFDocument.load(fs.readFileSync(file))).getForm()
    expect((form.getField('name') as PDFTextField).getText()).toBe('Youssef')
    expect((form.getField('agree') as PDFCheckBox).isChecked()).toBe(true)
    expect(errors, errors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})

test('annotations, a text replacement and page operations are saved into the file', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-pdf-edits-')), 'form.pdf')
  fs.copyFileSync(fixture, file)
  const { app, page, errors } = await launch(file)
  try {
    await page.getByRole('button', { name: 'Edit', exact: true }).click()
    const textTool = page.getByRole('radio', { name: 'Add a text box' })
    await expect(textTool).toBeEnabled({ timeout: 10_000 }) // pdf-lib loaded

    // A text box on page 1.
    await textTool.click()
    let box = await pageBox(page, 1)
    await page.mouse.click(box.x + box.width * 0.5, box.y + box.height * 0.45)
    const textEditor = page.getByRole('textbox', { name: 'Text box' })
    await expect(textEditor).toBeFocused()
    await page.keyboard.type('Hello from Atlas')
    await page.keyboard.press('Control+Enter')
    await expect(textEditor).toHaveCount(0)
    await expect(page.locator('.pdf-viewer__page-wrapper[data-page="1"] .pdf-edit-layer__svg text')).toHaveText(['Hello from Atlas'])

    // A drawing on page 1.
    await page.getByRole('radio', { name: 'Draw' }).click()
    box = await pageBox(page, 1)
    await page.mouse.move(box.x + 100, box.y + 450)
    await page.mouse.down()
    for (let step = 1; step <= 10; step += 1) await page.mouse.move(box.x + 100 + step * 15, box.y + 450 + (step % 2) * 20)
    await page.mouse.up()
    await expect(page.locator('.pdf-viewer__page-wrapper[data-page="1"] .pdf-edit-layer__svg polyline')).not.toHaveCount(0)

    // Rewrite the existing heading on page 1, in place.
    await page.getByRole('radio', { name: /^Edit text/ }).click()
    const heading = page.locator('.pdf-viewer__page-wrapper[data-page="1"] .pdf-viewer__text-layer span', { hasText: 'Page 1 original' })
    const headingBox = (await heading.boundingBox())!
    await page.mouse.click(headingBox.x + 4, headingBox.y + headingBox.height / 2)
    const paragraphEditor = page.getByRole('textbox', { name: 'Paragraph text' })
    await expect(paragraphEditor).toHaveText('Page 1 original')
    await paragraphEditor.evaluate((node) => {
      node.querySelector('[data-style]')!.textContent = 'Page 1 edited'
    })
    await page.keyboard.press('Control+Enter')
    await expect(paragraphEditor).toHaveCount(0)

    // Highlight a line on page 2 by selecting it.
    await page.getByRole('radio', { name: 'Highlight: select text to highlight it' }).click()
    await goToPage(page, 2)
    const line = page.locator('.pdf-viewer__page-wrapper[data-page="2"] .pdf-viewer__text-layer span', { hasText: 'Highlight me' })
    const lineBox = (await line.boundingBox())!
    await page.mouse.move(lineBox.x + 2, lineBox.y + lineBox.height / 2)
    await page.mouse.down()
    await page.mouse.move(lineBox.x + lineBox.width - 2, lineBox.y + lineBox.height / 2, { steps: 8 })
    await page.mouse.up()
    await expect(page.locator('.pdf-viewer__page-wrapper[data-page="2"] .pdf-edit-layer__highlights rect')).not.toHaveCount(0)

    // Delete page 3, then turn page 1 a quarter clockwise.
    await page.getByRole('radio', { name: 'Select, move and delete edits' }).click()
    await goToPage(page, 3)
    await page.getByRole('button', { name: 'Delete this page' }).click()
    await expect(page.locator('.pdf-viewer__page-wrapper')).toHaveCount(2)
    await goToPage(page, 1)
    await page.getByRole('button', { name: 'Rotate this page right (saved)' }).click()

    await page.keyboard.press('Control+s')
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeDisabled({ timeout: 10_000 })

    const saved = new Uint8Array(fs.readFileSync(file))
    const reread = await PDFDocument.load(saved)
    expect(reread.getPageCount()).toBe(2)
    expect(reread.getPage(0).getRotation().angle).toBe(90)

    const texts = await savedTextOf(saved)
    expect(texts[0]).toContain('Page 1 edited')
    // Rewritten, not painted over: the old words are gone from the file.
    expect(texts[0]).not.toContain('Page 1 original')
    expect(texts[1]).toContain('Page 2 original')
    expect(texts.join(' ')).not.toContain('Page 3 original')

    const subtypes = await annotationSubtypes(saved)
    expect(subtypes[0]).toEqual(expect.arrayContaining(['FreeText', 'Ink']))
    expect(subtypes[1]).toContain('Highlight')
    expect(errors, errors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})

test('Ctrl+S with a text box still being typed into saves that text box', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-pdf-flush-')), 'form.pdf')
  fs.copyFileSync(fixture, file)
  const { app, page, errors } = await launch(file)
  try {
    await page.getByRole('button', { name: 'Edit', exact: true }).click()
    const textTool = page.getByRole('radio', { name: 'Add a text box' })
    await expect(textTool).toBeEnabled({ timeout: 10_000 })
    await textTool.click()
    const box = await pageBox(page, 1)
    await page.mouse.click(box.x + box.width * 0.5, box.y + box.height * 0.45)
    await expect(page.getByRole('textbox', { name: 'Text box' })).toBeFocused()
    await page.keyboard.type('Typed then saved')
    // No Ctrl+Enter, no click away: the editor is still open.
    await page.keyboard.press('Control+s')

    await expect.poll(async () => (await annotationSubtypes(new Uint8Array(fs.readFileSync(file))))[0], { timeout: 10_000 }).toContain('FreeText')
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeDisabled()
    expect(errors, errors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})

test('undo takes an edit back and leaves the document clean', async () => {
  const { app, page, errors } = await launch(fixture)
  try {
    await page.getByRole('button', { name: 'Edit', exact: true }).click()
    const rotate = page.getByRole('button', { name: 'Rotate this page right (saved)' })
    await rotate.click()
    const save = page.getByRole('button', { name: 'Save', exact: true })
    await expect(save).toBeEnabled()
    await page.keyboard.press('Control+z')
    await expect(save).toBeDisabled()
    await page.keyboard.press('Control+y')
    await expect(save).toBeEnabled()
    expect(errors, errors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})
