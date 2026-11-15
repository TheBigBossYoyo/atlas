// @vitest-environment node
/**
 * Round-trips the edit writer through real pdf-lib (to build and re-read the
 * file) and real pdf.js (the engine Atlas itself reopens the result with), so
 * these assert on what a reader actually finds in the saved bytes — not on
 * what the writer meant to put there.
 */
import { afterEach, describe, expect, it } from 'vitest'
import * as pdfLib from 'pdf-lib'
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist/legacy/build/pdf.mjs'

import { createInitialEditState, deletePage, movePage, rotatePage } from '../editState'
import { applyEditsToPdf, buildEditedPdf, UnsupportedCharacterError } from '../writePdf'
import type { PdfEditState } from '../types'

GlobalWorkerOptions.workerSrc = new URL(
  '../../../../../node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs',
  import.meta.url,
).href

const BLACK = { r: 0, g: 0, b: 0 }
const YELLOW = { r: 1, g: 0.9, b: 0.2 }

/** A small document whose pages are told apart by size and by their own text. */
async function makePdf(options: { readonly withForm?: boolean } = {}): Promise<Uint8Array> {
  const doc = await pdfLib.PDFDocument.create()
  const font = await doc.embedFont(pdfLib.StandardFonts.Helvetica)
  const sizes: Array<[number, number]> = [[600, 800], [601, 801], [602, 802]]
  sizes.forEach(([width, height], index) => {
    const page = doc.addPage([width, height])
    page.drawText(`Page ${index + 1} original`, { x: 72, y: 700, size: 12, font })
  })
  if (options.withForm) {
    const form = doc.getForm()
    form.createTextField('onFirst').addToPage(doc.getPage(0), { x: 72, y: 600, width: 200, height: 20 })
    form.createTextField('onSecond').addToPage(doc.getPage(1), { x: 72, y: 600, width: 200, height: 20 })
  }
  return doc.save()
}

const openTasks: Array<ReturnType<typeof getDocument>> = []

/** pdfjs-dist 6 tears a document down through its loading task, so keep each one for `afterEach`. */
async function openWithPdfjs(bytes: Uint8Array) {
  const task = getDocument({ data: bytes.slice(0) })
  openTasks.push(task)
  return task.promise
}

afterEach(async () => {
  await Promise.all(openTasks.splice(0).map((task) => task.destroy()))
})

async function pageText(doc: Awaited<ReturnType<typeof openWithPdfjs>>, pageNumber: number): Promise<string> {
  const page = await doc.getPage(pageNumber)
  const content = await page.getTextContent()
  return content.items.map((item) => ('str' in item ? item.str : '')).join(' ')
}

function withState(state: PdfEditState, patch: Partial<PdfEditState>): PdfEditState {
  return { ...state, ...patch }
}

describe('applyEditsToPdf', () => {
  it('returns the input bytes untouched when nothing was edited', async () => {
    const original = await makePdf()
    const result = await applyEditsToPdf(pdfLib, original, createInitialEditState(3), 3)
    expect(result).toBe(original)
  })

  it('reorders, deletes and rotates pages, and the result reads back that way', async () => {
    const original = await makePdf()
    let state = createInitialEditState(3)
    state = movePage(state, 2, -2) // 3,1,2
    state = deletePage(state, 2) // 3,1
    state = rotatePage(state, 1, 90) // page "1" turned a quarter
    const result = await applyEditsToPdf(pdfLib, original, state, 3)

    const reread = await pdfLib.PDFDocument.load(result)
    const pages = reread.getPages()
    expect(pages.map((page) => page.getWidth())).toEqual([602, 600])
    expect(pages.map((page) => page.getRotation().angle)).toEqual([0, 90])

    const pdf = await openWithPdfjs(result)
    expect(pdf.numPages).toBe(2)
    expect(await pageText(pdf, 1)).toContain('Page 3 original')
    expect(await pageText(pdf, 2)).toContain('Page 1 original')
  })

  it('adds a turn on top of a page that already had its own /Rotate', async () => {
    const doc = await pdfLib.PDFDocument.create()
    doc.addPage([600, 800]).setRotation(pdfLib.degrees(270))
    const original = await doc.save()
    const state = rotatePage(rotatePage(createInitialEditState(1), 0, 90), 0, 90)
    const result = await applyEditsToPdf(pdfLib, original, state, 1)
    const reread = await pdfLib.PDFDocument.load(result)
    expect(reread.getPage(0).getRotation().angle).toBe(90)
  })

  it('writes highlight, ink and free-text annotations pdf.js recognises', async () => {
    const original = await makePdf()
    const state = withState(createInitialEditState(3), {
      annotations: [
        { kind: 'highlight', id: 'h1', page: 1, rects: [{ x: 70, y: 695, width: 120, height: 16 }], color: YELLOW },
        {
          kind: 'ink',
          id: 'i1',
          page: 1,
          points: [{ x: 100, y: 100 }, { x: 150, y: 140 }, { x: 200, y: 100 }],
          color: { r: 1, g: 0, b: 0 },
          width: 2,
        },
        {
          kind: 'text',
          id: 't1',
          page: 2,
          anchor: { x: 72, y: 500 },
          rotation: 0,
          text: 'Hello\nworld',
          fontSize: 14,
          color: BLACK,
        },
      ],
    })
    const pdf = await openWithPdfjs(await applyEditsToPdf(pdfLib, original, state, 3))

    const first = await (await pdf.getPage(1)).getAnnotations()
    expect(first.map((annotation) => annotation.subtype).sort()).toEqual(['Highlight', 'Ink'])
    const highlight = first.find((annotation) => annotation.subtype === 'Highlight')
    expect(highlight.hasAppearance).toBe(true)

    const second = await (await pdf.getPage(2)).getAnnotations()
    expect(second).toHaveLength(1)
    expect(second[0].subtype).toBe('FreeText')
    expect(second[0].hasAppearance).toBe(true)
    expect(second[0].contentsObj.str).toBe('Hello\nworld')
  })

  it('places a rotated text box so its bounding rect matches the turned box', async () => {
    const original = await makePdf()
    const state = withState(createInitialEditState(3), {
      annotations: [
        { kind: 'text', id: 't', page: 1, anchor: { x: 300, y: 300 }, rotation: 90, text: 'Up', fontSize: 10, color: BLACK },
      ],
    })
    const pdf = await openWithPdfjs(await applyEditsToPdf(pdfLib, original, state, 3))
    const [annotation] = await (await pdf.getPage(1)).getAnnotations()
    const [x1, y1, x2, y2] = annotation.rect as number[]
    // Turned 90° counter-clockwise about its top-left corner, the box extends
    // up from the anchor (its width) and right of it (its height).
    expect(x1).toBeCloseTo(300, 3)
    expect(y1).toBeCloseTo(300, 3)
    expect(x2 - x1).toBeGreaterThan(10)
    expect(y2 - y1).toBeGreaterThan(10)
  })

  it('refuses characters the standard fonts cannot draw, with a typed error', async () => {
    const original = await makePdf()
    const state = withState(createInitialEditState(3), {
      annotations: [
        { kind: 'text', id: 't', page: 1, anchor: { x: 72, y: 500 }, rotation: 0, text: 'Привет', fontSize: 12, color: BLACK },
      ],
    })
    await expect(applyEditsToPdf(pdfLib, original, state, 3)).rejects.toBeInstanceOf(UnsupportedCharacterError)
  })

  it('removes a deleted page from the file itself, not just from the page list', async () => {
    const source = await pdfLib.PDFDocument.load(await makePdf())
    // Something that still points at page 2 after it is deleted: a named destination.
    const doomed = source.getPage(1).ref
    source.catalog.set(pdfLib.PDFName.of('Dests'), source.context.obj({ Chapter2: [doomed, 'Fit'] }))
    const original = await source.save({ useObjectStreams: false })

    const result = await applyEditsToPdf(pdfLib, original, deletePage(createInitialEditState(3), 1), 3)

    const reread = await pdfLib.PDFDocument.load(result)
    const pageObjects = reread.context
      .enumerateIndirectObjects()
      .filter(([, object]) => object instanceof pdfLib.PDFDict && object.get(pdfLib.PDFName.of('Type')) === pdfLib.PDFName.of('Page'))
    expect(pageObjects).toHaveLength(2)
    const dests = reread.catalog.lookup(pdfLib.PDFName.of('Dests'), pdfLib.PDFDict)
    expect(dests.lookup(pdfLib.PDFName.of('Chapter2'), pdfLib.PDFArray).get(0)).toBe(pdfLib.PDFNull)
    expect(await pageText(await openWithPdfjs(result), 2)).toContain('Page 3 original')
  })

  it('removes form fields whose only widget was on a deleted page', async () => {
    const original = await makePdf({ withForm: true })
    const state = deletePage(createInitialEditState(3), 1)
    const reread = await pdfLib.PDFDocument.load(await applyEditsToPdf(pdfLib, original, state, 3))
    expect(reread.getForm().getFields().map((field) => field.getName())).toEqual(['onFirst'])
  })
})

describe('buildEditedPdf', () => {
  it('skips pdf-lib entirely for a forms-only save', async () => {
    const original = await makePdf()
    const formBytes = new Uint8Array([1, 2, 3])
    let loaded = false
    const result = await buildEditedPdf({
      original,
      saveForms: async () => formBytes,
      edits: createInitialEditState(3),
      sourcePageCount: 3,
      loadPdfLib: async () => {
        loaded = true
        return pdfLib
      },
    })
    expect(result).toBe(formBytes)
    expect(loaded).toBe(false)
  })

  it('applies page edits on top of the filled-in form bytes', async () => {
    const original = await makePdf({ withForm: true })
    const filled = await pdfLib.PDFDocument.load(original)
    filled.getForm().getTextField('onFirst').setText('typed value')
    const filledBytes = await filled.save()

    const result = await buildEditedPdf({
      original,
      saveForms: async () => filledBytes,
      edits: rotatePage(createInitialEditState(3), 0, 90),
      sourcePageCount: 3,
      loadPdfLib: async () => pdfLib,
    })
    const reread = await pdfLib.PDFDocument.load(result)
    expect(reread.getForm().getTextField('onFirst').getText()).toBe('typed value')
    expect(reread.getPage(0).getRotation().angle).toBe(90)
  })
})
