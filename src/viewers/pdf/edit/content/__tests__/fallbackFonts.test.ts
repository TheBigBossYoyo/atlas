// @vitest-environment node
/**
 * Every bundled fallback font, embedded through Atlas's own subsetter, must
 * still have every glyph it was asked for. pdf-lib's subsetter silently lost
 * most glyphs of two of these families ("Added with Atlas" came out as "as"),
 * and nothing short of asking the renderer catches that: text extraction goes
 * through the ToUnicode map and reads fine either way.
 */
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import * as pdfLib from 'pdf-lib'
import { getDocument, GlobalWorkerOptions, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'

import { createInitialEditState } from '../../editState'
import { applyEditsToPdf } from '../../writePdf'
import type { FontStyle } from '../../types'
import { newTextEdit } from '../textEdits'
import { fallbackForTests } from './fallbackForTests'

const ROOT = path.resolve(__dirname, '..', '..', '..', '..', '..', '..')
GlobalWorkerOptions.workerSrc = pathToFileURL(path.join(ROOT, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs')).href

const tasks: Array<ReturnType<typeof getDocument>> = []
afterEach(async () => {
  await Promise.all(tasks.splice(0).map((task) => task.destroy()))
})

/** The glyphs pdf.js will draw, with `!` after any the embedded font does not contain. */
async function drawnGlyphs(bytes: Uint8Array): Promise<string> {
  const task = getDocument({ data: bytes.slice(0) })
  tasks.push(task)
  const page = await (await task.promise).getPage(1)
  const ops = await page.getOperatorList()
  let out = ''
  ops.fnArray.forEach((fn, index) => {
    if (fn !== OPS.showText) return
    for (const glyph of ops.argsArray[index][0] as Array<{ unicode?: string; isInFont?: boolean } | number>) {
      if (typeof glyph === 'number' || glyph.unicode === ' ') continue
      out += `${glyph.unicode}${glyph.isInFont ? '' : '!'}`
    }
  })
  return out
}

const STYLES: FontStyle[] = (['helvetica', 'times', 'courier'] as const).flatMap((family) =>
  [false, true].flatMap((bold) => [false, true].map((italic) => ({ family, bold, italic }))),
)

describe('bundled fallback fonts', () => {
  it.each(STYLES.map((style) => [`${style.family}${style.bold ? ' bold' : ''}${style.italic ? ' italic' : ''}`, style] as const))(
    '%s keeps every glyph of new text',
    async (_label, style) => {
      const doc = await pdfLib.PDFDocument.create()
      doc.addPage([400, 200])
      const blank = await doc.save()
      const text = 'Added with Atlas, Été — Жарко'
      const edit = newTextEdit('n', 1, { x: 20, y: 100 }, { x: 1, y: 0 }, text, style, 14, { r: 0, g: 0, b: 0 })
      const bytes = await applyEditsToPdf(pdfLib, blank, { ...createInitialEditState(1), textEdits: [edit] }, 1, fallbackForTests)
      expect(await drawnGlyphs(bytes)).toBe(text.replace(/ /g, ''))
    },
  )
})
