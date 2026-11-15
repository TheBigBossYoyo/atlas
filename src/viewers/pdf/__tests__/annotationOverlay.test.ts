import { afterEach, describe, expect, it, vi } from 'vitest'

import { renderAnnotationOverlay } from '../annotationOverlay'
import type { RawPdfAnnotation } from '../annotations'
import type { PdfDocument } from '../types'

/** A pdf.js-shaped annotation storage: a plain map with the two methods the overlay calls. */
function fakeDoc(fieldObjects: Record<string, Array<{ id: string }>> = {}) {
  const values = new Map<string, Record<string, unknown>>()
  const doc = {
    annotationStorage: {
      getValue: (key: string, fallback: object) => values.get(key) ?? fallback,
      setValue: (key: string, value: Record<string, unknown>) => values.set(key, { ...values.get(key), ...value }),
    },
    getFieldObjects: () => Promise.resolve(fieldObjects),
  }
  return { doc: doc as unknown as PdfDocument, values }
}

const viewport = {
  scale: 1,
  convertToViewportPoint: (x: number, y: number) => [x, 1000 - y],
} as unknown as Parameters<typeof renderAnnotationOverlay>[2]

const textField: RawPdfAnnotation = {
  id: 'name',
  subtype: 'Widget',
  fieldType: 'Tx',
  fieldName: 'name',
  fieldValue: 'from file',
  rect: [0, 0, 100, 20],
}

const callbacks = () => ({ onInternalNavigate: vi.fn(), onFormChange: vi.fn() })

afterEach(() => {
  document.body.replaceChildren()
})

describe('renderAnnotationOverlay — form fields', () => {
  it('rebuilds a field from what the user typed, not from the file, after its page re-renders', () => {
    const { doc } = fakeDoc()
    const container = document.createElement('div')
    const cb = callbacks()
    renderAnnotationOverlay(container, [textField], viewport, doc, cb)
    const input = container.querySelector('input') as HTMLInputElement
    expect(input.value).toBe('from file')

    input.value = 'typed by the user'
    input.dispatchEvent(new Event('input'))
    expect(cb.onFormChange).toHaveBeenCalledTimes(1)

    // The page scrolled out of the render window and back.
    renderAnnotationOverlay(container, [textField], viewport, doc, cb)
    expect((container.querySelector('input') as HTMLInputElement).value).toBe('typed by the user')
  })

  it('stores checkbox state as a boolean — the type pdf.js saves and prints from', () => {
    const { doc, values } = fakeDoc()
    // Attached, as the real overlay is: jsdom fires no change event for a detached input.
    const container = document.body.appendChild(document.createElement('div'))
    renderAnnotationOverlay(
      container,
      [{ id: 'agree', subtype: 'Widget', fieldType: 'Btn', checkBox: true, fieldValue: 'Yes', exportValue: 'Yes', rect: [0, 0, 10, 10] }],
      viewport,
      doc,
      callbacks(),
    )
    const box = container.querySelector('input') as HTMLInputElement
    expect(box.checked).toBe(true)
    box.click()
    expect(values.get('agree')).toEqual({ value: false })
  })

  it('turns the other buttons of a radio group off in storage, including ones on other pages', async () => {
    const { doc, values } = fakeDoc({ size: [{ id: 'small' }, { id: 'large' }, { id: 'elsewhere' }] })
    const page = document.createElement('div')
    document.body.appendChild(page)
    const radio = (id: string, rect: [number, number, number, number]): RawPdfAnnotation => ({
      id,
      subtype: 'Widget',
      fieldType: 'Btn',
      radioButton: true,
      fieldName: 'size',
      buttonValue: id,
      fieldValue: 'small',
      rect,
    })
    renderAnnotationOverlay(page, [radio('small', [0, 0, 10, 10]), radio('large', [20, 0, 30, 10])], viewport, doc, callbacks())
    const [small, large] = Array.from(page.querySelectorAll('input'))
    expect(small.checked).toBe(true)

    large.click()
    await vi.waitFor(() => expect(values.get('elsewhere')).toEqual({ value: false }))
    expect(values.get('large')).toEqual({ value: true })
    expect(values.get('small')).toEqual({ value: false })
    expect(small.checked).toBe(false)
  })

  it('renders multi-line, read-only and length-limited text fields', () => {
    const { doc } = fakeDoc()
    const container = document.createElement('div')
    renderAnnotationOverlay(
      container,
      [{ ...textField, id: 'notes', multiLine: true, readOnly: true, maxLen: 12 }],
      viewport,
      doc,
      callbacks(),
    )
    const area = container.querySelector('textarea') as HTMLTextAreaElement
    expect(area.readOnly).toBe(true)
    expect(area.maxLength).toBe(12)
  })

  it('fills a choice field and stores the chosen export value', () => {
    const { doc, values } = fakeDoc()
    const container = document.createElement('div')
    const cb = callbacks()
    renderAnnotationOverlay(
      container,
      [
        {
          id: 'country',
          subtype: 'Widget',
          fieldType: 'Ch',
          combo: true,
          fieldValue: ['FR'],
          options: [
            { exportValue: 'FR', displayValue: 'France' },
            { exportValue: 'LB', displayValue: 'Lebanon' },
          ],
          rect: [0, 0, 100, 20],
        },
      ],
      viewport,
      doc,
      cb,
    )
    const select = container.querySelector('select') as HTMLSelectElement
    expect(select.value).toBe('FR')
    select.value = 'LB'
    select.dispatchEvent(new Event('change'))
    expect(values.get('country')).toEqual({ value: 'LB' })
    expect(cb.onFormChange).toHaveBeenCalled()
  })
})
