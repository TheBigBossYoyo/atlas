import { describe, expect, it } from 'vitest'

import {
  annotationRectToViewportBox,
  classifyAnnotation,
  isValidAnnotationRect,
  isWidgetChecked,
  type RawPdfAnnotation,
} from '../annotations'

describe('classifyAnnotation', () => {
  it('classifies a Link with a sanitized url as external-link', () => {
    expect(
      classifyAnnotation({ id: '1', subtype: 'Link', url: 'https://example.com' }),
    ).toBe('external-link')
  })

  it('classifies a Link with a dest (and no url) as internal-link', () => {
    expect(classifyAnnotation({ id: '1', subtype: 'Link', dest: 'chapter-1' })).toBe(
      'internal-link',
    )
    expect(classifyAnnotation({ id: '1', subtype: 'Link', dest: [0, 'Fit'] })).toBe(
      'internal-link',
    )
  })

  it('never trusts an unvalidated URL field — only pdf.js-sanitized `url` counts', () => {
    // pdf.js exposes the raw, unvalidated action URL as `unsafeUrl` even when
    // it rejected the URL as unsafe (e.g. a javascript: scheme) and left
    // `url` unset. RawPdfAnnotation deliberately has no `unsafeUrl` field —
    // this module must never treat one as clickable even if present at
    // runtime, so the cast below simulates that real pdf.js shape.
    const annotation = {
      id: '1',
      subtype: 'Link',
      unsafeUrl: 'javascript:alert(1)',
    } as unknown as RawPdfAnnotation
    expect(classifyAnnotation(annotation)).toBe('ignored')
  })

  it('ignores a Link annotation with neither url nor dest', () => {
    expect(classifyAnnotation({ id: '1', subtype: 'Link' })).toBe('ignored')
  })

  it('classifies Widget subtypes by fieldType', () => {
    expect(classifyAnnotation({ id: '1', subtype: 'Widget', fieldType: 'Tx' })).toBe(
      'text-field',
    )
    expect(
      classifyAnnotation({ id: '1', subtype: 'Widget', fieldType: 'Btn', checkBox: true }),
    ).toBe('checkbox')
    expect(
      classifyAnnotation({ id: '1', subtype: 'Widget', fieldType: 'Btn', radioButton: true }),
    ).toBe('radio')
  })

  it('ignores annotations flagged NoView or Hidden regardless of subtype', () => {
    expect(
      classifyAnnotation({ id: '1', subtype: 'Link', url: 'https://example.com', noView: true }),
    ).toBe('ignored')
    expect(
      classifyAnnotation({ id: '1', subtype: 'Link', url: 'https://example.com', hidden: true }),
    ).toBe('ignored')
  })

  it('ignores unsupported subtypes (e.g. free text, popup)', () => {
    expect(classifyAnnotation({ id: '1', subtype: 'Popup' })).toBe('ignored')
    expect(classifyAnnotation({ id: '1', subtype: 'Widget', fieldType: 'Ch' })).toBe('ignored')
  })
})

describe('isWidgetChecked', () => {
  it('a checkbox is checked when its value matches its export value', () => {
    expect(isWidgetChecked({ id: '1', checkBox: true, fieldValue: 'Yes', exportValue: 'Yes' })).toBe(
      true,
    )
    expect(isWidgetChecked({ id: '1', checkBox: true, fieldValue: 'Off', exportValue: 'Yes' })).toBe(
      false,
    )
  })

  it('a radio button is checked when its value matches its button value', () => {
    expect(
      isWidgetChecked({ id: '1', radioButton: true, fieldValue: 'A', buttonValue: 'A' }),
    ).toBe(true)
    expect(
      isWidgetChecked({ id: '1', radioButton: true, fieldValue: 'B', buttonValue: 'A' }),
    ).toBe(false)
  })

  it('is false for a non-button widget', () => {
    expect(isWidgetChecked({ id: '1' })).toBe(false)
  })
})

describe('isValidAnnotationRect', () => {
  it('accepts a 4-element array of finite numbers', () => {
    expect(isValidAnnotationRect([1, 2, 3, 4])).toBe(true)
  })

  it('rejects undefined, wrong length, or non-finite entries', () => {
    expect(isValidAnnotationRect(undefined)).toBe(false)
    expect(isValidAnnotationRect([1, 2, 3])).toBe(false)
    expect(isValidAnnotationRect([1, 2, 3, NaN])).toBe(false)
  })
})

describe('annotationRectToViewportBox', () => {
  it('converts both corners and returns their bounding box', () => {
    const viewport = {
      // Simulates a simple flip (as an unrotated PDF viewport does): y grows
      // downward in viewport space while PDF y grows upward.
      convertToViewportPoint: (x: number, y: number) => [x * 2, 100 - y * 2] as const,
    }

    const box = annotationRectToViewportBox(viewport, [10, 10, 30, 20])

    // (10,10) -> (20, 80); (30,20) -> (60, 60)
    expect(box).toEqual({ left: 20, top: 60, width: 40, height: 20 })
  })

  it('handles rect corners given in either order', () => {
    const viewport = { convertToViewportPoint: (x: number, y: number) => [x, y] as const }
    const box = annotationRectToViewportBox(viewport, [30, 20, 10, 5])
    expect(box).toEqual({ left: 10, top: 5, width: 20, height: 15 })
  })
})
