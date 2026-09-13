import { describe, expect, it } from 'vitest'

import { IMAGE_RELATIONSHIP_TYPE, imageMimeType, resolveImagePartPath } from '../mediaParts'

describe('resolveImagePartPath', () => {
  it('resolves a document-relative media target under word/', () => {
    expect(resolveImagePartPath({ id: 'rId1', type: IMAGE_RELATIONSHIP_TYPE, target: 'media/image1.png' })).toBe(
      'word/media/image1.png',
    )
  })

  it('keeps targets that are already package paths', () => {
    expect(resolveImagePartPath({ id: 'rId1', type: IMAGE_RELATIONSHIP_TYPE, target: 'word/media/image1.png' })).toBe(
      'word/media/image1.png',
    )
  })

  it('treats a leading slash as package-absolute instead of doubling the prefix', () => {
    expect(resolveImagePartPath({ id: 'rId1', type: IMAGE_RELATIONSHIP_TYPE, target: '/word/media/image1.png' })).toBe(
      'word/media/image1.png',
    )
  })

  it('normalizes parent-directory segments', () => {
    expect(resolveImagePartPath({ id: 'rId1', type: IMAGE_RELATIONSHIP_TYPE, target: '../customMedia/pic.jpeg' })).toBe(
      'customMedia/pic.jpeg',
    )
  })

  it('ignores external (linked) images', () => {
    expect(
      resolveImagePartPath({
        id: 'rId1',
        type: IMAGE_RELATIONSHIP_TYPE,
        target: 'https://example.com/pic.png',
        targetMode: 'External',
      }),
    ).toBeNull()
  })

  it('ignores non-image relationships', () => {
    expect(
      resolveImagePartPath({
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
        target: 'header1.xml',
      }),
    ).toBeNull()
  })
})

describe('imageMimeType', () => {
  it.each([
    ['word/media/a.png', 'image/png'],
    ['word/media/a.JPG', 'image/jpeg'],
    ['word/media/a.jpeg', 'image/jpeg'],
    ['word/media/a.gif', 'image/gif'],
    ['word/media/a.bmp', 'image/bmp'],
    ['word/media/a.webp', 'image/webp'],
    ['word/media/a.svg', 'image/svg+xml'],
  ])('maps %s to %s', (path, mime) => {
    expect(imageMimeType(path)).toBe(mime)
  })

  it.each(['word/media/a.emf', 'word/media/a.wmf', 'word/media/a.tiff', 'word/media/noext'])(
    'returns null for formats browsers cannot paint (%s)',
    (path) => {
      expect(imageMimeType(path)).toBeNull()
    },
  )
})
