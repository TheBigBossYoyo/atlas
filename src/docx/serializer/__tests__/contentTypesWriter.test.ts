/**
 * Atlas — contentTypesWriter unit tests (Wave D.3)
 */

import { describe, expect, it } from 'vitest'
import {
  MEDIA_CONTENT_TYPES,
  addOverride,
  ensureMediaContentType,
  writeContentTypesXml,
} from '../contentTypesWriter'
import type { ContentTypesPart } from '../contentTypesWriter'

const empty: ContentTypesPart = { defaults: [], overrides: [] }

// ---------------------------------------------------------------------------
// writeContentTypesXml
// ---------------------------------------------------------------------------

describe('writeContentTypesXml', () => {
  it('empty types → XML decl + Types root', () => {
    const xml = writeContentTypesXml(empty)
    expect(xml).toContain('<?xml version="1.0"')
    expect(xml).toContain('Types')
    expect(xml).toContain('http://schemas.openxmlformats.org/package/2006/content-types')
  })

  it('Default entries are emitted before Override entries', () => {
    const types: ContentTypesPart = {
      defaults: [{ extension: 'rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' }],
      overrides: [{ partName: '/word/document.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }],
    }
    const xml = writeContentTypesXml(types)
    expect(xml.indexOf('Default')).toBeLessThan(xml.indexOf('Override'))
  })

  it('Default extension and contentType are present', () => {
    const types: ContentTypesPart = {
      defaults: [{ extension: 'png', contentType: 'image/png' }],
      overrides: [],
    }
    const xml = writeContentTypesXml(types)
    expect(xml).toContain('png')
    expect(xml).toContain('image/png')
  })
})

// ---------------------------------------------------------------------------
// ensureMediaContentType
// ---------------------------------------------------------------------------

describe('ensureMediaContentType', () => {
  it('adds png default entry when not present', () => {
    const result = ensureMediaContentType(empty, 'png')
    expect(result.defaults).toHaveLength(1)
    expect(result.defaults[0]?.extension).toBe('png')
    expect(result.defaults[0]?.contentType).toBe('image/png')
  })

  it('is idempotent — does not duplicate existing extension', () => {
    const withPng = ensureMediaContentType(empty, 'png')
    const again = ensureMediaContentType(withPng, 'png')
    expect(again.defaults).toHaveLength(1)
  })

  it('does not mutate original', () => {
    ensureMediaContentType(empty, 'png')
    expect(empty.defaults).toHaveLength(0)
  })

  it('unknown extension returns same shape unchanged', () => {
    const result = ensureMediaContentType(empty, 'xyz')
    expect(result.defaults).toHaveLength(0)
  })
})

// ---------------------------------------------------------------------------
// addOverride
// ---------------------------------------------------------------------------

describe('addOverride', () => {
  it('pure-adds override to the list', () => {
    const result = addOverride(empty, '/word/document.xml', 'application/some+xml')
    expect(result.overrides).toHaveLength(1)
    expect(result.overrides[0]?.partName).toBe('/word/document.xml')
  })

  it('does not mutate original', () => {
    addOverride(empty, '/word/document.xml', 'application/some+xml')
    expect(empty.overrides).toHaveLength(0)
  })
})

// ---------------------------------------------------------------------------
// MEDIA_CONTENT_TYPES
// ---------------------------------------------------------------------------

describe('MEDIA_CONTENT_TYPES', () => {
  it('has all 7 required entries', () => {
    const keys = Object.keys(MEDIA_CONTENT_TYPES)
    expect(keys).toContain('png')
    expect(keys).toContain('jpg')
    expect(keys).toContain('jpeg')
    expect(keys).toContain('gif')
    expect(keys).toContain('bmp')
    expect(keys).toContain('tiff')
    expect(keys).toContain('svg')
  })

  it('jpeg and jpg both map to image/jpeg', () => {
    expect(MEDIA_CONTENT_TYPES['jpeg']).toBe('image/jpeg')
    expect(MEDIA_CONTENT_TYPES['jpg']).toBe('image/jpeg')
  })
})
