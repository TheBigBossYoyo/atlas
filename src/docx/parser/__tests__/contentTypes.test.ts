/**
 * Tests for src/docx/parser/contentTypes.ts
 */

import { describe, it, expect } from 'vitest'
import { parseContentTypes } from '../contentTypes'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const FULL_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"
    ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`

const SINGLE_DEFAULT_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`

const EMPTY_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
</Types>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseContentTypes', () => {
  it('parses the correct number of defaults', () => {
    const ct = parseContentTypes(FULL_CONTENT_TYPES)
    expect(ct.defaults).toHaveLength(2)
  })

  it('parses the correct number of overrides', () => {
    const ct = parseContentTypes(FULL_CONTENT_TYPES)
    expect(ct.overrides).toHaveLength(3)
  })

  it('returns correct default extensions', () => {
    const ct = parseContentTypes(FULL_CONTENT_TYPES)
    const extensions = ct.defaults.map((d) => d.extension)
    expect(extensions).toContain('rels')
    expect(extensions).toContain('xml')
  })

  it('returns correct override partNames', () => {
    const ct = parseContentTypes(FULL_CONTENT_TYPES)
    const partNames = ct.overrides.map((o) => o.partName)
    expect(partNames).toContain('/word/document.xml')
    expect(partNames).toContain('/word/styles.xml')
    expect(partNames).toContain('/word/settings.xml')
  })

  it('returns correct contentType for a default', () => {
    const ct = parseContentTypes(FULL_CONTENT_TYPES)
    const relsDefault = ct.defaults.find((d) => d.extension === 'rels')
    expect(relsDefault?.contentType).toBe(
      'application/vnd.openxmlformats-package.relationships+xml',
    )
  })

  it('handles single Default entry (not wrapped in array)', () => {
    const ct = parseContentTypes(SINGLE_DEFAULT_CONTENT_TYPES)
    expect(ct.defaults).toHaveLength(1)
    expect(ct.defaults[0].extension).toBe('rels')
    expect(ct.overrides).toHaveLength(1)
  })

  it('returns empty arrays for empty <Types> element', () => {
    const ct = parseContentTypes(EMPTY_CONTENT_TYPES)
    expect(ct.defaults).toHaveLength(0)
    expect(ct.overrides).toHaveLength(0)
  })
})
