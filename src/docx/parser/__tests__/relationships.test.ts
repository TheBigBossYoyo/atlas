/**
 * Tests for src/docx/parser/relationships.ts
 */

import { describe, it, expect } from 'vitest'
import { parseRelationships } from '../relationships'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com"
    TargetMode="External"/>
</Relationships>`

const EMPTY_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseRelationships', () => {
  it('parses a single relationship (package-level _rels/.rels)', () => {
    const rels = parseRelationships(PACKAGE_RELS)
    expect(rels).toHaveLength(1)
    expect(rels[0]).toMatchObject({
      id: 'rId1',
      target: 'word/document.xml',
    })
  })

  it('parses multiple relationships (word/_rels/document.xml.rels)', () => {
    const rels = parseRelationships(DOCUMENT_RELS)
    expect(rels).toHaveLength(3)
  })

  it('returns correct ids', () => {
    const rels = parseRelationships(DOCUMENT_RELS)
    const ids = rels.map((r) => r.id)
    expect(ids).toEqual(['rId1', 'rId2', 'rId3'])
  })

  it('returns correct targets', () => {
    const rels = parseRelationships(DOCUMENT_RELS)
    expect(rels[0].target).toBe('styles.xml')
    expect(rels[1].target).toBe('settings.xml')
    expect(rels[2].target).toBe('https://example.com')
  })

  it('maps TargetMode="External" correctly', () => {
    const rels = parseRelationships(DOCUMENT_RELS)
    expect(rels[2].targetMode).toBe('External')
  })

  it('omits targetMode when not present', () => {
    const rels = parseRelationships(DOCUMENT_RELS)
    expect(rels[0].targetMode).toBeUndefined()
  })

  it('returns empty array for empty <Relationships> element', () => {
    const rels = parseRelationships(EMPTY_RELS)
    expect(rels).toHaveLength(0)
  })
})
