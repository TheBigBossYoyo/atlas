/**
 * Atlas — relsWriter unit tests (Wave D.3)
 */

import { describe, expect, it } from 'vitest'
import {
  addRelationship,
  allocateRelationshipId,
  writeRelationshipsXml,
} from '../relsWriter'
import type { Relationship } from '../../parser/relationships'

// ---------------------------------------------------------------------------
// writeRelationshipsXml
// ---------------------------------------------------------------------------

describe('writeRelationshipsXml', () => {
  it('empty rels → XML decl + Relationships root with xmlns', () => {
    const xml = writeRelationshipsXml([])
    expect(xml).toContain('<?xml version="1.0"')
    expect(xml).toContain('Relationships')
    expect(xml).toContain('http://schemas.openxmlformats.org/package/2006/relationships')
  })

  it('single internal rel is present in output', () => {
    const rels: Relationship[] = [
      {
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
        target: 'word/document.xml',
      },
    ]
    const xml = writeRelationshipsXml(rels)
    expect(xml).toContain('rId1')
    expect(xml).toContain('word/document.xml')
    expect(xml).not.toContain('TargetMode')
  })

  it('external rel includes TargetMode="External"', () => {
    const rels: Relationship[] = [
      {
        id: 'rId2',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        target: 'https://example.com',
        targetMode: 'External',
      },
    ]
    const xml = writeRelationshipsXml(rels)
    expect(xml).toContain('TargetMode')
    expect(xml).toContain('External')
    expect(xml).toContain('https://example.com')
  })
})

// ---------------------------------------------------------------------------
// allocateRelationshipId
// ---------------------------------------------------------------------------

describe('allocateRelationshipId', () => {
  it('returns rId1 from empty collection', () => {
    expect(allocateRelationshipId([])).toBe('rId1')
  })

  it('returns rIdN+1 from existing sequential ids', () => {
    const rels: Relationship[] = [
      { id: 'rId1', type: 't', target: 'a' },
      { id: 'rId2', type: 't', target: 'b' },
      { id: 'rId3', type: 't', target: 'c' },
    ]
    expect(allocateRelationshipId(rels)).toBe('rId4')
  })

  it('handles non-sequential ids (picks max+1)', () => {
    const rels: Relationship[] = [
      { id: 'rId1', type: 't', target: 'a' },
      { id: 'rId5', type: 't', target: 'b' },
      { id: 'rId3', type: 't', target: 'c' },
    ]
    expect(allocateRelationshipId(rels)).toBe('rId6')
  })
})

// ---------------------------------------------------------------------------
// addRelationship
// ---------------------------------------------------------------------------

describe('addRelationship', () => {
  it('returns a new array (immutability)', () => {
    const original: ReadonlyArray<Relationship> = []
    const { rels } = addRelationship(original, 'type', 'target')
    expect(rels).not.toBe(original)
  })

  it('assigns the next id to the added entry', () => {
    const existing: ReadonlyArray<Relationship> = [
      { id: 'rId1', type: 't', target: 'a' },
    ]
    const { id, rels } = addRelationship(existing, 'someType', 'someTarget')
    expect(id).toBe('rId2')
    expect(rels[rels.length - 1]?.id).toBe('rId2')
  })

  it('sets targetMode when mode is External', () => {
    const { rels } = addRelationship([], 't', 'https://x.com', 'External')
    expect(rels[0]?.targetMode).toBe('External')
  })

  it('does not set targetMode when mode is omitted', () => {
    const { rels } = addRelationship([], 't', 'word/doc.xml')
    expect(rels[0]?.targetMode).toBeUndefined()
  })

  it('does not mutate original array', () => {
    const original: Relationship[] = [{ id: 'rId1', type: 't', target: 'a' }]
    addRelationship(original, 't2', 'b')
    expect(original).toHaveLength(1)
  })
})
