import { describe, expect, it } from 'vitest'

import { DocxParseError } from '../unzip'
import { MAX_XML_PART_LENGTH } from '../xmlSizeGuard'
import { parseSettings } from '../settings'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function settingsXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NS}>${body}</w:settings>`
}

describe('parseSettings', () => {
  it('returns trackChanges: false when the element is absent', () => {
    expect(parseSettings(settingsXml('<w:zoom w:percent="100"/>'))).toEqual({ trackChanges: false })
  })

  it('returns trackChanges: true for a bare self-closing element', () => {
    expect(parseSettings(settingsXml('<w:trackChanges/>'))).toEqual({ trackChanges: true })
  })

  it('returns trackChanges: true for an element with no w:val attribute, paired-tag form', () => {
    expect(parseSettings(settingsXml('<w:trackChanges></w:trackChanges>'))).toEqual({ trackChanges: true })
  })

  it('returns trackChanges: false when w:val="false"', () => {
    expect(parseSettings(settingsXml('<w:trackChanges w:val="false"/>'))).toEqual({ trackChanges: false })
  })

  it('returns trackChanges: false when w:val="0"', () => {
    expect(parseSettings(settingsXml('<w:trackChanges w:val="0"/>'))).toEqual({ trackChanges: false })
  })

  it('returns trackChanges: true when w:val="true"', () => {
    expect(parseSettings(settingsXml('<w:trackChanges w:val="true"/>'))).toEqual({ trackChanges: true })
  })

  it('throws DocxParseError for malformed XML', () => {
    expect(() => parseSettings('<<<not xml')).toThrow(DocxParseError)
  })

  it('rejects a part over the XML size guard limit', () => {
    const oversized = `${'a'.repeat(MAX_XML_PART_LENGTH + 1)}`
    expect(() => parseSettings(oversized)).toThrow(DocxParseError)
  })
})
