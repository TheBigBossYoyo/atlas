import { describe, expect, it } from 'vitest'

import { parseSettings } from '../../parser/settings'
import { writeSettingsXml } from '../settingsWriter'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function settingsXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NS}>${body}</w:settings>`
}

describe('writeSettingsXml', () => {
  it('inserts a bare trackChanges element when enabling on a document that lacked one', () => {
    const original = settingsXml('<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>')
    const result = writeSettingsXml(original, true)

    expect(result).toContain('<w:trackChanges/>')
    // Everything else about the document is preserved untouched.
    expect(result).toContain('<w:zoom w:percent="100"/>')
    expect(result).toContain('<w:defaultTabStop w:val="720"/>')
    expect(parseSettings(result as string)).toEqual({ trackChanges: true })
  })

  // ECMA-376 §17.15.1's CT_Settings sequence places w:trackChanges after
  // w:zoom but before w:defaultTabStop — inserting it unconditionally as the
  // very first child (as this module used to) would put it BEFORE w:zoom,
  // an out-of-schema-order settings.xml a strict OOXML consumer could reject
  // or flag for repair.
  it('inserts trackChanges after a known-preceding element instead of always as the first child', () => {
    const original = settingsXml('<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>')
    const result = writeSettingsXml(original, true) as string

    const zoomEnd = result.indexOf('<w:zoom w:percent="100"/>') + '<w:zoom w:percent="100"/>'.length
    const trackChangesStart = result.indexOf('<w:trackChanges/>')
    const defaultTabStopStart = result.indexOf('<w:defaultTabStop')

    expect(trackChangesStart).toBe(zoomEnd)
    expect(trackChangesStart).toBeLessThan(defaultTabStopStart)
  })

  it('inserts trackChanges after the LAST known-preceding element when several are present', () => {
    const original = settingsXml('<w:view w:val="print"/><w:zoom w:percent="100"/><w:mailMerge/>')
    const result = writeSettingsXml(original, true) as string

    const mailMergeEnd = result.indexOf('<w:mailMerge/>') + '<w:mailMerge/>'.length
    expect(result.indexOf('<w:trackChanges/>')).toBe(mailMergeEnd)
  })

  it('still inserts as the first child when the document has none of the known-preceding elements', () => {
    const original = settingsXml('<w:defaultTabStop w:val="720"/>')
    const result = writeSettingsXml(original, true) as string

    expect(result.indexOf('<w:trackChanges/>')).toBeLessThan(result.indexOf('<w:defaultTabStop'))
  })

  it('is a no-op when disabling on a document that never had the element', () => {
    const original = settingsXml('<w:zoom w:percent="100"/>')
    expect(writeSettingsXml(original, false)).toBe(original)
  })

  it('removes an existing bare trackChanges element when disabling', () => {
    const original = settingsXml('<w:zoom w:percent="100"/><w:trackChanges/><w:defaultTabStop w:val="720"/>')
    const result = writeSettingsXml(original, false)

    expect(result).not.toContain('trackChanges')
    expect(result).toContain('<w:zoom w:percent="100"/>')
    expect(result).toContain('<w:defaultTabStop w:val="720"/>')
    expect(parseSettings(result as string)).toEqual({ trackChanges: false })
  })

  it('removes a trackChanges element with w:val="false" when disabling (idempotent)', () => {
    const original = settingsXml('<w:trackChanges w:val="false"/>')
    const result = writeSettingsXml(original, false)

    expect(result).not.toContain('trackChanges')
  })

  it('normalizes an existing w:val="false" element to bare presence when enabling', () => {
    const original = settingsXml('<w:trackChanges w:val="false"/>')
    const result = writeSettingsXml(original, true)

    expect(parseSettings(result as string)).toEqual({ trackChanges: true })
  })

  it('leaves an already-on bare element untouched when enabling again (idempotent)', () => {
    const original = settingsXml('<w:trackChanges/>')
    expect(writeSettingsXml(original, true)).toBe(original)
  })

  it('synthesizes a minimal settings.xml when enabling and no original part exists', () => {
    const result = writeSettingsXml(undefined, true)

    expect(result).toBeDefined()
    expect(parseSettings(result as string)).toEqual({ trackChanges: true })
  })

  it('returns undefined when disabling and no original part exists (nothing to write)', () => {
    expect(writeSettingsXml(undefined, false)).toBeUndefined()
  })
})
