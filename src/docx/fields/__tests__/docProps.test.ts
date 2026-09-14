import { describe, expect, it } from 'vitest'

import { parseCoreProps } from '../docProps'

describe('parseCoreProps', () => {
  it('extracts dc:creator and dc:title', () => {
    const xml =
      '<?xml version="1.0"?><cp:coreProperties xmlns:dc="x" xmlns:cp="y">'
      + '<dc:title>Quarterly Report</dc:title>'
      + '<dc:creator>A. Author</dc:creator>'
      + '</cp:coreProperties>'

    expect(parseCoreProps(xml)).toEqual({ author: 'A. Author', title: 'Quarterly Report' })
  })

  it('returns an empty object when the elements are absent', () => {
    expect(parseCoreProps('<cp:coreProperties/>')).toEqual({})
  })

  it('returns an empty object for undefined input', () => {
    expect(parseCoreProps(undefined)).toEqual({})
  })

  it('decodes XML entities in the extracted text', () => {
    const xml = '<dc:creator>Smith &amp; Jones</dc:creator>'
    expect(parseCoreProps(xml)).toEqual({ author: 'Smith & Jones' })
  })
})
