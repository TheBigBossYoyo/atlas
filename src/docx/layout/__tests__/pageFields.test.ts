import { describe, expect, it } from 'vitest'

import { resolvePageField, type PageFieldContext } from '../pageFields'

const CTX: PageFieldContext = {
  pageNumber: 3,
  totalPages: 12,
  sectionPageNumber: 1,
  sectionTotalPages: 4,
}

describe('resolvePageField', () => {
  it('resolves PAGE to the current page number', () => {
    expect(resolvePageField('PAGE', CTX)).toBe('3')
  })

  it('resolves NUMPAGES to the total page count', () => {
    expect(resolvePageField('NUMPAGES', CTX)).toBe('12')
  })

  it('resolves SECTIONPAGES to the owning section\'s page count', () => {
    expect(resolvePageField('SECTIONPAGES', CTX)).toBe('4')
  })

  it('falls back to totalPages for SECTIONPAGES when section page count is unknown', () => {
    expect(resolvePageField('SECTIONPAGES', { pageNumber: 3, totalPages: 12 })).toBe('12')
  })

  it('is case-insensitive on the keyword', () => {
    expect(resolvePageField('page', CTX)).toBe('3')
  })

  it('tolerates leading/trailing whitespace in the instruction', () => {
    expect(resolvePageField('  PAGE  ', CTX)).toBe('3')
  })

  it('returns undefined for an instruction this resolver does not handle', () => {
    expect(resolvePageField('DATE', CTX)).toBeUndefined()
    expect(resolvePageField('', CTX)).toBeUndefined()
  })

  it('formats with an upper-roman numeric switch', () => {
    expect(resolvePageField('PAGE \\* ROMAN', CTX)).toBe('III')
  })

  it('formats with a lower-roman numeric switch', () => {
    expect(resolvePageField('PAGE \\* roman', CTX)).toBe('iii')
  })

  it('formats with an upper-alphabetic numeric switch', () => {
    expect(resolvePageField('PAGE \\* ALPHABETIC', CTX)).toBe('C')
  })

  it('formats with a lower-alphabetic numeric switch', () => {
    expect(resolvePageField('PAGE \\* alphabetic', CTX)).toBe('c')
  })

  it('ignores an unrecognized switch (e.g. MERGEFORMAT) and falls back to decimal', () => {
    expect(resolvePageField('NUMPAGES \\* MERGEFORMAT', CTX)).toBe('12')
  })
})
