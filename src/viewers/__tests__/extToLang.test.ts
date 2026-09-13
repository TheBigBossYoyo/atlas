import { describe, it, expect } from 'vitest'
import { extToLang, getLangForExt } from '../extToLang'

describe('extToLang', () => {
  it('should cover at least 40 extensions', () => {
    expect(Object.keys(extToLang).length).toBeGreaterThanOrEqual(40)
  })

  it('should return correct mapping for standard extensions', () => {
    expect(getLangForExt('.ts')).toBe('typescript')
    expect(getLangForExt('tsx')).toBe('tsx')
    expect(getLangForExt('PY')).toBe('python')
    expect(getLangForExt('.rs')).toBe('rust')
  })

  it('should return undefined for unknown extensions', () => {
    expect(getLangForExt('.unknown')).toBeUndefined()
  })
})
