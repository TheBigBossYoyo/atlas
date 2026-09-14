import { describe, it, expect } from 'vitest'
import { extToLang, getLangForExt } from '../extToLang'
import { EXTENSION_MANIFEST } from '../../formats/extensionManifest'

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

  // T6/DAT-14 — every code-routed extension in the canonical manifest must
  // resolve to a highlighting language; previously 22 of 61 silently fell
  // back to unhighlighted plain text.
  it.each(
    EXTENSION_MANIFEST.filter((entry) => entry.format === 'code').map((entry) => entry.ext),
  )('resolves a language for every code-routed extension: .%s', (ext) => {
    expect(getLangForExt(ext)).toBeTruthy()
  })

  it('previously-uncovered extensions now resolve (DAT-14 regression)', () => {
    const previouslyMissing = [
      'mjs', 'cjs', 'pyw', 'kts', 'cc', 'exs', 'hs', 'ml',
      'bash', 'zsh', 'fish', 'bat', 'cmd', 'jsonc', 'sass', 'less',
      'graphql', 'proto', 'dockerfile', 'cfg', 'conf', 'htm',
    ]
    for (const ext of previouslyMissing) {
      expect(getLangForExt(ext)).toBeTruthy()
    }
  })
})
