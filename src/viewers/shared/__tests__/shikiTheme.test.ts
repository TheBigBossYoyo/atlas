import { describe, expect, it } from 'vitest'
import type { Theme as ThemeId } from '../../../types'
import { getShikiThemeForAppTheme } from '../shikiTheme'

describe('getShikiThemeForAppTheme', () => {
  it.each([
    ['light', 'github-light'],
    ['dark', 'github-dark'],
    ['sepia', 'solarized-light'],
    ['nord', 'nord'],
    ['dracula', 'dracula'],
  ] as const)('maps %s to %s', (themeId, expected) => {
    expect(getShikiThemeForAppTheme(themeId)).toBe(expected)
  })

  it('throws for invalid theme ids', () => {
    expect(() => getShikiThemeForAppTheme('invalid' as ThemeId)).toThrow()
  })
})
